"""
Reconciles the QIC_ Actual_LP / QIC_ Actual_UP / QIC_ Actual_5&6 sheets of the
"Contractual vs Actual" master workbook against an uploaded "TOTAL MANPOWER"
workbook (which has a TOTAL MANPOWER sheet of active employees and a LEFT
sheet of departed employees), and keeps the TOTAL- LP , UP, 5&6 & PGC rollup
sheet in sync.

PGC is intentionally skipped everywhere. Existing Position labels, Contractual
Requirement values, formulas, merged cells and formatting are never touched --
every edit either fills an already-blank employee-slot cell, or APPENDS new
rows far below all existing content (never inserts mid-sheet, which protects
the workbook's many formulas and merged cells from silently breaking).

What happens on every run, per project (LP / UP / 5&6):
  1. REMOVE   - a named employee's Iqama is found in the LEFT sheet, or is
                active in TOTAL MANPOWER under a *different* project (moved).
                Their row is cleared and marked Vacant.
  2. PROMOTE  - if a named position's on-site slot becomes vacant and there's
                already an "Additional" employee with that same position, that
                Additional is promoted into the vacated on-site slot.
  3. ADD      - an employee active in TOTAL MANPOWER for this project who
                isn't yet listed here is placed into a vacant row whose
                Position label strictly matches their SAP POSITION or
                ASSIGNED WORK (exact match, or one is a whole-word subset).
  4. APPEND   - anyone left over (no vacant slot anywhere for their position
                - including genuinely new positions never seen before) is
                APPENDED as a new row in the Back-End/Additional area,
                Requirement=0, Status="Additional". If their position has no
                entry yet in the TOTAL-sheet mapping, a brand-new row is also
                appended to the TOTAL sheet.
  5. FLAG     - "unmatched": employees currently listed whose Iqama isn't found
                in TOTAL MANPOWER (active) or LEFT at all - left untouched.

update_total_sheet() then recomputes Onsite Qty / Additional Qty for every
mapped Position from the CURRENT state of the Actual sheets.
"""

import re
import io
import os
import json
import openpyxl

SHEET_NAMES = {"LP": "QIC_ Actual_LP", "UP": "QIC_ Actual_UP", "5&6": "QIC_ Actual_5&6"}

TOTAL_SHEET_NAME = "TOTAL- LP , UP, 5&6 & PGC"
TOTAL_SHEET_COLS = {"LP": (7, 8), "UP": (13, 14), "5&6": (19, 20)}

COL_POSITION = 1
COL_REQUIREMENT = 2
COL_ONSITE = 3
COL_VACANT = 4
COL_NAME = 5
COL_DECIPLINE = 6
COL_ID = 7
COL_IQAMA = 8
COL_STATUS = 9
COL_DEPLOY_DATE = 10
COL_SPONSOR = 11
COL_REMARKS = 12

_MAPPING_PATH = os.path.join(os.path.dirname(__file__), "position_mapping.json")
with open(_MAPPING_PATH) as _f:
    POSITION_MAPPING = json.load(_f)


def _iq(v):
    if v is None:
        return None
    s = str(v).strip()
    if s in ("", "-", "None"):
        return None
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    return s


def _norm(s):
    if s is None:
        return ""
    s = str(s).upper().strip()
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    if s.endswith("IES"):
        s = s[:-3] + "Y"
    elif s.endswith("S") and not s.endswith("SS"):
        s = s[:-1]
    return s


def _get_requirement_value(val):
    """Convert requirement value (could be number, formula string, or None) to int."""
    if val is None:
        return 0
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return 0


def _strict_match(a, b):
    if not a or not b:
        return False
    if a == b:
        return True
    ta, tb = set(a.split()), set(b.split())
    return ta.issubset(tb) or tb.issubset(ta)


def _project_match(proj, key):
    if not proj:
        return False
    p = str(proj).upper()
    if key == "LP":
        return "LOWER" in p
    if key == "UP":
        return "UPPER" in p
    if key == "5&6":
        return "5&6" in p or "5 & 6" in p
    return False


def _parse_actual_sheet(ws):
    records = []
    current_position = None
    current_req = None
    skip_next = False
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=COL_POSITION).value
        b = ws.cell(row=r, column=COL_REQUIREMENT).value
        e = ws.cell(row=r, column=COL_NAME).value
        f = ws.cell(row=r, column=COL_DECIPLINE).value
        g = ws.cell(row=r, column=COL_ID).value
        h = ws.cell(row=r, column=COL_IQAMA).value
        i = ws.cell(row=r, column=COL_STATUS).value

        if skip_next:
            skip_next = False
            continue
        if a == "Position":
            continue
        if isinstance(a, str) and a.strip().startswith("Total Team"):
            current_position = None
            skip_next = True
            continue
        if all(v is None for v in [a, b, e, f, g, h, i]):
            continue
        if a is not None:
            current_position = a
            current_req = b
        records.append(dict(row=r, position=current_position, requirement=current_req,
                             employee_name=e, decipline=f, id=g, iqama=h, status=i))
    return records


def _load_manpower(manpower_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(manpower_bytes), data_only=True)
    ws_tm = wb["TOTAL MANPOWER"]
    hdr = {c.value: i + 1 for i, c in enumerate(ws_tm[1])}

    tm_employees = []
    all_tm_iqamas = set()
    for r in range(2, ws_tm.max_row + 1):
        iqama = _iq(ws_tm.cell(row=r, column=hdr["Iqama No / PP"]).value)
        if not iqama:
            continue
        all_tm_iqamas.add(iqama)
        status = ws_tm.cell(row=r, column=hdr["STATUS"]).value
        if status and str(status).upper() != "ACTIVE":
            continue
        tm_employees.append(dict(
            iqama=iqama,
            emp_id=ws_tm.cell(row=r, column=hdr["Emp ID"]).value,
            name=ws_tm.cell(row=r, column=hdr["NAME"]).value,
            sap_position=ws_tm.cell(row=r, column=hdr["SAP POSITION"]).value,
            assigned_work=ws_tm.cell(row=r, column=hdr["ASSIGNED WORK"]).value,
            sponsor=ws_tm.cell(row=r, column=hdr["VISA SPONSOR"]).value,
            project=ws_tm.cell(row=r, column=hdr["PROJECT"]).value,
            joining_date=ws_tm.cell(row=r, column=hdr["Joining Date"]).value,
            department=ws_tm.cell(row=r, column=hdr.get("DEPARTMENT")).value,
        ))

    ws_left = wb["LEFT"]
    hdrL = {c.value: i + 1 for i, c in enumerate(ws_left[1])}
    left_iqamas = set()
    for r in range(2, ws_left.max_row + 1):
        iqama = _iq(ws_left.cell(row=r, column=hdrL["IQAMA"]).value)
        if iqama:
            left_iqamas.add(iqama)

    tm_by_project = {"LP": [], "UP": [], "5&6": []}
    for rec in tm_employees:
        for k in tm_by_project:
            if _project_match(rec["project"], k):
                tm_by_project[k].append(rec)

    tm_iqama_to_proj = {}
    for k, lst in tm_by_project.items():
        for r in lst:
            tm_iqama_to_proj[r["iqama"]] = k

    return dict(tm_employees=tm_employees, left_iqamas=left_iqamas,
                tm_by_project=tm_by_project, tm_iqama_to_proj=tm_iqama_to_proj,
                all_tm_iqamas=all_tm_iqamas)


def build_plan(master_wb, manpower_bytes):
    mp = _load_manpower(manpower_bytes)
    all_tm_iqamas = mp["all_tm_iqamas"]

    plan = {}
    for proj, sn in SHEET_NAMES.items():
        ws = master_wb[sn]
        recs = _parse_actual_sheet(ws)

        remove_rows = []
        for rec in recs:
            if not rec["employee_name"]:
                continue
            i = _iq(rec["iqama"])
            if not i:
                continue
            if i in mp["left_iqamas"]:
                remove_rows.append((rec["row"], "left"))
            else:
                ap = mp["tm_iqama_to_proj"].get(i)
                if ap and ap != proj:
                    remove_rows.append((rec["row"], f"moved_to_{ap}"))
        remove_row_nums = {r for r, _ in remove_rows}

        existing_iqamas = {_iq(r["iqama"]) for r in recs if _iq(r["iqama"])}

        vacant_rows = []
        for rec in recs:
            i = _iq(rec["iqama"])
            is_vacant_now = (not rec["employee_name"]) and (not i or rec["status"] == "Vacant")
            will_be_vacant = rec["row"] in remove_row_nums
            if is_vacant_now or will_be_vacant:
                vacant_rows.append(dict(row=rec["row"], position=rec["position"]))

        joiners = [r for r in mp["tm_by_project"][proj] if r["iqama"] not in existing_iqamas]

        assignments, unplaced, used_rows = [], [], set()
        for j in joiners:
            jn_sap, jn_aw = _norm(j["sap_position"]), _norm(j["assigned_work"])
            best = None
            for v in vacant_rows:
                if v["row"] in used_rows:
                    continue
                vn = _norm(v["position"])
                if _strict_match(jn_sap, vn) or _strict_match(jn_aw, vn):
                    best = v
                    break
            if best:
                assignments.append((j, best))
                used_rows.add(best["row"])
            else:
                unplaced.append(j)

        unmatched = []
        for rec in recs:
            if not rec["employee_name"]:
                continue
            i = _iq(rec["iqama"])
            if not i or i in mp["left_iqamas"]:
                continue
            if i not in all_tm_iqamas:
                unmatched.append(rec)

        plan[proj] = dict(remove_rows=remove_rows, assignments=assignments,
                           unplaced=unplaced, unmatched=unmatched)
    return plan


def apply_plan(master_wb, plan):
    change_log = {}
    for proj, sn in SHEET_NAMES.items():
        ws = master_wb[sn]
        p = plan[proj]
        log = []

        # REMOVE leavers and movers
        for row, reason in p["remove_rows"]:
            name_before = ws.cell(row=row, column=COL_NAME).value
            ws.cell(row=row, column=COL_NAME).value = None
            ws.cell(row=row, column=COL_ID).value = "-"
            ws.cell(row=row, column=COL_IQAMA).value = "-"
            ws.cell(row=row, column=COL_STATUS).value = "Vacant"
            ws.cell(row=row, column=COL_DEPLOY_DATE).value = None
            ws.cell(row=row, column=COL_SPONSOR).value = "-"
            ws.cell(row=row, column=COL_REMARKS).value = None
            log.append(f"Row {row}: REMOVED {name_before} ({reason})")

        # PROMOTE: Additional -> on-site when on-site slot becomes vacant
        recs = _parse_actual_sheet(ws)
        by_position_req = {}
        by_position_add = {}
        for rec in recs:
            if not rec["employee_name"]:
                continue
            i = _iq(rec["iqama"])
            if not i:
                continue
            req_val = _get_requirement_value(rec["requirement"])
            if req_val > 0:
                by_position_req.setdefault(rec["position"], []).append(rec)
            elif rec["status"] == "Additional":
                by_position_add.setdefault(rec["position"], []).append(rec)

        for position, onsite_recs in by_position_req.items():
            req = _get_requirement_value(onsite_recs[0]["requirement"])
            named = sum(1 for r in onsite_recs if r["employee_name"])
            vacant_in_block = req - named
            if vacant_in_block > 0 and position in by_position_add:
                additional_recs = by_position_add[position]
                for add_rec in additional_recs[:vacant_in_block]:
                    if add_rec["employee_name"]:
                        promote_row = add_rec["row"]
                        ws.cell(row=promote_row, column=COL_STATUS).value = "Approved"
                        log.append(
                            f"Row {promote_row}: PROMOTED {add_rec['employee_name']} "
                            f"from Additional to on-site {position}"
                        )

        # ADD to vacant slots
        for j, v in p["assignments"]:
            row = v["row"]
            jd = j.get("joining_date")
            ws.cell(row=row, column=COL_NAME).value = j["name"]
            ws.cell(row=row, column=COL_ID).value = j.get("emp_id") or "-"
            ws.cell(row=row, column=COL_IQAMA).value = j["iqama"]
            ws.cell(row=row, column=COL_STATUS).value = "Approved"
            ws.cell(row=row, column=COL_DEPLOY_DATE).value = jd if jd else None
            ws.cell(row=row, column=COL_SPONSOR).value = j.get("sponsor") or "-"
            ws.cell(row=row, column=COL_REMARKS).value = None
            log.append(f"Row {row}: ADDED {j['name']} as {v['position']} (was vacant)")

        change_log[proj] = log
    return change_log


def append_new_positions(master_wb, plan, manpower_bytes):
    """Append unplaced joiners to Back-End section. Update mapping and TOTAL sheet."""
    mp = _load_manpower(manpower_bytes)
    mapping = json.loads(json.dumps(POSITION_MAPPING))

    change_log = {}
    ws_total = master_wb[TOTAL_SHEET_NAME]

    for proj, sn in SHEET_NAMES.items():
        ws = master_wb[sn]
        p = plan[proj]
        log = []

        if not p["unplaced"]:
            change_log[proj] = log
            continue

        recs = _parse_actual_sheet(ws)
        max_row = max(r["row"] for r in recs) if recs else 3
        insert_row = max_row + 1

        for j in p["unplaced"]:
            jd = j.get("joining_date")
            pos = j.get("sap_position") or j.get("assigned_work") or "Additional"

            ws.cell(row=insert_row, column=COL_POSITION).value = pos
            ws.cell(row=insert_row, column=COL_REQUIREMENT).value = 0
            ws.cell(row=insert_row, column=COL_NAME).value = j["name"]
            ws.cell(row=insert_row, column=COL_DECIPLINE).value = j.get("department")
            ws.cell(row=insert_row, column=COL_ID).value = j.get("emp_id") or "-"
            ws.cell(row=insert_row, column=COL_IQAMA).value = j["iqama"]
            ws.cell(row=insert_row, column=COL_STATUS).value = "Additional"
            ws.cell(row=insert_row, column=COL_DEPLOY_DATE).value = jd if jd else None
            ws.cell(row=insert_row, column=COL_SPONSOR).value = j.get("sponsor") or "-"

            log.append(
                f"Row {insert_row}: APPENDED {j['name']} as {pos} "
                f"(new position, marked Additional)"
            )

            if pos not in mapping.get(proj, {}):
                total_row = find_or_create_total_row(ws_total, pos, proj)
                if proj not in mapping:
                    mapping[proj] = {}
                mapping[proj][pos] = total_row
                log.append(f"  └─ Added TOTAL sheet row {total_row} for position '{pos}'")

            insert_row += 1

        change_log[proj] = log
    return change_log, mapping


def find_or_create_total_row(ws_total, position, proj):
    """Find existing row by position name, or create new one if not found."""
    pos_norm = _norm(position)

    for r in range(4, ws_total.max_row + 1):
        fms_label = ws_total.cell(row=r, column=4).value
        if fms_label and _norm(fms_label) == pos_norm:
            return r

    new_row = ws_total.max_row + 1
    ws_total.cell(row=new_row, column=1).value = ""
    ws_total.cell(row=new_row, column=2).value = "NEW"
    ws_total.cell(row=new_row, column=3).value = "NEW POSITION"
    ws_total.cell(row=new_row, column=4).value = position

    return new_row


def update_total_sheet(master_wb, position_mapping):
    change_log = {}
    ws_total = master_wb[TOTAL_SHEET_NAME]

    for proj, sn in SHEET_NAMES.items():
        mapping = position_mapping.get(proj, {})
        if not mapping:
            change_log[proj] = []
            continue

        ws = master_wb[sn]
        recs = _parse_actual_sheet(ws)

        by_position = {}
        for rec in recs:
            if rec["position"] is None:
                continue
            by_position.setdefault(rec["position"], []).append(rec)

        totals_by_row = {}
        for pos, total_row in mapping.items():
            rows = by_position.get(pos, [])
            if not rows:
                continue
            requirement = _get_requirement_value(rows[0]["requirement"])
            named = sum(1 for r in rows if r["employee_name"])
            onsite = min(named, requirement) if requirement > 0 else 0
            additional = max(named - requirement, 0)
            acc = totals_by_row.setdefault(total_row, [0, 0])
            acc[0] += onsite
            acc[1] += additional

        onsite_col, additional_col = TOTAL_SHEET_COLS[proj]
        log = []
        for total_row, (onsite, additional) in totals_by_row.items():
            before_o = ws_total.cell(row=total_row, column=onsite_col).value
            before_a = ws_total.cell(row=total_row, column=additional_col).value
            if before_o != onsite or before_a != additional:
                ws_total.cell(row=total_row, column=onsite_col).value = onsite
                ws_total.cell(row=total_row, column=additional_col).value = additional
                fms_label = ws_total.cell(row=total_row, column=4).value
                log.append(
                    f"Row {total_row} ({fms_label}): Onsite {before_o}->{onsite}, "
                    f"Additional {before_a}->{additional}"
                )
        change_log[proj] = log
    return change_log


def summary_markdown(plan, actual_append_log, total_sheet_log):
    lines = ["# Manpower Reconciliation Summary — LP / UP / 5&6", ""]
    for proj in ["LP", "UP", "5&6"]:
        p = plan[proj]
        lines.append(f"## {proj}")
        lines.append("")
        lines.append(f"**Removed — departed or moved ({len(p['remove_rows'])}):**")
        for row, reason in p["remove_rows"]:
            lines.append(f"- Row {row} — {reason.replace('_', ' ')}")
        lines.append("")
        lines.append(f"**Filled into existing vacant slots ({len(p['assignments'])}):**")
        for j, v in p["assignments"]:
            lines.append(f"- {j['name']} → row {v['row']} ({v['position']})")
        lines.append("")
        lines.append(f"**Appended as new Additional positions ({len(p['unplaced'])}):**")
        for j in p["unplaced"]:
            lines.append(f"- {j['name']} — {j.get('sap_position')} (Iqama {j['iqama']})")
        lines.append("")
        lines.append(f"**Unmatched — listed here but not in TOTAL MANPOWER or LEFT ({len(p['unmatched'])}):**")
        for rec in p["unmatched"]:
            lines.append(f"- {rec['employee_name']} — Iqama/PP: {rec['iqama']} — Position: {rec['position']}")
        if proj in actual_append_log and actual_append_log[proj]:
            lines.append("")
            lines.append(f"**New rows appended to Actual sheet:**")
            for line in actual_append_log[proj]:
                lines.append(f"- {line}")
        if proj in total_sheet_log and total_sheet_log[proj]:
            lines.append("")
            lines.append(f"**TOTAL sheet rows updated ({len(total_sheet_log[proj])}):**")
            for line in total_sheet_log[proj][:10]:
                lines.append(f"- {line}")
            if len(total_sheet_log[proj]) > 10:
                lines.append(f"- ... and {len(total_sheet_log[proj]) - 10} more")
        lines.append("")
    return "\n".join(lines)


def reconcile(master_bytes, manpower_bytes):
    """Full reconciliation pipeline."""
    wb = openpyxl.load_workbook(io.BytesIO(master_bytes), data_only=False)
    plan = build_plan(wb, manpower_bytes)
    change_log = apply_plan(wb, plan)
    actual_append_log, updated_mapping = append_new_positions(wb, plan, manpower_bytes)
    total_sheet_log = update_total_sheet(wb, updated_mapping)

    for proj in change_log:
        change_log[proj].extend(actual_append_log.get(proj, []))
        change_log[proj].extend(total_sheet_log.get(proj, []))

    out = io.BytesIO()
    wb.save(out)

    summary = summary_markdown(plan, actual_append_log, total_sheet_log)
    return plan, out.getvalue(), change_log, summary, updated_mapping
