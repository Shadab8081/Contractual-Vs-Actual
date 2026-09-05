"""
Reconciles the QIC_ Actual_LP / QIC_ Actual_UP / QIC_ Actual_5&6 sheets of the
"Contractual vs Actual" master workbook against an uploaded "TOTAL MANPOWER"
workbook (which has a TOTAL MANPOWER sheet of active employees and a LEFT
sheet of departed employees).

PGC is intentionally skipped. Only employee-slot cells (Employee Name, ID,
Iqama, Status, Deployment Date, Sponsor, Remarks) are ever edited. Position
labels, Contractual Requirement values, formulas (On-Site Availability /
Vacant Position / subtotals), merged cells and all formatting are left
completely untouched.

Three kinds of changes are made automatically:
  1. REMOVE  - a named employee's Iqama is found in the LEFT sheet, or is
               active in TOTAL MANPOWER under a *different* project (moved).
               Their row is cleared and marked Vacant.
  2. ADD     - an employee active in TOTAL MANPOWER for this project who
               isn't yet listed here is placed into a vacant row whose
               Position label strictly matches their SAP POSITION or
               ASSIGNED WORK (exact match, or one is a whole-word subset of
               the other). No fuzzy/character-similarity matching is used,
               to avoid misplacing people into a similar-looking but wrong
               role.
  3. FLAG    - anything that doesn't cleanly resolve is left untouched and
               reported instead of guessed at:
                 - "unplaced": active in TOTAL MANPOWER, no matching vacant
                   slot found (likely genuine new/overflow headcount).
                 - "unmatched": currently listed here, but their Iqama isn't
                   found in TOTAL MANPOWER (active) or LEFT at all.
"""

import re
import io
import os
import json
import openpyxl

SHEET_NAMES = {"LP": "QIC_ Actual_LP", "UP": "QIC_ Actual_UP", "5&6": "QIC_ Actual_5&6"}

TOTAL_SHEET_NAME = "TOTAL- LP , UP, 5&6 & PGC"
# (onsite_qty_col, additional_qty_col) per project, on the TOTAL sheet
TOTAL_SHEET_COLS = {"LP": (7, 8), "UP": (13, 14), "5&6": (19, 20)}

_MAPPING_PATH = os.path.join(os.path.dirname(__file__), "position_mapping.json")
with open(_MAPPING_PATH) as _f:
    POSITION_MAPPING = json.load(_f)  # {"LP": {"Position label": total_row_num, ...}, ...}

COL_POSITION = 1
COL_REQUIREMENT = 2
COL_NAME = 5
COL_DECIPLINE = 6
COL_ID = 7
COL_IQAMA = 8
COL_STATUS = 9
COL_DEPLOY_DATE = 10
COL_SPONSOR = 11
COL_REMARKS = 12


def _iq(v):
    """Normalize an Iqama/passport value to a comparable string, or None."""
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
    """Forward-fills Position/Requirement across each block of rows, skips
    header rows, 'Total Team ...' subtotal rows (and the numeric row right
    below each one), and fully blank rows."""
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
    """Read-only: computes what WOULD change. Does not modify master_wb."""
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


def update_total_sheet(master_wb):
    """
    Recomputes Onsite Qty / Additional Qty on the TOTAL- LP , UP, 5&6 & PGC
    sheet from the CURRENT state of the Actual sheets (call this AFTER
    apply_plan has made its edits), using the confirmed Position ->
    TOTAL-row mapping in position_mapping.json.

    For each Position label with a mapping entry:
      - blocks with a non-zero Contractual Requirement contribute their
        named-employee headcount to Onsite Qty (capped implicitly, since a
        block's row count equals its Requirement)
      - blocks with a ZERO Requirement (the Back-End/Additional overflow
        rows for that same position) contribute their named-employee
        headcount to Additional Qty
      - if a non-zero-Requirement block ever has more named employees than
        its Requirement (shouldn't normally happen, given fixed row counts,
        but handled defensively), the excess also counts as Additional

    Multiple raw Position labels that map to the same TOTAL-sheet row (e.g.
    "Driver" and "DRIVER" both mapping to row 96) have their headcounts
    summed into that one row. Expat Qty / Saudi Qty (the approved
    requirement) on the TOTAL sheet are never touched.

    Returns a change log: {project: [description, ...]}.
    """
    change_log = {}
    ws_total = master_wb[TOTAL_SHEET_NAME]

    for proj, sn in SHEET_NAMES.items():
        mapping = POSITION_MAPPING.get(proj, {})
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

        totals_by_row = {}  # total_row -> [onsite, additional]
        for pos, total_row in mapping.items():
            rows = by_position.get(pos, [])
            if not rows:
                continue
            requirement = rows[0]["requirement"] or 0
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


def apply_plan(master_wb, plan):
    """Mutates master_wb in place according to a plan from build_plan()."""
    change_log = {}
    for proj, sn in SHEET_NAMES.items():
        ws = master_wb[sn]
        p = plan[proj]
        log = []
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


def summary_markdown(plan, total_sheet_log=None):
    lines = ["# Manpower Reconciliation Summary — LP / UP / 5&6", ""]
    for proj in ["LP", "UP", "5&6"]:
        p = plan[proj]
        lines.append(f"## {proj}")
        lines.append("")
        lines.append(f"**Removed — departed or moved to another project ({len(p['remove_rows'])}):**")
        for row, reason in p["remove_rows"]:
            lines.append(f"- Row {row} — {reason.replace('_', ' ')}")
        lines.append("")
        lines.append(f"**Filled into existing vacant slots ({len(p['assignments'])}):**")
        for j, v in p["assignments"]:
            lines.append(f"- {j['name']} → row {v['row']} ({v['position']})")
        lines.append("")
        lines.append(f"**Needs manual placement — active, no matching vacant slot ({len(p['unplaced'])}):**")
        for j in p["unplaced"]:
            lines.append(f"- {j['name']} — {j.get('sap_position')} (Iqama {j['iqama']})")
        lines.append("")
        lines.append(f"**Unmatched — listed here but not found in TOTAL MANPOWER or LEFT ({len(p['unmatched'])}):**")
        for rec in p["unmatched"]:
            lines.append(f"- {rec['employee_name']} — Iqama/PP: {rec['iqama']} — Position: {rec['position']}")
        if total_sheet_log:
            changes = total_sheet_log.get(proj, [])
            lines.append("")
            lines.append(f"**TOTAL sheet rows updated ({len(changes)}):**")
            for line in changes:
                lines.append(f"- {line}")
        lines.append("")
    return "\n".join(lines)


def reconcile(master_bytes, manpower_bytes):
    """Convenience wrapper: returns (plan, updated_master_bytes, change_log, summary_md)."""
    wb = openpyxl.load_workbook(io.BytesIO(master_bytes), data_only=False)
    plan = build_plan(wb, manpower_bytes)
    change_log = apply_plan(wb, plan)
    total_sheet_log = update_total_sheet(wb)
    for proj in change_log:
        change_log[proj].extend(total_sheet_log.get(proj, []))
    out = io.BytesIO()
    wb.save(out)
    summary = summary_markdown(plan, total_sheet_log)
    return plan, out.getvalue(), change_log, summary
