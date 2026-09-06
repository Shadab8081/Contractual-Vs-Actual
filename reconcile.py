"""
Reconciles Actual sheets (LP/UP/5&6) ONLY.
Does NOT touch the TOTAL- LP, UP, 5&6 & PGC sheet.
User will update TOTAL sheet manually.
"""

import re
import io
import os
import json
import openpyxl

SHEET_NAMES = {"LP": "QIC_ Actual_LP", "UP": "QIC_ Actual_UP", "5&6": "QIC_ Actual_5&6"}

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


def _strict_match(a, b):
    if not a or not b:
        return False
    if a == b:
        return True
    ta, tb = set(a.split()), set(b.split())
    return ta.issubset(tb) or tb.issubset(ta)


def _get_requirement_value(val):
    """Convert requirement to int, handling formulas/strings."""
    if val is None:
        return 0
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return 0


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

        # REMOVE
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

        # PROMOTE
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
                        log.append(f"Row {promote_row}: PROMOTED {add_rec['employee_name']} from Additional to on-site {position}")

        # ADD
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
    mp = _load_manpower(manpower_bytes)
    change_log = {}

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

            log.append(f"Row {insert_row}: APPENDED {j['name']} as {pos} (new position, marked Additional)")
            insert_row += 1

        change_log[proj] = log
    return change_log


def summary_markdown(plan, actual_append_log):
    lines = ["# Manpower Reconciliation Summary — LP / UP / 5&6", ""]
    for proj in ["LP", "UP", "5&6"]:
        p = plan[proj]
        lines.append(f"## {proj}")
        lines.append(f"**Removed:** {len(p['remove_rows'])} | **Added:** {len(p['assignments'])} | **Appended:** {len(p['unplaced'])} | **Unmatched:** {len(p['unmatched'])}")
        if proj in actual_append_log and actual_append_log[proj]:
            lines.append("")
            lines.append("**New rows appended:**")
            for line in actual_append_log[proj][:15]:
                lines.append(f"- {line}")
            if len(actual_append_log[proj]) > 15:
                lines.append(f"- ... and {len(actual_append_log[proj]) - 15} more")
        lines.append("")
    lines.append("📌 **TOTAL- LP, UP, 5&6 & PGC sheet was NOT updated. Please update the numbers manually.**")
    return "\n".join(lines)


def reconcile(master_bytes, manpower_bytes):
    """Reconcile Actual sheets only. TOTAL sheet untouched."""
    wb = openpyxl.load_workbook(io.BytesIO(master_bytes), data_only=False)
    plan = build_plan(wb, manpower_bytes)
    change_log = apply_plan(wb, plan)
    actual_append_log = append_new_positions(wb, plan, manpower_bytes)

    for proj in change_log:
        change_log[proj].extend(actual_append_log.get(proj, []))

    out = io.BytesIO()
    wb.save(out)

    summary = summary_markdown(plan, actual_append_log)
    return plan, out.getvalue(), change_log, summary
