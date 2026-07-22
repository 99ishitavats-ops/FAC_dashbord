"""
Parser for the Ampyr FAC Tracker (Gantt) Excel workbook.

The workbook is a weekly Gantt sheet:
  * Columns A-H hold metadata (Site Name, FAC Date, O&M Contractor, Task,
    Sub-Task, Responsibility, Accepted By, Dates as per testing).
  * Columns 9.. are one-week-each Gantt columns. Coloured cells form the task
    bars; the fill colour encodes the task category (see CATEGORY_COLORS).
  * A "today marker" cell sits at column 89 which corresponds to 14-Jul-2026,
    letting us map every Gantt column to a real calendar date
    (7 days per column).

The parser re-detects the anchor from the pink today-marker fill each run, so
it survives small layout shifts when an updated file is uploaded.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta

import openpyxl
import pandas as pd

# Legend fill colours -> category
CATEGORY_COLORS = {
    "FFC4B5FD": "Prerequisite",
    "FF6EE7B7": "Compliance",
    "FFFDE68A": "Closure",
    "FFFED7AA": "Financial",
    "FFA5F3FC": "Technical",
    "FFFBCFE8": "Testing",
    "FFBFDBFE": "Document",
}

TODAY_MARKER_FILL = "FFFEE2E2"
DEFAULT_ANCHOR_COL = 89
DEFAULT_ANCHOR_DATE = date(2026, 7, 14)
FIRST_METADATA_ROW = 4
FIRST_GANTT_COL = 9


def _fill_hex(cell):
    fill = cell.fill
    if fill and fill.patternType:
        rgb = fill.fgColor.rgb
        if isinstance(rgb, str):
            return rgb
    return None


def _find_anchor(ws):
    for r in range(1, 4):
        for c in range(FIRST_GANTT_COL, ws.max_column + 1):
            if _fill_hex(ws.cell(row=r, column=c)) == TODAY_MARKER_FILL:
                return c, DEFAULT_ANCHOR_DATE
    return DEFAULT_ANCHOR_COL, DEFAULT_ANCHOR_DATE


def _col_to_date(col, anchor_col, anchor_date):
    return anchor_date + timedelta(weeks=(col - anchor_col))


def _parse_header(text):
    parts = [p.strip() for p in str(text).split("|")]
    name = parts[0].strip() if parts else str(text).strip()
    country, fac_raw, contractor = "", "", ""
    for p in parts[1:]:
        low = p.lower()
        if low.startswith("fac"):
            fac_raw = p.split(":", 1)[1].strip() if ":" in p else ""
        elif low.startswith("o&m") or low.startswith("o & m"):
            contractor = p.split(":", 1)[1].strip() if ":" in p else ""
        elif p and not country:
            country = p
    return name, country, contractor, fac_raw


def _parse_date_any(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.date() if isinstance(v, datetime) else v
    s = str(v).strip()
    if not s or s in {"—", "-", "TBC", "TBD"}:
        return None
    for fmt in ("%d-%b-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _status_for(d, today):
    if d is None:
        return "TBC"
    delta = (d - today).days
    if delta < 0:
        return "Overdue"
    if delta <= 90:
        return "Due soon"
    return "On track"


def parse_workbook(source, today=None):
    if today is None:
        today = date.today()
    if isinstance(today, datetime):
        today = today.date()

    wb = openpyxl.load_workbook(source, data_only=True)
    ws = wb[wb.sheetnames[0]]
    anchor_col, anchor_date = _find_anchor(ws)

    sites, milestones, tasks = [], [], []
    cur_site = cur_phase = cur_task = None

    for r in range(FIRST_METADATA_ROW, ws.max_row + 1):
        A = ws.cell(row=r, column=1).value
        D = ws.cell(row=r, column=4).value
        E = ws.cell(row=r, column=5).value
        F = ws.cell(row=r, column=6).value
        G = ws.cell(row=r, column=7).value
        H = ws.cell(row=r, column=8).value

        a_txt = str(A).strip() if A is not None else ""
        d_txt = str(D).strip() if D is not None else ""
        e_txt = str(E).strip() if E is not None else ""

        if a_txt and "|" in a_txt:
            name, country, contractor, fac_raw = _parse_header(a_txt)
            cur_site = {
                "Site": name, "Country": country, "Contractor": contractor,
                "FAC Date": _parse_date_any(fac_raw), "FYT Date": None,
            }
            sites.append(cur_site)
            cur_phase = cur_task = None
            continue

        if cur_site is None:
            continue

        if a_txt and a_txt.lower() == cur_site["Site"].lower():
            continue

        if d_txt.startswith("▸") or d_txt.startswith(">"):
            up = d_txt.upper()
            if "FIRST YEAR" in up:
                cur_phase = "First Year Testing"
            elif "FINAL ACCEPTANCE" in up:
                cur_phase = "FAC"
            else:
                cur_phase = d_txt.lstrip("▸> ").strip()
            phase_date = _parse_date_any(H)
            if cur_phase == "First Year Testing":
                cur_site["FYT Date"] = phase_date
            milestones.append({
                "Site": cur_site["Site"], "Country": cur_site["Country"],
                "Contractor": cur_site["Contractor"],
                "Milestone": cur_phase, "Date": phase_date,
            })
            cur_task = None
            continue

        if d_txt:
            cur_task = d_txt
        if not (d_txt or e_txt):
            continue

        start_col = end_col = None
        counts = {}
        for c in range(FIRST_GANTT_COL, ws.max_column + 1):
            hexc = _fill_hex(ws.cell(row=r, column=c))
            if hexc in CATEGORY_COLORS:
                if start_col is None:
                    start_col = c
                end_col = c
                cat = CATEGORY_COLORS[hexc]
                counts[cat] = counts.get(cat, 0) + 1
        category = max(counts, key=counts.get) if counts else ""
        start_date = _col_to_date(start_col, anchor_col, anchor_date) if start_col else None
        end_date = (_col_to_date(end_col, anchor_col, anchor_date) + timedelta(days=6)) if end_col else None

        tasks.append({
            "Site": cur_site["Site"], "Country": cur_site["Country"],
            "Contractor": cur_site["Contractor"], "Phase": cur_phase or "",
            "Task": cur_task or "", "Sub-Task": e_txt,
            "Responsibility": str(F).strip() if F else "",
            "Accepted By": str(G).strip() if G else "",
            "Category": category, "Start": start_date, "End": end_date,
        })

    df_sites = pd.DataFrame(sites)
    df_ms = pd.DataFrame(milestones)
    df_tasks = pd.DataFrame(tasks)

    if not df_sites.empty:
        df_sites["FAC Status"] = df_sites["FAC Date"].apply(lambda d: _status_for(d, today))
        df_sites["FYT Status"] = df_sites["FYT Date"].apply(lambda d: _status_for(d, today))
    if not df_ms.empty:
        df_ms["Status"] = df_ms["Date"].apply(lambda d: _status_for(d, today))

    return {"sites": df_sites, "milestones": df_ms, "tasks": df_tasks,
            "today": today, "anchor_col": anchor_col, "anchor_date": anchor_date}


if __name__ == "__main__":
    import sys
    data = parse_workbook(sys.argv[1], today=date(2026, 7, 21))
    print("SITES\n", data["sites"].to_string())
    print("\nMILESTONES\n", data["milestones"].to_string())
    print("\nTASKS:", len(data["tasks"]), "rows")
    print(data["tasks"].head(12).to_string())
