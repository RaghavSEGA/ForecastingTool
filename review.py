"""The review report that accompanies the master (both modes).

Tabs: Submissions · Changes (update mode) · Repairs (update mode) · Exceptions ·
Comments · Pivot Data. Kept separate from the master so the master stays clean.
"""
from __future__ import annotations

from datetime import date, datetime
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

from extractor import ANNUAL, IPS, QUARTERS, SEVERITY_ORDER, Submission

MONEY = '$#,##0;($#,##0);"-"'
NAVY = "1F3864"
FILLS = {"Error": "F8CBAD", "Warning": "FFE699", "Overwritten": "FFE699",
         "Not applied": "F8CBAD", "Filled blank": "E2EFDA"}
WRAP = Alignment(wrap_text=True, vertical="top")


def _font(**kw):
    return Font(name="Arial", size=10, **kw)


def _table(wb, title, headers, widths, rows, money_cols=(), empty="Nothing to show.",
           fill_col=None):
    ws = wb.create_sheet(title)
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(1, c, h)
        cell.font = _font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[L(c)].width = w
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            if isinstance(v, str) and v.startswith("="):
                cell.data_type = "s"  # show formulas as text, don't evaluate them
            cell.font = _font()
            cell.alignment = WRAP
            if c in money_cols and isinstance(v, (int, float)):
                cell.number_format = MONEY
            if isinstance(v, date):
                cell.number_format = "yyyy-mm-dd"
        if fill_col and FILLS.get(row[fill_col - 1]):
            ws.cell(r, fill_col).fill = PatternFill("solid", fgColor=FILLS[row[fill_col - 1]])
    if not rows:
        ws.cell(2, 1, empty).font = _font(italic=True)
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{L(len(headers))}{len(rows) + 1}"
    return ws


def build_review(entries: list[dict], changes=None, repairs=None, log=None,
                 generated: datetime | None = None) -> bytes:
    """entries: one dict per included submission with keys
    sub (Submission), name (str), destination (str), reason (str), received (date|None)."""
    generated = generated or datetime.now()
    wb = Workbook()
    wb.remove(wb.active)

    _table(wb, "Submissions",
           ["Licensee", "Written to", "Match evidence", "Received", "2027 Total", "Errors",
            "Warnings", "Source File"],
           [30, 26, 34, 12, 14, 8, 9, 44],
           [[e["name"], e["destination"], e["reason"], e["received"], e["sub"].total,
             e["sub"].count("Error"), e["sub"].count("Warning"), e["sub"].file] for e in entries],
           money_cols=(5,))
    ws = wb["Submissions"]
    note = f"Generated {generated:%b %d, %Y %H:%M}."
    if log:
        note += " " + " ".join(log)
    ws.cell(len(entries) + 3, 1, note).font = _font(italic=True, color="595959")

    if changes is not None:
        _table(wb, "Changes",
               ["Status", "Licensee", "Cell", "Quarter", "IP", "Before", "After"],
               [13, 30, 8, 8, 14, 14, 14],
               [[status, c.licensee, c.cell, c.quarter, c.ip, c.before, c.after]
                for c, status in changes],
               money_cols=(6, 7), empty="No royalty values changed.", fill_col=1)
    if repairs is not None:
        _table(wb, "Repairs", ["Licensee", "Cell", "Was", "Restored formula"],
               [30, 8, 14, 30],
               [[r.licensee, r.cell, r.current, r.formula] for r in repairs],
               money_cols=(3,), empty="No typed-over totals found.")

    items = sorted(((i, e) for e in entries for i in e["sub"].issues),
                   key=lambda t: (SEVERITY_ORDER[t[0].severity], t[1]["name"].lower()))
    _table(wb, "Exceptions", ["Severity", "Licensee", "Cell", "Issue", "Source File"],
           [10, 30, 8, 90, 44],
           [[i.severity, e["name"], i.cell, i.message, e["sub"].file] for i, e in items],
           empty="No exceptions.", fill_col=1)

    _table(wb, "Comments", ["Licensee", "Section", "Comment", "Source File"], [30, 10, 90, 44],
           [[e["name"], k if k != ANNUAL else "Full year", v, e["sub"].file]
            for e in entries for k, v in e["sub"].comments.items()],
           empty="No licensee comments.")

    _table(wb, "Pivot Data", ["Licensee", "Quarter", "IP", "Royalty", "Source File"],
           [30, 9, 15, 14, 44],
           [[e["name"], q, ip, e["sub"].value(q, ip), e["sub"].file]
            for e in entries for q in QUARTERS for ip in IPS],
           money_cols=(4,))

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
