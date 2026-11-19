"""Update an existing master workbook in place.

Only these cells are ever written: the quarterly royalty inputs of matched rows,
a blank Received date, typed-over total formulas (if the reviewer approves), and
new licensee rows. Notes, emails, other licensees, other tabs and all remaining
formulas and formatting are left exactly as they were.
"""
from __future__ import annotations

import re
from collections import Counter
from copy import copy
from dataclasses import dataclass
from datetime import date
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter as L

from extractor import IPS, QUARTERS, WARNING, Submission
from layout import DATE_FMT, Layout, RosterRow

NEW = "new"
_REF = re.compile(r"(?<![A-Za-z_])\$?[A-Z]{1,3}\$?\d+")
HIGHLIGHT = PatternFill("solid", fgColor="FFF2CC")


class UpdateError(Exception):
    pass


@dataclass
class Assignment:
    sub: Submission
    target: object                # roster row number, NEW, or None (skip)
    received: date | None = None
    name: str | None = None       # used for NEW rows


@dataclass
class Change:
    key: str
    licensee: str
    row: int
    cell: str
    quarter: str
    ip: str
    before: object                # number, text, or None
    after: float
    kind: str                     # "fill" (blank cell) or "overwrite"


@dataclass
class Repair:
    cell: str
    licensee: str
    current: object
    formula: str


# A submission this many times the master's current figures (or this small a fraction)
# is worth a second look: Komar's sales-dollar form was 8.6x its royalty row.
JUMP, JUMP_MIN = 4.0, 100_000


def check_against_master(sub: Submission, row: RosterRow) -> None:
    """Warn when a submission is wildly different from the figures already in its row."""
    old, new = row.total, sub.total
    if old <= 0 or new <= 0:
        return
    if new >= JUMP * old and new - old >= JUMP_MIN:
        sub.add(WARNING, f"Total ${new:,.0f} is {new / old:.1f}x the ${old:,.0f} already in the "
                         f"master for {row.name}. Check these are royalty dollars, not sales.")
    elif new * JUMP <= old and old - new >= JUMP_MIN:
        sub.add(WARNING, f"Total ${new:,.0f} is only {100 * new / old:.0f}% of the ${old:,.0f} "
                         f"already in the master for {row.name}. Check nothing is missing.")


def _is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def _number(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    return v  # text or an uncalculated formula: shown as-is


def plan_changes(ws_f, ws_v, layout: Layout, assignments: list[Assignment],
                 roster: dict) -> list[Change]:
    changes = []
    for a in assignments:
        if not isinstance(a.target, int):
            continue
        r = a.target
        for q in QUARTERS:
            for ip in IPS:
                col = layout.input_col(q, ip)
                new = round(a.sub.value(q, ip), 6)
                cached, raw = ws_v.cell(r, col).value, ws_f.cell(r, col).value
                before = _number(cached)
                if before is None and _is_formula(raw):
                    before = raw  # formula saved without a result
                if before is None:
                    if not new:
                        continue
                    kind = "fill"
                elif isinstance(before, float) and abs(before - new) < 0.005:
                    continue
                else:
                    kind = "overwrite"
                changes.append(Change(f"{r}:{col}", roster[r].name, r, f"{L(col)}{r}", q, ip,
                                      before, new, kind))
    return changes


def computed_formulas(ws_f, layout: Layout, rows: list[int]) -> dict:
    """{column: (formula, origin cell)} for columns where most licensee rows hold a
    formula that references other cells (the totals), in its most common form."""
    start = min(min(b.values()) for b in layout.blocks.values())
    found = {}
    for c in range(start, ws_f.max_column + 1):
        refs = [(r, ws_f.cell(r, c).value) for r in rows]
        refs = [(r, v) for r, v in refs if _is_formula(v) and _REF.search(v[1:])]
        if not refs or len(refs) < 0.6 * len(rows):
            continue
        origin = f"{L(c)}{rows[0]}"
        forms = Counter(Translator(v, origin=f"{L(c)}{r}").translate_formula(origin)
                        for r, v in refs)
        found[c] = (forms.most_common(1)[0][0], origin)
    return found


def find_repairs(ws_f, layout: Layout, roster: list[RosterRow]) -> list[Repair]:
    rows = [r.row for r in roster]
    names = {r.row: r.name for r in roster}
    repairs = []
    for c, (formula, origin) in computed_formulas(ws_f, layout, rows).items():
        for r in rows:
            v = ws_f.cell(r, c).value
            if not _is_formula(v):
                target = f"{L(c)}{r}"
                repairs.append(Repair(target, names[r], v,
                                      Translator(formula, origin=origin).translate_formula(target)))
    return repairs


def _retarget(v, total_row: int, move: int, last_row: int, grow: int):
    """Adjust a TOTAL-row formula: references to the TOTAL row itself move with it,
    and ranges ending at the last licensee row grow to include new rows."""
    if not _is_formula(v):
        return v

    def fix(m):
        text = m.group(0)
        row = int(re.search(r"\d+$", text).group())
        if row == total_row:
            row += move
        elif row == last_row and m.start() > 0 and v[m.start() - 1] == ":":
            row += grow
        return re.sub(r"\d+$", str(row), text)
    return _REF.sub(fix, v)


def apply(master: bytes, layout: Layout, roster: list[RosterRow],
          assignments: list[Assignment], approved: set, repairs: list[Repair],
          highlight: bool = False) -> tuple[bytes, list[str]]:
    wb = load_workbook(BytesIO(master))
    values = load_workbook(BytesIO(master), data_only=True)[layout.sheet]
    ws = wb[layout.sheet]
    log = []
    by_row = {r.row: r for r in roster}

    for rep in repairs:
        ws[rep.cell].value = rep.formula
    if repairs:
        log.append(f"Restored {len(repairs)} typed-over total formula(s).")

    # --- matched rows ------------------------------------------------------ #
    received_fmt = DATE_FMT
    if layout.cols["received"]:
        for r in by_row:
            cell = ws.cell(r, layout.cols["received"])
            if isinstance(cell.value, date):
                received_fmt = cell.number_format
                break
    changes = plan_changes(ws, values, layout, assignments, by_row)
    written = 0
    for ch in changes:
        if ch.kind == "overwrite" and ch.key not in approved:
            continue
        cell = ws[ch.cell]
        cell.value = ch.after if ch.after else None
        if highlight:
            cell.fill = HIGHLIGHT
        written += 1
    log.append(f"Wrote {written} royalty value(s) into existing rows.")

    for a in assignments:
        if isinstance(a.target, int) and a.received and layout.cols["received"]:
            cell = ws.cell(a.target, layout.cols["received"])
            if cell.value in (None, "") or (isinstance(cell.value, str) and not cell.value.strip()):
                cell.value = a.received
                cell.number_format = received_fmt

    # --- new licensee rows -------------------------------------------------- #
    news = [a for a in assignments if a.target == NEW]
    if news:
        formulas = computed_formulas(ws, layout, [r.row for r in roster])
        style_row, max_col, n = layout.last_row, ws.max_column, len(news)
        t = layout.total_row
        if t:
            below = [c.coordinate for row in ws.iter_rows(min_row=t + 1)
                     for c in row if c.value not in (None, "")]
            if below:
                raise UpdateError("There's content below the TOTAL row, so new licensee rows "
                                  "can't be added safely. Add blank rows above TOTAL in Excel, "
                                  "or skip the new licensees.")
            if any(m.min_row <= t <= m.max_row for m in ws.merged_cells.ranges):
                raise UpdateError("The TOTAL row has merged cells, so it can't be moved to "
                                  "make room for new licensees.")
            move = max(0, n - (t - layout.last_row - 1))   # use blank rows above TOTAL first
            height = ws.row_dimensions[t].height
            snapshot = [(ws.cell(t, c).value, copy(ws.cell(t, c)._style)) for c in range(1, max_col + 1)]
            for c, (value, style) in enumerate(snapshot, 1):
                if move:
                    ws.cell(t, c).value = None
                    ws.cell(t, c)._style = copy(ws.cell(style_row, c)._style)
                dst = ws.cell(t + move, c)
                dst.value = _retarget(value, t, move, layout.last_row, n)
                dst._style = style
            if move and height:
                ws.row_dimensions[t + move].height = height
            if move:
                log.append(f"Moved the TOTAL row from {t} to {t + move}.")
        for k, a in enumerate(news):
            r = layout.last_row + 1 + k
            for c in range(1, max_col + 1):
                ws.cell(r, c)._style = copy(ws.cell(style_row, c)._style)
                ws.cell(r, c).value = None
            ws.cell(r, layout.cols["licensee"], a.name or a.sub.company)
            if layout.cols["received"] and a.received:
                ws.cell(r, layout.cols["received"], a.received).number_format = received_fmt
            for q in QUARTERS:
                for ip in IPS:
                    v = a.sub.value(q, ip)
                    cell = ws.cell(r, layout.input_col(q, ip), v if v else None)
                    if highlight:
                        cell.fill = HIGHLIGHT
            for c, (formula, origin) in formulas.items():
                ws.cell(r, c).value = Translator(formula, origin=origin).translate_formula(f"{L(c)}{r}")
            log.append(f"Added '{a.name or a.sub.company}' as new row {r}.")

    wb.calculation.fullCalcOnLoad = True
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue(), log