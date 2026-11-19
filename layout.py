"""The master forecast layout (as in "2027 Calendar Year Forecast.xlsx").

One row per licensee:  Notes | Received | Licensee | Email Address |
Q1..Q4 2027 (Movie, Sonic Prime, Sonic Classic, Sonic Modern, Other IP, Legendary, RGG, Total) |
2027 TOTAL (same 8) | <forecast label> | FY27 | Latest forecast- FY27, then a TOTAL row.

`find_layout` recognises that structure by its labels in an existing file. Fresh masters
are written by fresh.py from a template cut from a real master.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from extractor import ANNUAL, IPS, QUARTERS, YEAR, _norm

BLOCKS = QUARTERS + [ANNUAL]
MONEY = '"$"#,##0_);[Red]\\("$"#,##0\\)'
DATE_FMT = "d-mmm"
_GROUP_RE = re.compile(rf"^(q[1-4]) {YEAR}( royalty)?$")


class NotAMaster(Exception):
    """The sheet doesn't have the master forecast layout."""


@dataclass
class Layout:
    sheet: str
    group_row: int              # "Q1 2027" ... "2027 TOTAL"
    label_row: int              # Movie, Sonic Prime, ...
    first_row: int              # first licensee row
    last_row: int               # last row with a licensee name
    total_row: int | None       # the TOTAL row, if any
    cols: dict                  # notes / received / licensee / email -> column (or None)
    blocks: dict                # "Q1".."Q4", "2027" -> {ip or "Total": column}

    def input_col(self, quarter: str, ip: str) -> int:
        return self.blocks[quarter][ip]


@dataclass
class RosterRow:
    row: int
    name: str
    email: str
    received: object
    values: dict = field(default_factory=dict)   # {(quarter, ip): float}

    @property
    def total(self) -> float:
        return sum(self.values.values())


def _map_block(ws, label_row: int, start: int, block: str) -> dict:
    cols = {}
    for off in range(10):
        n = _norm(ws.cell(label_row, start + off).value or "")
        for ip in IPS:
            if n == ip.lower():
                cols[ip] = start + off
        if n == "total":
            cols["Total"] = start + off
            break
    missing = [ip for ip in IPS if ip not in cols]
    if missing:
        raise NotAMaster(f"The {block} section has no column for: {', '.join(missing)}.")
    return cols


def find_layout(ws, allow_empty: bool = False) -> Layout:
    """allow_empty: accept a sheet with no licensee rows yet (a template)."""
    heads = {}
    for row in ws.iter_rows(max_row=10):
        for c in row:
            if c.value is None:
                continue
            n = _norm(c.value)
            m = _GROUP_RE.match(n)
            if m:
                heads[m.group(1).upper()] = (c.row, c.column)
            elif n == f"{YEAR} total":
                heads[ANNUAL] = (c.row, c.column)
    if len(heads) != 5 or len({r for r, _ in heads.values()}) != 1:
        raise NotAMaster(f"No 'Q1–Q4 {YEAR}' and '{YEAR} TOTAL' headers found.")
    group_row = next(iter(heads.values()))[0]
    label_row = group_row + 1
    blocks = {b: _map_block(ws, label_row, col, b) for b, (_, col) in heads.items()}

    cols = {"notes": None, "received": None, "licensee": None, "email": None}
    first_block_col = min(col for _, col in heads.values())
    for r in (group_row, label_row):
        for c in range(1, first_block_col):
            n = _norm(ws.cell(r, c).value or "")
            if n == "notes":
                cols["notes"] = c
            elif n.startswith("received"):
                cols["received"] = c
            elif n in ("licensee", "company", "company name"):
                cols["licensee"] = c
            elif n.startswith("email"):
                cols["email"] = c
    if not cols["licensee"]:
        raise NotAMaster("No 'Licensee' column found.")

    first_row = label_row + 1
    total_row, last_row = None, label_row
    for r in range(first_row, ws.max_row + 1):
        if any(_norm(ws.cell(r, c).value or "") == "total" for c in range(1, cols["licensee"] + 1)):
            total_row = r
            break
        if ws.cell(r, cols["licensee"]).value not in (None, ""):
            last_row = r
    if last_row < first_row and not allow_empty:
        raise NotAMaster("The sheet has no licensee rows.")
    return Layout(ws.title, group_row, label_row, first_row, last_row, total_row, cols, blocks)


def master_sheets(wb) -> dict:
    """{sheet name: Layout} for every sheet that has the master layout."""
    found = {}
    for ws in wb.worksheets:
        try:
            found[ws.title] = find_layout(ws)
        except NotAMaster:
            pass
    return found


def read_roster(ws_values, layout: Layout) -> list[RosterRow]:
    roster = []
    for r in range(layout.first_row, layout.last_row + 1):
        name = ws_values.cell(r, layout.cols["licensee"]).value
        if name in (None, ""):
            continue
        email = ws_values.cell(r, layout.cols["email"]).value if layout.cols["email"] else ""
        received = ws_values.cell(r, layout.cols["received"]).value if layout.cols["received"] else None
        values = {}
        for q in QUARTERS:
            for ip in IPS:
                v = ws_values.cell(r, layout.input_col(q, ip)).value
                values[(q, ip)] = float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
        roster.append(RosterRow(r, str(name).strip(), str(email or ""), received, values))
    return roster