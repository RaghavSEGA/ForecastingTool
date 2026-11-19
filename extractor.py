"""Read SEGA 2027 royalty forecast templates. Plain Python, no AI.

The parser finds the form by its labels ("Q1 2027 Royalty", "Company Name:", ...)
rather than fixed cell addresses, so licensee edits such as typing their details
in column E instead of G are still read correctly. Anything it can't read cleanly
becomes an Issue for a person to review instead of a silent guess.
"""
from __future__ import annotations

import re
import statistics
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from io import BytesIO

from openpyxl import load_workbook

YEAR = 2027
ANNUAL = str(YEAR)
IPS = ["Movie", "Sonic Prime", "Sonic Classic", "Sonic Modern", "Other IP", "Legendary", "RGG"]
QUARTERS = ["Q1", "Q2", "Q3", "Q4"]

ERROR, WARNING, INFO = "Error", "Warning", "Info"
SEVERITY_ORDER = {ERROR: 0, WARNING: 1, INFO: 2}

HEADER_FIELDS = {"company": "company name", "category": "product category",
                 "completed_by": "completed by", "completed_date": "completed date"}
SEGA_FIELDS = {"please return to", "due date"}  # pre-filled by SEGA, not the licensee
ALL_LABELS = set(HEADER_FIELDS.values()) | SEGA_FIELDS

# Template text that is expected to be on the sheet and carries no licensee data.
STATIC_TEXT = {"sega 2027 forecast", "please list in royalty dollars",
               "other ip", "legendary", "rgg", "sega dreamcast", "sega genesis",
               "sega master system", "sega megadrive", "sega saturn", "altered beast",
               "crazy taxi", "golden axe", "jet set radio", "shinobi", "streets of rage",
               "virtua fighter", "yakuza (like a dragon)"}

DATE_FORMATS = ("%m/%d/%Y", "%m.%d.%Y", "%m-%d-%Y", "%Y-%m-%d", "%m/%d/%y",
                "%m.%d.%y", "%m-%d-%y", "%B %d, %Y", "%b %d, %Y", "%d %B %Y",
                "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%b %d %Y")
# Displayed without a year (Excel's d-mmm, e.g. "18-Sep", as it appears in HTML exports)
YEARLESS_FORMATS = ("%d-%b", "%d %b", "%b %d", "%B %d", "%m/%d")
_BLOCK_RE = re.compile(rf"^(q[1-4]) {YEAR} royalty$")
_COMMENT_PREFIX = re.compile(r"^\s*comments?\s*:?\s*", re.I)


class NotATemplate(Exception):
    """The file isn't a SEGA forecast template (or its structure was broken)."""


@dataclass
class Issue:
    severity: str
    message: str
    cell: str = ""


@dataclass
class Submission:
    file: str
    company: str | None = None
    category: str | None = None
    completed_by: str | None = None
    completed_date: date | str | None = None
    royalties: dict = field(default_factory=dict)   # {(quarter, ip): float}
    comments: dict = field(default_factory=dict)    # {"Q1".."Q4" or "2027": text}
    issues: list = field(default_factory=list)

    def add(self, severity: str, message: str, cell: str = "") -> None:
        self.issues.append(Issue(severity, message, cell))

    def value(self, quarter: str, ip: str) -> float:
        return self.royalties.get((quarter, ip), 0.0)

    def quarter_total(self, quarter: str) -> float:
        return sum(self.value(quarter, ip) for ip in IPS)

    def ip_total(self, ip: str) -> float:
        return sum(self.value(q, ip) for q in QUARTERS)

    @property
    def total(self) -> float:
        return sum(self.quarter_total(q) for q in QUARTERS)

    def count(self, severity: str) -> int:
        return sum(i.severity == severity for i in self.issues)


# --------------------------------------------------------------------------- #
def _norm(v) -> str:
    return re.sub(r"\s+", " ", str(v)).strip().lower().rstrip(":").strip()


def _blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def _money(v: float) -> str:
    return f"-${-v:,.0f}" if v < 0 else f"${v:,.0f}"


def _find_blocks(ws) -> dict:
    blocks = {}
    for row in ws.iter_rows(max_row=25):
        for c in row:
            if _blank(c.value):
                continue
            n = _norm(c.value)
            m = _BLOCK_RE.match(n)
            if m:
                blocks[m.group(1).upper()] = (c.row, c.column)
            elif n == f"{YEAR} total":
                blocks[ANNUAL] = (c.row, c.column)
    return blocks


def _find_form(wb):
    for ws in wb.worksheets:
        blocks = _find_blocks(ws)
        if len(blocks) == 5 and len({r for r, _ in blocks.values()}) == 1:
            return ws, blocks
    raise NotATemplate(f"No sheet with the Q1–Q4 {YEAR} Royalty and {YEAR} TOTAL sections.")


def _block_columns(ws, label_row: int, start_col: int, block: str) -> dict:
    cols = {}
    for off in range(10):
        n = _norm(ws.cell(label_row, start_col + off).value or "")
        for ip in IPS:
            if n == ip.lower():
                cols[ip] = start_col + off
        if n == "total":
            cols["Total"] = start_col + off
            break
    missing = [ip for ip in IPS if ip not in cols]
    if missing:
        raise NotATemplate(f"The {block} section is missing column(s): {', '.join(missing)}.")
    return cols


def _to_number(value, formula, coord: str, sub: Submission | None) -> float:
    """Convert a cell to a number. With sub=None, convert quietly (for checks)."""
    def note(sev, msg):
        if sub is not None:
            sub.add(sev, msg, coord)

    if value is None:
        if isinstance(formula, str) and formula.startswith("="):
            note(ERROR, f"Formula {formula} has no saved result, so it was read as $0. "
                        "Open the file in Excel, save it, and upload it again.")
        return 0.0
    if isinstance(value, bool):
        note(ERROR, f"TRUE/FALSE value read as $0.")
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text in ("", "-", "--"):
        return 0.0
    cleaned = re.sub(r"[,$\s]", "", text)
    negative = cleaned.startswith("(") and cleaned.endswith(")")
    try:
        number = float(cleaned.strip("()"))
    except ValueError:
        note(ERROR, f"Couldn't read '{text}' as a dollar amount, so it was read as $0.")
        return 0.0
    note(INFO, f"Text '{text}' read as {_money(-number if negative else number)}.")
    return -number if negative else number


def _parse_date(value, sub: Submission, coord: str):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 1 < value < 2958465:  # unconverted Excel serial
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    text = re.sub(r"(?i)\bsept\b\.?", "Sep", str(value).strip())
    try:
        return datetime.fromisoformat(text).date()     # "2026-09-18 00:00:00"
    except ValueError:
        pass
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    for fmt in YEARLESS_FORMATS:
        try:
            parsed = datetime.strptime(f"{text} 2000", f"{fmt} %Y").date()
        except ValueError:
            continue
        today = date.today()
        guess = parsed.replace(year=today.year)
        if guess > today + timedelta(days=31):   # e.g. "18-Dec" read in January
            guess = guess.replace(year=today.year - 1)
        sub.add(INFO, f"Completed date '{text}' has no year; read as {guess:%b %d, %Y}.", coord)
        return guess
    sub.add(WARNING, f"Completed date '{text}' isn't a recognisable date; kept as text.", coord)
    return text


def _header_value(ws, label_cell):
    """First non-empty cell to the right of a label, stopping at the next label."""
    for col in range(label_cell.column + 1, label_cell.column + 9):
        cell = ws.cell(label_cell.row, col)
        if _blank(cell.value):
            continue
        if _norm(cell.value) in ALL_LABELS:
            return None
        return cell
    return None


# --------------------------------------------------------------------------- #
def parse_file(name: str, data: bytes, rounded: bool = False) -> Submission:
    """rounded: the values were rounded for display (HTML sources), so totals on the form
    may differ from the sum of the cells by up to half a unit per cell."""
    try:
        values_wb = load_workbook(BytesIO(data), data_only=True)
        formulas_wb = load_workbook(BytesIO(data))
    except Exception as exc:
        raise NotATemplate(f"Couldn't open the file as an Excel workbook ({exc}).") from exc

    ws, blocks = _find_form(values_wb)
    fx = formulas_wb[ws.title]
    sub = Submission(file=name)
    known: set[str] = set()

    header_row = next(iter(blocks.values()))[0]
    label_row, data_row = header_row + 1, header_row + 2

    # --- licensee details ------------------------------------------------- #
    for row in ws.iter_rows(max_row=header_row - 1):
        for c in row:
            if _blank(c.value) or _norm(c.value) not in ALL_LABELS:
                continue
            known.add(c.coordinate)
            value_cell = _header_value(ws, c)
            if value_cell is None:
                continue
            known.add(value_cell.coordinate)
            field_name = next((k for k, v in HEADER_FIELDS.items() if v == _norm(c.value)), None)
            if field_name == "completed_date":
                sub.completed_date = _parse_date(value_cell.value, sub, value_cell.coordinate)
            elif field_name:
                setattr(sub, field_name, re.sub(r"\s+", " ", str(value_cell.value)).strip())

    if not sub.company:
        sub.add(ERROR, "No company name on the form. Enter it before building the master.")
    if not sub.category:
        sub.add(WARNING, "Product category is blank.")
    if not sub.completed_by:
        sub.add(WARNING, "'Completed By' is blank.")
    if sub.completed_date is None:
        sub.add(WARNING, "Completed date is blank.")
    elif isinstance(sub.completed_date, date) and sub.completed_date > date.today() + timedelta(days=1):
        sub.add(WARNING, f"Completed date {sub.completed_date:%b %d, %Y} is in the future, "
                         "probably a typo on the form.")

    # --- royalties -------------------------------------------------------- #
    columns = {b: _block_columns(ws, label_row, col, b) for b, (_, col) in blocks.items()}
    for q in QUARTERS:
        for ip in IPS:
            col = columns[q][ip]
            cell = ws.cell(data_row, col)
            amount = _to_number(cell.value, fx.cell(data_row, col).value, cell.coordinate, sub)
            if amount < 0:
                sub.add(WARNING, f"Negative royalty ({_money(amount)}) for {ip} in {q}.",
                        cell.coordinate)
            sub.royalties[(q, ip)] = amount

    # The master recalculates every total, but a mismatch on the form means the
    # licensee overrode a formula, and their intended numbers may differ.
    def check(label, col, expected, cells):
        if col is None:
            return
        cell = ws.cell(data_row, col)
        if cell.value is None and str(fx.cell(data_row, col).value or "").startswith("="):
            return  # formula saved without a result: nothing to compare against
        shown = _to_number(cell.value, None, cell.coordinate, None)
        if abs(shown - expected) > (0.5 * cells + 0.01 if rounded else 0.5):
            sub.add(WARNING, f"{label} on the form is {_money(shown)}, but the royalty values "
                             f"add up to {_money(expected)}. The master uses the royalty values.",
                    cell.coordinate)

    for q in QUARTERS:
        check(f"{q} total", columns[q].get("Total"), sub.quarter_total(q), len(IPS))
    for ip in IPS:
        check(f"{YEAR} {ip} total", columns[ANNUAL][ip], sub.ip_total(ip), len(QUARTERS))
    check(f"{YEAR} total", columns[ANNUAL].get("Total"), sub.total, len(IPS) * len(QUARTERS))

    # --- comments (anything typed in each section's comment area) ---------- #
    for block, (_, start) in blocks.items():
        end = max(columns[block].values())
        pieces = []
        for r in range(data_row + 1, data_row + 8):
            for col in range(start, end + 1):
                cell = ws.cell(r, col)
                if _blank(cell.value):
                    continue
                known.add(cell.coordinate)
                text = _COMMENT_PREFIX.sub("", str(cell.value)).strip()
                if text:
                    pieces.append(text)
        if pieces:
            sub.comments[block] = "\n".join(pieces)
    if sub.total == 0:
        sub.add(WARNING, "Every royalty value is $0. " +
                ("See the licensee's comments." if sub.comments else "There are no comments."))

    # --- anything else on the sheet --------------------------------------- #
    stray = []
    for row in ws.iter_rows():
        for c in row:
            if _blank(c.value) or c.coordinate in known or header_row <= c.row <= data_row:
                continue
            if _norm(c.value) in STATIC_TEXT:
                continue
            shown = str(c.value)
            stray.append(f"{c.coordinate}='{shown[:30]}{'…' if len(shown) > 30 else ''}'")
    if stray:
        listed = ", ".join(stray[:8]) + (f" (+{len(stray) - 8} more)" if len(stray) > 8 else "")
        sub.add(INFO, f"Content outside the form was ignored: {listed}.")

    return sub


def _company_key(name: str) -> str:
    n = re.sub(r"[^a-z0-9 ]", " ", name.lower())
    n = re.sub(r"\b(inc|llc|ltd|co|corp|corporation|company|the)\b", " ", n)
    return re.sub(r"\s+", "", n)


def category_key(category: str | None) -> str:
    """'Sleepwear, underwear, socks' -> 'sleepwear' (first listed category, normalised)."""
    first = re.split(r"[,/;&]| and ", (category or "").lower())[0]
    return re.sub(r"[^a-z0-9]", "", first)


# A submission this far above the rest is more likely sales dollars than royalties.
# Calibrated on the Sept 2026 cycle: Komar's sales-dollar form was 166x the median and
# 3.3x the next largest; the largest genuine licensee was 51x and at most 2.6x.
OUTLIER_VS_MEDIAN, OUTLIER_VS_NEXT, OUTLIER_MIN = 100, 2.0, 1_000_000


def check_batch(subs: list[Submission]) -> None:
    """Cross-file checks: likely resubmissions, divisions of the same company, and
    submissions far larger than the rest (sales entered instead of royalties)."""
    by_name, by_numbers = defaultdict(list), defaultdict(list)
    for s in subs:
        if s.company:
            by_name[_company_key(s.company)].append(s)
        if s.total:
            by_numbers[tuple(sorted(s.royalties.items()))].append(s)

    for group in by_name.values():
        for s in group if len(group) > 1 else []:
            others = [o for o in group if o is not s]
            # Different product categories = different divisions (e.g. Centric Brands
            # Beauty / Sleepwear / Socks), unless the figures are identical.
            same = [o for o in others if not (s.category and o.category)
                    or category_key(o.category) == category_key(s.category)]
            if same:
                s.add(WARNING, f"Has the same company name as {', '.join(o.file for o in same)}. "
                               "Possible resubmission: include only one to avoid double-counting.")
            if len(same) < len(others):
                divs = " and ".join(f"{o.file} ({o.category})" for o in others if o not in same)
                s.add(INFO, f"Same company as {divs}, but this form is for {s.category}, so it's "
                            "treated as a separate division.")
    for group in by_numbers.values():
        for s in group if len(group) > 1 else []:
            others = ", ".join(o.file for o in group if o is not s)
            s.add(WARNING, f"Has identical royalty figures as {others}. Possible resubmission: "
                           "include only one to avoid double-counting.")

    totals = [s for s in subs if s.total > 0]
    if len(totals) >= 5:
        for s in totals:
            rest = sorted(o.total for o in totals if o is not s)
            median = statistics.median(rest)
            if (s.total >= OUTLIER_MIN and s.total > OUTLIER_VS_MEDIAN * median
                    and s.total > OUTLIER_VS_NEXT * rest[-1]):
                s.add(WARNING, f"The ${s.total:,.0f} total is {s.total / median:,.0f}x the typical "
                               f"submission and {s.total / rest[-1]:.1f}x the next largest. Check "
                               "these are royalty dollars, not sales.")