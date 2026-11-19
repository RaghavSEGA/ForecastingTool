"""Read forecast templates supplied as HTML tables or pasted text.

Covers .html/.htm files (Excel's "Save as Web Page", Google Sheets' HTML download, a saved
email) and tables copied straight out of an email or web page and pasted into the app,
which arrive as HTML or tab-separated text. Each table becomes a worksheet grid (merged
header cells stay at their top-left position, as in Excel), and that grid goes through the
same parser and checks as a spreadsheet. Uses only the standard library.
"""
from __future__ import annotations

import csv
import re
from html.parser import HTMLParser
from io import BytesIO, StringIO

from openpyxl import Workbook

from extractor import INFO, NotATemplate, Submission, parse_file

_NUMBER = re.compile(r"\(?\s*-?\s*[$£€]?\s*-?\s*(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?\s*\)?")


def _int(v, default=1) -> int:
    try:
        return max(1, min(int(str(v).strip()), 50))
    except (TypeError, ValueError):
        return default


class _Tables(HTMLParser):
    """Collect every <table> as rows of [text parts, colspan, rowspan]."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables, self.stack, self.cell, self.skip = [], [], None, 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("script", "style"):
            self.skip += 1
        elif tag == "table":
            self.stack.append([])
        elif tag == "tr" and self.stack:
            self.stack[-1].append([])
        elif tag in ("td", "th") and self.stack:
            if not self.stack[-1]:
                self.stack[-1].append([])
            self.cell = [[], _int(a.get("colspan")), _int(a.get("rowspan"))]
            self.stack[-1][-1].append(self.cell)
        elif tag == "br" and self.cell is not None:
            self.cell[0].append("\n")
        elif tag in ("p", "div", "li") and self.cell is not None and self.cell[0]:
            self.cell[0].append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self.skip = max(0, self.skip - 1)
        elif tag in ("td", "th"):
            self.cell = None
        elif tag == "table" and self.stack:
            table = self.stack.pop()
            self.cell = None
            if any(table):
                self.tables.append(table)

    def handle_data(self, data):
        if self.cell is not None and not self.skip:
            self.cell[0].append(data)


def _clean(text: str) -> str:
    lines = [re.sub(r"[ \t\r\f\v\u00a0]+", " ", line).strip()
             for line in text.replace("\u00a0", " ").split("\n")]
    return "\n".join(line for line in lines if line)


def _grid(rows) -> list[list[str]]:
    """Lay out HTML rows on a grid, honouring colspan/rowspan."""
    cells, width = {}, 0
    for r, row in enumerate(rows):
        c = 0
        for parts, colspan, rowspan in row:
            while (r, c) in cells:
                c += 1
            text = _clean("".join(parts))
            for dr in range(rowspan):
                for dc in range(colspan):
                    cells[(r + dr, c + dc)] = text if (dr, dc) == (0, 0) else ""
            c += colspan
            width = max(width, c)
    height = max((r for r, _ in cells), default=-1) + 1
    return [[cells.get((r, c), "") for c in range(width)] for r in range(height)]


def _value(text: str):
    """'$1,200.50' -> 1200.5, '(300)' -> -300, '12,000' -> 12000; anything else stays text."""
    t = (text or "").strip()
    if not t:
        return None
    if _NUMBER.fullmatch(t) and re.search(r"\d", t):
        negative = t.startswith("(") and t.endswith(")") or "-" in t
        number = float(re.sub(r"[^\d.]", "", t))
        number = -number if negative else number
        return int(number) if number.is_integer() else number
    return t


def _decode(data: bytes) -> str:
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", errors="replace")
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data.decode("utf-16", errors="replace")
    head = data[:4000].decode("ascii", errors="ignore")
    m = re.search(r"charset=[\"']?([\w-]+)", head, re.I)
    for enc in ([m.group(1)] if m else []) + ["utf-8", "cp1252"]:
        try:
            return data.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", errors="replace")


def _tables(html: str) -> list[list[list[str]]]:
    parser = _Tables()
    parser.feed(html)
    parser.close()
    return [_grid(t) for t in parser.tables]


def _parse_grids(name: str, grids: list, source: str, rounded: bool = True) -> Submission:
    """Stack the grids on one sheet (a blank row between tables) and parse it."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    top = 1
    for grid in grids:
        for i, row in enumerate(grid):
            for j, text in enumerate(row):
                v = _value(text)
                if v is not None:
                    ws.cell(top + i, j + 1, v)
        top += len(grid) + 1
    buf = BytesIO()
    wb.save(buf)
    sub = parse_file(name, buf.getvalue(), rounded=rounded)
    sub.add(INFO, f"Read from {source}. HTML holds amounts as displayed, so they may be rounded "
                  "to the dollar or cent. Cell references count the table's rows and columns "
                  "from A1.")
    return sub


def parse_html(name: str, data: bytes) -> Submission:
    grids = _tables(_decode(data))
    if not grids:
        raise NotATemplate("There's no table in this HTML file.")
    return _parse_grids(name, grids, "an HTML table")


def parse_pasted(name: str, text: str) -> Submission:
    """A table copied from an email, web page or spreadsheet: HTML or tab-separated."""
    if re.search(r"<\s*table", text, re.I):
        grids = _tables(text)
        if not grids:
            raise NotATemplate("The pasted HTML has no table in it.")
        return _parse_grids(name, grids, "pasted HTML")
    rows = list(csv.reader(StringIO(text.strip("\n")), delimiter="\t"))
    if not rows or max(len(r) for r in rows) < 2:
        raise NotATemplate("That doesn't look like a table. Copy the cells themselves from the "
                           "email or spreadsheet (not a screenshot) and paste again.")
    return _parse_grids(name, [[[_clean(c) for c in r] for r in rows]], "a pasted table")