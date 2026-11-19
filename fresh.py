"""Fresh masters, styled by a template cut from a real master.

`build_template` takes an existing master (e.g. "2027 Calendar Year Forecast.xlsx") and keeps
only what defines its look: the header rows, one prototype licensee row in each column's
house style, the TOTAL row, column widths, zoom and the workbook theme (so theme colours
render exactly as in the original). No licensee data is kept.

`write_fresh` fills a copy of that template with submissions. A template made from your
master ships with the tool (templates/master_template.xlsx); any master can be used instead.
"""
from __future__ import annotations

import re
from collections import Counter
from copy import copy
from datetime import date
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.cell_range import CellRange

from extractor import ANNUAL, IPS, QUARTERS, Submission
from layout import NotAMaster, find_layout, master_sheets
from updater import computed_formulas

DEFAULT_TEMPLATE = Path(__file__).with_name("templates") / "master_template.xlsx"
_SUM = re.compile(r"^=SUM\(([A-Z]{1,3})\d+:([A-Z]{1,3})\d+\)$")


class TemplateError(Exception):
    pass


def _style_key(cell):
    f, fl, b, a = cell.font, cell.fill, cell.border, cell.alignment
    return (f.name, f.sz, f.b, f.i, str(f.color), fl.fill_type, str(fl.fgColor),
            tuple(getattr(b, s).style for s in ("left", "right", "top", "bottom")),
            a.horizontal, a.vertical, a.wrap_text, cell.number_format)


# --------------------------------------------------------------------------- #
def build_template(master: bytes, sheet: str | None = None) -> bytes:
    """Cut a blank template out of an existing master workbook."""
    wb = load_workbook(BytesIO(master))
    layouts = master_sheets(wb)
    if not layouts:
        raise TemplateError("No sheet in this file has the master layout.")
    name = sheet or next(iter(layouts))
    lay = layouts[name]
    ws = wb[name]
    for other in [s for s in wb.worksheets if s.title != name]:
        wb.remove(other)

    rows = [r for r in range(lay.first_row, lay.last_row + 1)
            if ws.cell(r, lay.cols["licensee"]).value not in (None, "")]
    max_col = ws.max_column
    proto, total = lay.label_row + 1, lay.label_row + 2

    # Each column's house style = its most common style across licensee rows
    # (masters edited by hand drift; the majority is the intended look).
    styles = {}
    for c in range(1, max_col + 1):
        counts = Counter(_style_key(ws.cell(r, c)) for r in rows)
        best = counts.most_common(1)[0][0]
        src = next(r for r in rows if _style_key(ws.cell(r, c)) == best)
        styles[c] = copy(ws.cell(src, c)._style)
    formulas = computed_formulas(ws, lay, rows)
    total_cells = []
    if lay.total_row:
        total_cells = [(ws.cell(lay.total_row, c).value, copy(ws.cell(lay.total_row, c)._style))
                       for c in range(1, max_col + 1)]
    total_height = ws.row_dimensions[lay.total_row].height if lay.total_row else None
    header_heights = {r: ws.row_dimensions[r].height for r in range(1, proto)}

    # Clear everything below the headers: values, styles, merges, formatting rules.
    for m in [m for m in ws.merged_cells.ranges if m.min_row >= proto]:
        ws.unmerge_cells(str(m))
    ws.delete_rows(proto, ws.max_row - proto + 1)
    for r in list(ws.row_dimensions):
        if r >= proto:
            del ws.row_dimensions[r]
    for r, h in header_heights.items():
        ws.row_dimensions[r].height = h
    ws.conditional_formatting = type(ws.conditional_formatting)()
    ws.data_validations.dataValidation = []

    for c in range(1, max_col + 1):
        cell = ws.cell(proto, c)
        cell._style = styles[c]
        if c in formulas:
            formula, origin = formulas[c]
            cell.value = Translator(formula, origin=origin).translate_formula(f"{L(c)}{proto}")
    for c, (value, style) in enumerate(total_cells, 1):
        cell = ws.cell(total, c)
        cell._style = style
        m = _SUM.match(value) if isinstance(value, str) else None
        cell.value = f"=SUM({m.group(1)}{proto}:{m.group(2)}{proto})" if m else value
    if total_height:
        ws.row_dimensions[total].height = total_height

    ws.freeze_panes = f"A{proto}"          # keep the headers visible
    wb.properties.creator = wb.properties.lastModifiedBy = "SEGA forecast consolidator"
    wb.properties.title = "SEGA forecast master template"
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
def write_fresh(subs: list[Submission], names: dict, received: dict, notes: dict,
                sheet_title: str, forecast_label: str | None = None,
                template: bytes | None = None) -> bytes:
    """Fill a copy of the template with one row per submission (keyed by file name)."""
    data = template or DEFAULT_TEMPLATE.read_bytes()
    wb = load_workbook(BytesIO(data))
    ws = wb.worksheets[0]
    try:
        lay = find_layout(ws, allow_empty=True)
    except NotAMaster as exc:
        raise TemplateError(f"The template isn't a master layout: {exc}") from exc
    proto = lay.label_row + 1
    has_total = any(str(ws.cell(proto + 1, c).value or "").upper() == "TOTAL"
                    for c in range(1, lay.cols["licensee"] + 1))
    max_col = ws.max_column
    proto_cells = [(ws.cell(proto, c).value, copy(ws.cell(proto, c)._style))
                   for c in range(1, max_col + 1)]
    total_cells = ([(ws.cell(proto + 1, c).value, copy(ws.cell(proto + 1, c)._style))
                    for c in range(1, max_col + 1)] if has_total else [])
    total_height = ws.row_dimensions[proto + 1].height

    ws.title = (sheet_title or ws.title)[:31]
    if forecast_label:
        label_col = lay.blocks[ANNUAL]["Total"] + 1
        if label_col <= max_col:
            ws.cell(lay.group_row, label_col).value = forecast_label

    # clear prototype + total rows, then write licensee rows in their place
    for r in (proto, proto + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None
    first = proto
    for i, s in enumerate(subs):
        r = first + i
        for c, (value, style) in enumerate(proto_cells, 1):
            cell = ws.cell(r, c)
            cell._style = copy(style)
            cell.value = (Translator(value, origin=f"{L(c)}{proto}").translate_formula(f"{L(c)}{r}")
                          if isinstance(value, str) and value.startswith("=") else None)
        if lay.cols["notes"]:
            ws.cell(r, lay.cols["notes"]).value = notes.get(s.file) or None
        if lay.cols["received"]:
            ws.cell(r, lay.cols["received"]).value = received.get(s.file)
        ws.cell(r, lay.cols["licensee"]).value = names.get(s.file) or s.company
        for q in QUARTERS:
            for ip in IPS:
                v = s.value(q, ip)
                ws.cell(r, lay.input_col(q, ip)).value = v if v else None

    last = first + len(subs) - 1
    if total_cells:
        t = last + 1
        for c, (value, style) in enumerate(total_cells, 1):
            cell = ws.cell(t, c)
            cell._style = copy(style)
            m = _SUM.match(value) if isinstance(value, str) else None
            cell.value = f"=SUM({m.group(1)}{first}:{m.group(2)}{last})" if m else value
        if total_height:
            ws.row_dimensions[t].height = total_height
            if t != proto + 1:
                ws.row_dimensions[proto + 1].height = None

    wb.properties.creator = wb.properties.lastModifiedBy = "SEGA forecast consolidator"
    wb.properties.title = f"SEGA forecast master ({date.today():%b %d, %Y})"
    wb.calculation.fullCalcOnLoad = True
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
_LEGAL = re.compile(r",?\s*\b(inc|incorporated|llc|l\.l\.c|ltd|limited|corp|corporation|co|"
                    r"company|plc|gmbh|lp|llp)\b\.?", re.I)


def tidy_name(name: str | None) -> str | None:
    """House style for licensee names, matching the master: 'TRENDS INTERNATIONAL, LLC'
    -> 'Trends International', 'MAH Group Inc DBA WOLFpak' -> 'WOLFpak',
    'The Bentex Group' -> 'Bentex Group'. Acronyms such as ACI or H2W are kept."""
    if not name:
        return name
    n = re.split(r"\s+d/?b/?a\s+", name, maxsplit=1, flags=re.I)[-1]   # "X dba Y" -> Y
    n = _LEGAL.sub("", n)
    n = re.sub(r"^\s*the\s+", "", n, flags=re.I)
    n = re.sub(r"\s+", " ", n).strip(" ,.-")

    def fix(word):
        letters = re.sub(r"[^A-Za-z]", "", word)
        shouting = letters.isupper() and len(letters) > 3 and re.search(r"[AEIOUY]", letters)
        return word.capitalize() if shouting else word
    tidy = " ".join(fix(w) for w in n.split(" "))
    return tidy or name