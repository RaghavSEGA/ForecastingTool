"""SEGA forecast consolidator.

Two modes:
  • Update an existing master: matches each submission to its roster row and writes the
    figures into your file, leaving notes, emails, formulas and other tabs untouched.
  • Start a fresh master: builds a new file in the same layout from the submissions alone.
Both produce a separate review report (changes, exceptions, comments, pivot data).

Run:  streamlit run app.py
"""
import copy
import hashlib
import re
from datetime import date, datetime
from types import SimpleNamespace
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from extractor import (ANNUAL, ERROR, INFO, QUARTERS, SEVERITY_ORDER, WARNING, YEAR,
                       NotATemplate, category_key, check_batch, parse_file)
from fresh import TemplateError, build_template, tidy_name, write_fresh
from html_input import parse_html, parse_pasted
from layout import master_sheets, read_roster
from matcher import auto_pick, suggest
from review import build_review
from updater import (NEW, Assignment, UpdateError, apply, check_against_master, find_repairs,
                     plan_changes)
from auth import require_login

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
NEW_LABEL = "➕ Add as a new licensee row"

st.set_page_config(page_title=f"SEGA {YEAR} forecast consolidator", page_icon="📊",
                   layout="wide")
require_login()   # @segaamerica.com email + SES one-time code; stops here until signed in


# ------------------------------------------------------------------ caching --
@st.cache_data(show_spinner=False, max_entries=500)
def parse_cached(name: str, data: bytes):
    return parse_file(name, data)


@st.cache_data(show_spinner=False, max_entries=500)
def parse_html_cached(name: str, data: bytes):
    return parse_html(name, data)


@st.cache_data(show_spinner=False, max_entries=200)
def parse_pasted_cached(name: str, text: str):
    return parse_pasted(name, text)


@st.cache_resource(show_spinner=False, max_entries=3)
def open_master(data: bytes):
    return load_workbook(BytesIO(data)), load_workbook(BytesIO(data), data_only=True)


@st.cache_data(show_spinner=False, max_entries=3)
def master_info(data: bytes):
    wf, wv = open_master(data)
    info = {}
    for name, lay in master_sheets(wf).items():
        roster = read_roster(wv[name], lay)
        info[name] = (lay, roster, find_repairs(wf[name], lay, roster))
    return info


def to_date(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    if isinstance(v, datetime):
        return v.date()
    return v if isinstance(v, date) else None


def digest(*parts) -> str:
    h = hashlib.md5()
    for p in parts:
        h.update(p if isinstance(p, bytes) else str(p).encode())
    return h.hexdigest()[:12]


def category_label(category: str | None) -> str:
    """'sleepwear, underwear, socks' -> 'Sleepwear'."""
    first = re.split(r"[,/;]| and ", category or "")[0].strip()
    return first[:1].upper() + first[1:]


def drop_missing_name_error(s):
    s.issues = [i for i in s.issues if not i.message.startswith("No company name")]


_MONTHS = (r"(jan(uary)?|feb(ruary)?|mar(ch)?|apr(il)?|may|june?|july?|aug(ust)?|"
           r"sep(t(ember)?)?|oct(ober)?|nov(ember)?|dec(ember)?)(?=\d|\b)")
_NOISE = re.compile(r"\b(20\d\d|sega|forecast|template|final|budget|due|submitted|copy|updated?|"
                    r"revised|v\d+|q[1-4]|fy\d{2,4})\b", re.I)


def name_from_filename(filename: str) -> str:
    """Best guess at a licensee name from a file name, e.g. '2027 Sega Forecast
    Template- Pan Oceanic.xlsx' -> 'Pan Oceanic'. Returns '' when there's nothing
    name-like left once dates, copy numbers and template words are removed."""
    stem = Path(filename).stem.replace("_", " ")
    stem = re.sub(r"\[.*?\]|\(.*?\)", " ", stem)                        # [22], (1)
    stem = re.sub(r"\b\d{1,2}[.\-/ ]\d{1,2}[.\-/ ]\d{2,4}\b", " ", stem)  # 9.11.26
    stem = re.sub(rf"(?i)\b{_MONTHS}\.?\s*(\d{{1,2}}\b)?(,?\s*20\d\d)?", " ", stem)  # Sep 18 2026
    stem = _NOISE.sub(" ", stem)
    stem = re.sub(r"\b\d+\b", " ", stem)
    stem = re.sub(r"\s*[-–—]+\s*", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip(" -_.,&")
    return stem if len(re.sub(r"[^A-Za-z]", "", stem)) >= 2 else ""


def form_details(s) -> str:
    """What's on an unnamed form, to help recognise which licensee sent it."""
    parts = [f"Completed by {s.completed_by}" if s.completed_by else "",
             s.category or "",
             f"dated {s.completed_date:%b %d, %Y}" if isinstance(s.completed_date, date)
             else (f"dated {s.completed_date}" if s.completed_date else "")]
    known = " · ".join(p for p in parts if p) or "The form has no contact details either."
    quarters = " · ".join(f"{q} ${s.quarter_total(q):,.0f}" for q in QUARTERS)
    return f"{known}  \n{quarters} · {YEAR} total ${s.total:,.0f}"


# --------------------------------------------------------------------- mode --
st.title(f"SEGA {YEAR} forecast consolidator")
mode = st.radio("What would you like to do?",
                ["Update an existing master", "Start a fresh master"], horizontal=True)
update = mode.startswith("Update")

if update:
    st.subheader("Current master")
    master_file = st.file_uploader("Master workbook", type=["xlsx", "xlsm"], key="master",
                                   label_visibility="collapsed")
    if not master_file:
        st.info("Upload the master you're working in, e.g. 2027 Calendar Year Forecast.xlsx.")
        st.stop()
    master_bytes = master_file.getvalue()
    try:
        sheets = master_info(master_bytes)
    except Exception as exc:
        st.error(f"Couldn't open **{master_file.name}**: {exc}")
        st.stop()
    if not sheets:
        st.error(f"No sheet in **{master_file.name}** has the master layout: a Licensee column "
                 f"with Q1–Q4 {YEAR} and {YEAR} TOTAL headers.")
        st.stop()
    sheet = (st.selectbox("Sheet to update", list(sheets)) if len(sheets) > 1
             else next(iter(sheets)))
    layout, roster, repairs = sheets[sheet]
    by_row = {r.row: r for r in roster}
    st.caption(f"**{sheet}** has {len(roster)} licensees, "
               f"{sum(1 for r in roster if r.total)} with figures so far.")
else:
    master_bytes = b""
    st.subheader("New master")
    c1, c2, c3 = st.columns([2, 2, 1.4], vertical_alignment="bottom")
    sheet_title = c1.text_input("Sheet name", f"Forecast 1 - {date.today():%B}")
    forecast_label = c2.text_input("Label for the forecast snapshot column",
                                   f"Forecast 1 {date.today().strftime('%b').upper()}")
    tidy = c3.toggle("Tidy licensee names", value=True,
                     help="Drops Inc/LLC/Ltd, 'The' and 'dba' prefixes, and un-capitalises "
                          "SHOUTED names, like the names in the master. You can still edit each one.")
    template_bytes = None
    with st.expander("Styling"):
        st.caption("Fresh masters copy the look of the 2027 Calendar Year Forecast: fonts, "
                   "colours, borders, column widths and number formats. To copy a different "
                   "master's look instead, upload it here. Only its formatting is used.")
        style_file = st.file_uploader("Master to copy the styling from", type=["xlsx", "xlsm"],
                                      key="style_master")
        if style_file:
            try:
                template_bytes = build_template(style_file.getvalue())
                st.success(f"Using the styling from **{style_file.name}**.")
            except Exception as exc:  # TemplateError or an unreadable file
                st.error(f"Couldn't use **{style_file.name}** for styling: {exc} "
                         "The standard styling will be used.")

# -------------------------------------------------------------- submissions --
st.subheader("Licensee submissions")
uploads = st.file_uploader("Forecast files", type=["xlsx", "xlsm", "html", "htm"],
                           accept_multiple_files=True, key="subs", label_visibility="collapsed")
pasted = st.session_state.setdefault("pasted", [])   # [(name, text)]


def add_paste():
    text = (st.session_state.get("paste_text") or "").strip()
    if not text:
        st.session_state["paste_msg"] = "Paste a table into the box first."
        return
    name = (st.session_state.get("paste_name") or "").strip() or f"Pasted table {len(pasted) + 1}"
    label = f"{name} (pasted)"
    if any(n == label for n, _ in pasted):
        label = f"{name} {len(pasted) + 1} (pasted)"
    pasted.append((label, text))
    st.session_state["paste_text"] = st.session_state["paste_name"] = ""
    st.session_state["paste_msg"] = ""


def remove_paste(i):
    pasted.pop(i)


with st.expander(f"Paste a table instead ({len(pasted)} added)" if pasted
                 else "Paste a table from an email instead", expanded=bool(pasted)):
    st.caption("Select the forecast table in the email or web page, copy it, and paste it here. "
               "It's read the same way as an uploaded file.")
    st.text_area("Pasted table", key="paste_text", height=130, label_visibility="collapsed",
                 placeholder="Paste the table here")
    c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
    c1.text_input("Name for this submission", key="paste_name",
                  placeholder="e.g. Thermos, from Sept 18 email")
    c2.button("Add pasted table", on_click=add_paste, width="stretch")
    if st.session_state.get("paste_msg"):
        st.warning(st.session_state["paste_msg"])
    for i, (label, _) in enumerate(pasted):
        c1, c2 = st.columns([5, 1], vertical_alignment="center")
        c1.markdown(f"📋 {label}")
        c2.button("Remove", key=f"rm_paste_{i}_{digest(label)}", on_click=remove_paste, args=(i,))

if not uploads and not pasted:
    st.info("Drop the licensee forecast files here: Excel templates, or HTML files with the "
            "forecast table. You can select them all at once.")
    st.stop()

parsed, rejected = [], []
for f in uploads or []:
    if f.name.startswith("~$"):
        continue
    try:
        if f.name.lower().endswith((".html", ".htm")):
            parsed.append(parse_html_cached(f.name, f.getvalue()))
        else:
            parsed.append(parse_cached(f.name, f.getvalue()))
    except NotATemplate as exc:
        rejected.append((f.name, str(exc)))
    except Exception as exc:
        rejected.append((f.name, f"Unexpected error while reading: {exc}"))
for label, text in pasted:
    try:
        parsed.append(parse_pasted_cached(label, text))
    except NotATemplate as exc:
        rejected.append((label, str(exc)))
    except Exception as exc:
        rejected.append((label, f"Unexpected error while reading: {exc}"))
for name, why in rejected:
    st.error(f"**{name}** was skipped. {why}")
if not parsed:
    st.stop()

# -------------------------------------------------- defaults + what needs attention --
preview = copy.deepcopy(parsed)
check_batch(preview)

labels, label_to_row, picks = {}, {}, {}
if update:
    labels = {r.row: f"{r.name} (row {r.row})" for r in roster}
    label_to_row = {v: k for k, v in labels.items()}
    for s in preview:
        cands = suggest(s, roster)
        pick = auto_pick(cands)
        picks[s.file] = (pick, cands)
        if pick:
            check_against_master(s, by_row[pick.row])

display = {}
for s in preview:
    display[s.file] = (tidy_name(s.company) if not update and tidy else s.company) or ""
if not update:
    # Same company, different divisions (e.g. three "Centric Brands" forms): add the category.
    seen = {}
    for s in preview:
        seen.setdefault(display[s.file].lower(), []).append(s)
    for group in seen.values():
        cats = {category_key(s.category) for s in group}
        if len(group) > 1 and len(cats) == len(group) and all(cats):
            for s in group:
                display[s.file] = f"{display[s.file]} ({category_label(s.category)})"


def attention(s):
    """(rank, reason): 0 = must fix, 1 = worth a look, 2 = fine. Uses parse-time state so
    rows don't jump around while you edit them."""
    errors = [i.message for i in s.issues if i.severity == ERROR]
    warnings = [i.message for i in s.issues if i.severity == WARNING]
    if update and not picks[s.file][0]:
        errors.insert(0, "Choose which roster row to write this to.")
    items = errors or warnings
    if not items:
        return 2, "✓"
    more = len(errors) + len(warnings) - 1
    text = items[0] if len(items[0]) <= 90 else items[0][:88].rstrip() + "…"
    return (0 if errors else 1), f"{'⛔' if errors else '⚠️'} {text}" + (f" (+{more} more)" if more else "")


status = {s.file: attention(s) for s in preview}


def default_row(s):
    d = s.completed_date
    row = {"include": True, "licensee": display[s.file],
           "received": d if isinstance(d, date) and d <= date.today() else date.today()}
    if update:
        pick, _ = picks[s.file]
        row["match"] = labels[pick.row] if pick else None
    else:
        row["notes"] = ""
    return row


def evidence(s):
    pick, cands = picks[s.file]
    return (pick.reason if pick else
            f"Unsure. Closest: {cands[0].name}" if cands else "No similar licensee")


EDITABLE = ["include", "licensee", "received"] + (["match"] if update else ["notes"])
key = digest(mode, *(s.file for s in parsed), master_bytes, not update and tidy)
overrides = st.session_state.setdefault(f"ovr_{key}", {})   # file -> reviewer's values
defaults = {s.file: default_row(s) for s in preview}


def current(file):
    return {**defaults[file], **overrides.get(file, {})}


# ------------------------------------------------------------------- review --
st.subheader("Review")
need = [s for s in preview if status[s.file][0] < 2]
must = sum(1 for s in need if status[s.file][0] == 0)
c1, c2 = st.columns([3, 2], vertical_alignment="center")
with c1:
    if need:
        st.markdown(f"**{len(need)} of {len(preview)} files need attention**"
                    + (f" ({must} must be fixed before downloading)" if must else "")
                    + ". They're listed first.")
    else:
        st.markdown(f"**All {len(preview)} files look good.**")
only = c2.toggle(f"Show only the {len(need)} that need attention", value=False,
                 disabled=not need, key=f"only_{key}")
st.caption("Check where each file will be written; blank ones need a choice. Received dates "
           "are only written where the master's cell is blank." if update else
           "Untick a file to leave it out. Licensee names and notes appear in the new master; "
           "files with no name on the form can be named under Download.")

# The editor shows a snapshot that's rebuilt only when the filter changes, so edits made
# in one view carry over to the other.
view = "attention" if only and need else "all"
if st.session_state.get(f"view_{key}") != view:
    st.session_state[f"view_{key}"] = view
    st.session_state[f"ver_{key}"] = st.session_state.get(f"ver_{key}", 0) + 1
ver = st.session_state[f"ver_{key}"]
snap_key = f"snap_{key}_{ver}"
if snap_key not in st.session_state:
    for k in [k for k in st.session_state if str(k).startswith(f"snap_{key}_")]:
        del st.session_state[k]
    ordered = sorted(preview, key=lambda s: (status[s.file][0], display[s.file].lower() or "~"))
    shown_rows = []
    for s in ordered:
        if view == "attention" and status[s.file][0] == 2:
            continue
        row = {c: current(s.file)[c] for c in EDITABLE}      # column order = EDITABLE
        if update:
            row["evidence"] = evidence(s)
        row.update({"check": status[s.file][1], "total": s.total, "file": s.file})
        shown_rows.append(row)
    st.session_state[snap_key] = pd.DataFrame(shown_rows)

config = {
    "include": st.column_config.CheckboxColumn("Include", width="small"),
    "licensee": st.column_config.TextColumn("Licensee (on form)" if update else "Licensee"),
    "received": st.column_config.DateColumn("Received", format="MMM D, YYYY"),
    "check": st.column_config.TextColumn("Needs attention", width="large"),
    "total": st.column_config.NumberColumn(f"{YEAR} total", format="dollar"),
    "file": st.column_config.TextColumn("File", width="medium"),
}
if update:
    config["match"] = st.column_config.SelectboxColumn(
        "Write to", options=list(labels.values()) + [NEW_LABEL], width="large")
    config["evidence"] = st.column_config.TextColumn("Why", width="medium")
    disabled = ["evidence", "check", "total", "file"]
else:
    config["notes"] = st.column_config.TextColumn("Notes")
    disabled = ["check", "total", "file"]

edited = st.data_editor(st.session_state[snap_key], key=f"review_{key}_{ver}", hide_index=True,
                        width="stretch", disabled=disabled, column_config=config)
for r in edited.to_dict("records"):
    overrides[r["file"]] = {c: r[c] for c in EDITABLE}

# ------------------------------------------------ resolve the reviewer's edits --
entries, assignments, problems = [], [], []
to_name = []  # (submission, input key, suggestion) for forms with no company name
for s in copy.deepcopy(parsed):
    row = SimpleNamespace(**current(s.file), evidence=evidence(s) if update else "")
    if not row.include:
        continue
    typed = (row.licensee or "").strip() or None
    # A name is needed in fresh mode, or when adding the file as a new roster row.
    needs_name = not update or row.match == NEW_LABEL
    if s.company is None and typed is None and needs_name:
        name_key = f"name_{key}_{digest(s.file)}"
        guess = name_from_filename(s.file.removesuffix(" (pasted)"))
        to_name.append((s, name_key, guess))
        typed = (st.session_state.get(name_key, guess) or "").strip() or None
    form_name = s.company
    if not update and s.company and typed == display[s.file]:
        s.company = typed  # the default (tidied) name, not a reviewer edit
    if typed != s.company:
        if s.company is None:
            drop_missing_name_error(s)
            s.add(INFO, f"Licensee name '{typed}' was entered during review.")
        else:
            s.add(INFO, f"Licensee name changed from '{s.company}' during review.")
        s.company = typed
    received = to_date(row.received)
    entry = {"sub": s, "received": received, "reason": "", "name": s.company or f"({s.file})",
             "form_name": form_name}

    if update:
        target = NEW if row.match == NEW_LABEL else label_to_row.get(row.match)
        if target is None:
            problems.append(f"Choose where **{s.file}** should be written.")
            entry["destination"] = "Not chosen"
        elif target == NEW:
            entry["destination"] = "New row"
            entry["reason"] = "Added as a new licensee"
        else:
            roster_name = by_row[target].name
            check_against_master(s, by_row[target])
            if not s.company:
                drop_missing_name_error(s)
                s.add(INFO, f"No name on the form; written to '{roster_name}' "
                            f"({row.evidence or 'chosen during review'}).")
            entry["name"] = roster_name
            entry["destination"] = f"Row {target} ({roster_name})"
            entry["reason"] = row.evidence or "Chosen during review"
        assignments.append(Assignment(s, target, received, s.company))
    else:
        entry["destination"] = "Fresh master"
        entry["notes"] = (row.notes or "").strip()
    entries.append(entry)

check_batch([e["sub"] for e in entries])
if update:
    used = {}
    for a in assignments:
        if isinstance(a.target, int):
            used.setdefault(a.target, []).append(a.sub.file)
    for r, files in used.items():
        if len(files) > 1:
            problems.append(f"{', '.join(files)} are all set to write to "
                            f"**{by_row[r].name}**. Keep one, or point the others elsewhere.")
errors = sum(e["sub"].count(ERROR) for e in entries)

# ------------------------------------------------------------------ changes --
approved, restore, highlight, changes = set(), False, False, []
if update and entries:
    wf, wv = open_master(master_bytes)
    changes = plan_changes(wf[sheet], wv[sheet], layout,
                           [a for a in assignments if isinstance(a.target, int)], by_row)
    fills = [c for c in changes if c.kind == "fill"]
    overwrites = [c for c in changes if c.kind == "overwrite"]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Submissions included", f"{len(entries)} of {len(parsed)}")
    c2.metric("Blank cells to fill", len(fills))
    c3.metric("Values to overwrite", len(overwrites))
    c4.metric("New licensee rows", sum(a.target == NEW for a in assignments))
    unchanged = sum(1 for a in assignments if isinstance(a.target, int)
                    and not any(c.row == a.target for c in changes))
    if unchanged:
        st.caption(f"{unchanged} submission(s) already match the master exactly.")

    if overwrites:
        st.subheader("Values that will change")
        st.caption("These cells already hold a different number. Untick any you want to keep.")
        ow = pd.DataFrame([{"apply": True, "licensee": c.licensee, "cell": c.cell,
                            "quarter": c.quarter, "ip": c.ip,
                            "before": c.before if isinstance(c.before, float) else None,
                            "after": c.after, "key": c.key} for c in overwrites])
        ow_edit = st.data_editor(
            ow, key=f"overwrites_{key}", hide_index=True, width="stretch",
            disabled=["licensee", "cell", "quarter", "ip", "before", "after"],
            column_order=["apply", "licensee", "cell", "quarter", "ip", "before", "after"],
            column_config={
                "apply": st.column_config.CheckboxColumn("Apply", width="small"),
                "licensee": "Licensee", "cell": "Cell", "quarter": "Quarter", "ip": "IP",
                "before": st.column_config.NumberColumn("In master", format="dollar"),
                "after": st.column_config.NumberColumn("In submission", format="dollar")})
        approved = set(ow_edit.loc[ow_edit["apply"], "key"])

    if repairs:
        st.subheader("Typed-over totals")
        st.warning(f"{len(repairs)} total cell(s) in the master hold a typed number instead of "
                   "their formula, so they won't update when figures change.")
        st.dataframe(pd.DataFrame([{"Licensee": r.licensee, "Cell": r.cell, "Holds": r.current,
                                    "Formula to restore": r.formula} for r in repairs]),
                     hide_index=True, width="stretch")
        restore = st.checkbox("Restore these formulas", value=True)
    highlight = st.toggle("Highlight updated cells in the master", value=False,
                          help="Fills every royalty cell the tool writes with light yellow.")
elif entries:
    c1, c2, c3 = st.columns(3)
    c1.metric("Submissions included", f"{len(entries)} of {len(parsed)}")
    c2.metric(f"{YEAR} royalty forecast", f"${sum(e['sub'].total for e in entries):,.0f}")
    c3.metric("Errors", errors)

# ---------------------------------------------------------- exceptions/notes --
st.subheader("Exceptions")
issues = sorted(({"Severity": i.severity, "Licensee": e["name"], "Cell": i.cell,
                  "Issue": i.message, "File": e["sub"].file}
                 for e in entries for i in e["sub"].issues),
                key=lambda r: (SEVERITY_ORDER[r["Severity"]], r["Licensee"].lower()))
if issues:
    show = st.segmented_control("Show", [ERROR, WARNING, INFO], default=[ERROR, WARNING],
                                selection_mode="multi", label_visibility="collapsed")
    shown = [r for r in issues if r["Severity"] in (show or [])]
    if shown:
        st.dataframe(pd.DataFrame(shown), hide_index=True, width="stretch",
                     column_config={"Issue": st.column_config.TextColumn(width="large")})
    else:
        st.caption("Nothing at the selected severity.")
elif entries:
    st.success("Every included file passed all checks.")

commented = [(e, k, v) for e in entries for k, v in e["sub"].comments.items()]
with st.expander(f"Licensee comments ({len(commented)})"):
    if not commented:
        st.caption("No licensee left comments.")
    for e, section, text in commented:
        st.markdown(f"**{e['name']}** · {section if section != ANNUAL else 'Full year'}")
        st.text(text)

# ----------------------------------------------------------------- download --
st.subheader("Download")
if not entries:
    st.warning("Tick at least one submission.")
    st.stop()

if to_name:
    unnamed = [s for s, _, _ in to_name if not s.company]
    if unnamed:
        st.error(f"{len(unnamed)} file(s) have no licensee name on the form. Enter each name "
                 "below, then press Enter.")
    else:
        st.success("All unnamed files now have a licensee name.")
    for s, name_key, guess in to_name:
        with st.container(border=True):
            left, right = st.columns([2, 3], vertical_alignment="center")
            with left:
                st.text_input(s.file, value=guess, key=name_key,
                              placeholder="Type the licensee's name")
            with right:
                st.caption(form_details(s))
                if guess and st.session_state.get(name_key, guess) == guess:
                    st.caption("Suggested from the file name. Check it's right.")
    if unnamed:
        st.stop()

for p in problems:
    st.error(p)
if problems:
    st.stop()
if errors:
    st.warning(f"{errors} error(s) remain. They're listed in the review report so you can "
               "follow up with the licensees.")

stamp = f"{date.today():%Y-%m-%d}"
if update:
    try:
        master_out, log = apply(master_bytes, layout, roster, assignments, approved,
                                repairs if restore else [], highlight)
    except UpdateError as exc:
        st.error(str(exc))
        st.stop()
    statuses = [(c, "Filled blank" if c.kind == "fill" else
                 "Overwritten" if c.key in approved else "Not applied") for c in changes]
    review_out = build_review(entries, statuses, repairs if restore else [], log)
    master_name = f"{Path(master_file.name).stem}_updated_{stamp}.xlsx"
else:
    ordered = sorted(entries, key=lambda e: e["name"].lower())
    master_out = write_fresh([e["sub"] for e in ordered],
                             {e["sub"].file: e["name"] for e in ordered},
                             {e["sub"].file: e["received"] for e in ordered},
                             {e["sub"].file: e.get("notes") for e in ordered},
                             sheet_title or "Forecast", forecast_label or None,
                             template=template_bytes)
    review_out = build_review(ordered)
    master_name = f"SEGA_{YEAR}_Forecast_Master_{stamp}.xlsx"

c1, c2 = st.columns(2)
c1.download_button("Download master", master_out, file_name=master_name, mime=XLSX,
                   type="primary", width="stretch")
c2.download_button("Download review report", review_out,
                   file_name=f"SEGA_{YEAR}_Forecast_Review_{stamp}.xlsx", mime=XLSX,
                   width="stretch")
st.caption("The review report lists every change, repair, exception and comment, plus a "
           "pivot-ready data tab. Totals in the master recalculate when you open it in Excel.")
