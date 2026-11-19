"""Command-line consolidation, both modes.

Fresh master:     python consolidate.py forecasts/ -o master.xlsx
Update a master:  python consolidate.py forecasts/ --master "2027 Calendar Year Forecast.xlsx"

In update mode, only confident matches are written; anything unsure is listed and
skipped (use the app to choose). Overwrites and formula repairs are applied unless
--keep-existing / --no-repair are given. A review report is always written too.
"""
import argparse
import sys
from datetime import date
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook

from extractor import ERROR, NotATemplate, check_batch, parse_file
from html_input import parse_html
from fresh import build_template, tidy_name, write_fresh
from layout import master_sheets, read_roster
from matcher import auto_pick, suggest
from review import build_review
from updater import (Assignment, UpdateError, apply, check_against_master, find_repairs,
                     plan_changes)


def collect(inputs):
    paths = []
    for p in map(Path, inputs):
        paths += (sorted(x for x in p.iterdir() if x.suffix.lower() in (".xlsx", ".xlsm", ".html", ".htm"))
                  if p.is_dir() else [p])
    return [p for p in paths if not p.name.startswith("~$")]


def received_date(s):
    """The form's completed date, unless it's missing or in the future (a typo)."""
    d = s.completed_date
    return d if isinstance(d, date) and d <= date.today() else None


def main() -> int:
    ap = argparse.ArgumentParser(description="Consolidate SEGA forecast templates.")
    ap.add_argument("inputs", nargs="+", help="forecast .xlsx/.html files and/or folders")
    ap.add_argument("--master", help="existing master to update (omit for a fresh master)")
    ap.add_argument("--sheet", help="sheet to update (default: first with the master layout)")
    ap.add_argument("-o", "--output", help="output master path")
    ap.add_argument("--review", help="output review report path")
    ap.add_argument("--keep-existing", action="store_true",
                    help="never overwrite a non-blank cell in the master")
    ap.add_argument("--no-repair", action="store_true",
                    help="don't restore typed-over total formulas")
    ap.add_argument("--label", default=f"Forecast 1 {date.today().strftime('%b').upper()}",
                    help="fresh mode: snapshot column label")
    ap.add_argument("--style-from", help="fresh mode: copy the styling of this master instead "
                                         "of the built-in one")
    ap.add_argument("--keep-names", action="store_true",
                    help="fresh mode: use names exactly as written on the forms")
    args = ap.parse_args()
    stamp = f"{date.today():%Y-%m-%d}"

    subs = []
    for p in collect(args.inputs):
        if args.master and p.resolve() == Path(args.master).resolve():
            continue
        try:
            reader = parse_html if p.suffix.lower() in (".html", ".htm") else parse_file
            subs.append(reader(p.name, p.read_bytes()))
        except NotATemplate as exc:
            print(f"SKIPPED   {p.name}: {exc}", file=sys.stderr)
    check_batch(subs)
    entries = []

    if args.master:
        data = Path(args.master).read_bytes()
        wf, wv = load_workbook(BytesIO(data)), load_workbook(BytesIO(data), data_only=True)
        sheets = master_sheets(wf)
        if not sheets:
            sys.exit(f"{args.master} has no sheet with the master layout.")
        name = args.sheet or next(iter(sheets))
        layout = sheets[name]
        roster = read_roster(wv[name], layout)
        by_row = {r.row: r for r in roster}
        assignments, taken = [], {}
        for s in subs:
            pick = auto_pick(suggest(s, roster))
            if not pick or pick.row in taken:
                why = f"also matches {taken[pick.row]}" if pick else "no confident match"
                print(f"UNMATCHED {s.file}: {why}; skipped (use the app to choose).")
                continue
            taken[pick.row] = s.file
            check_against_master(s, by_row[pick.row])
            if not s.company:
                s.issues = [i for i in s.issues if not i.message.startswith("No company name")]
            assignments.append(Assignment(s, pick.row, received_date(s)))
            entries.append({"sub": s, "name": pick.name, "reason": pick.reason,
                            "form_name": s.company,
                            "destination": f"Row {pick.row} ({pick.name})",
                            "received": assignments[-1].received})
            print(f"{s.file[:48]:<48} -> {pick.name} (row {pick.row}; {pick.reason})")
        changes = plan_changes(wf[name], wv[name], layout, assignments, by_row)
        approved = set() if args.keep_existing else {c.key for c in changes}
        repairs = [] if args.no_repair else find_repairs(wf[name], layout, roster)
        try:
            out, log = apply(data, layout, roster, assignments, approved, repairs)
        except UpdateError as exc:
            sys.exit(str(exc))
        statuses = [(c, "Filled blank" if c.kind == "fill" else
                     "Overwritten" if c.key in approved else "Not applied") for c in changes]
        output = args.output or f"{Path(args.master).stem}_updated_{stamp}.xlsx"
        review = build_review(entries, statuses, repairs, log)
        print("\n" + "\n".join(log))
    else:
        unnamed = [s.file for s in subs if not s.company]
        if unnamed:
            print("NO NAME   " + ", ".join(unnamed) + ": listed as UNIDENTIFIED.")
        def shown(s):
            n = s.company if args.keep_names else tidy_name(s.company)
            return n or f"UNIDENTIFIED ({s.file})"
        subs.sort(key=lambda s: shown(s).lower())
        names = {s.file: shown(s) for s in subs}
        received = {s.file: received_date(s) for s in subs}
        entries = [{"sub": s, "name": names[s.file], "form_name": s.company, "reason": "",
                    "destination": "Fresh master", "received": received[s.file]} for s in subs]
        template = build_template(Path(args.style_from).read_bytes()) if args.style_from else None
        out = write_fresh(subs, names, received, {}, f"Forecast 1 - {date.today():%B}",
                          args.label, template=template)
        output = args.output or f"SEGA_2027_Forecast_Master_{stamp}.xlsx"
        review = build_review(entries)
        print(f"{len(subs)} submissions, 2027 total ${sum(s.total for s in subs):,.2f}")

    Path(output).write_bytes(out)
    review_path = args.review or f"SEGA_2027_Forecast_Review_{stamp}.xlsx"
    Path(review_path).write_bytes(review)
    errors = sum(e["sub"].count(ERROR) for e in entries)
    print(f"\nWrote {output} and {review_path} ({errors} error(s) to follow up).")
    return 0


if __name__ == "__main__":
    sys.exit(main())