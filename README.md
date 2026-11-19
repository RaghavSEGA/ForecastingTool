# SEGA 2027 Forecast Consolidator

Combines licensee royalty forecast templates into the master forecast workbook
(the "2027 Calendar Year Forecast" layout). Plain Python: no AI, no cloud services.

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Command line, same two modes:

```bash
python consolidate.py forecasts/                                   # fresh master
python consolidate.py forecasts/ --master "2027 Calendar Year Forecast.xlsx"   # update
```

## Two modes

**Update an existing master.** Upload your current master and the new submissions.
Each submission is matched to its roster row using the evidence below, and you confirm or
change the match in the review table. Only these cells are written:

- the quarterly royalty cells of matched rows (E:AJ in the current layout)
- Received dates that are blank in the master
- typed-over total formulas, if you approve restoring them
- new licensee rows, added above TOTAL, with the TOTAL sums extended to include them

Notes, emails, other licensees, other tabs, formatting and conditional formatting stay
as they are. Before you download, every cell where a submission would **overwrite** an
existing value is listed, and you can untick any to keep the master's value.

**Start a fresh master.** Builds a new workbook in the same layout and styling from the
submissions alone (Notes and Email are blank; you can type notes in the review table).
The fresh file can be used as the master in update mode next time.

Both modes also produce a **review report**: Submissions (where each file went and why),
Changes (every cell written, before and after), Repairs, Exceptions, Comments, and Pivot Data.

## Matching evidence (update mode)

Strongest first: identical figures to what's already in the row; the same name once
suffixes like Inc/LLC/Group are dropped; one name containing the other ("Galaken" in
"Sunscape Eyewear/Galaken"); the company name appearing in a roster email domain
("H2W" in dan@h2wonline.com → Sunstaches); shared words; similar spelling. A match is
pre-selected only when it's clearly ahead of the runner-up. Otherwise you choose.

## Checks on each submission

| Severity | Check |
|---|---|
| Error | No company name (fresh mode needs one; update mode can identify the row by its figures) |
| Error | A royalty cell holds text that isn't an amount, or a formula saved without a result |
| Warning | A total on the form doesn't match its royalty values |
| Warning | Blank category, contact or date; unreadable date; negative royalty; all-zero forecast |
| Warning | Two files with the same company name or identical figures |
| Info | Text amounts converted; content outside the form ignored |

## Notes

- Totals in the downloaded master are formulas; Excel recalculates them on open. Some
  file previewers (e.g. email attachments) show them blank until the file is opened in Excel.
- Re-saving through openpyxl stores numbers at about 15 significant digits, so a value like
  3265.4700000000003 is written as 3265.47. No figure changes by more than a fraction of a cent.
- Layout labels, IP names and the year are constants at the top of `extractor.py` and
  `layout.py`. The master is found by its labels, so moving columns is fine as long as the
  "Licensee", "Q1 2027"… "2027 TOTAL" and IP headers stay.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit app (both modes) |
| `consolidate.py` | Command line (both modes) |
| `extractor.py` | Reads and checks licensee templates |
| `layout.py` | Recognises the master layout; writes fresh masters |
| `matcher.py` | Suggests the roster row for each submission |
| `updater.py` | Plans and applies changes to an existing master |
| `review.py` | Writes the review report |
