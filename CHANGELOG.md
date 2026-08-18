# Changelog

## 2026-08-18

The comparison tree moved to `output_comparison/`, cases can be named and
selected, and both layers stopped judging things that differ by construction.

### Report metadata is read, listed, and not judged

`metadata` in any `meta.json` names the cells that identify a run rather than
describe it — a report name, an id, a creation timestamp. They are read, listed
in `report.md` under **Not verified, on purpose** with both values, and left out
of the verdict.

The test is narrow: **does the value differ between two runs by construction?**
A minted id does. A creator name does not, when the same account generates every
report — a change there means the wrong account ran it, which is a finding. The
same goes for anything the figures depend on: view of risk, currency, model
version, the as-at date behind a "Data as of …" caption.

An entry is a label (`"Report ID"`, matching the cell and the value beside it,
including a fused `="Report ID: " & id`), or a cell reference (`"Cover!A3"`).
`*` covers a run of text, which is how one pattern reaches `Facility Report
Name`, `RiskPlay Report Name` and `Pro-Forma Report Name`. Either form takes a
sheet qualifier.

Exclusion reaches all three consumers — the keyed comparison, the address
sweep, and `differences.xlsx`. Each had to be told separately; a metadata cell
was briefly reported by one while another called it fine.

### `ignoreSheets` is gone from the shipped config

Cover, Contact, Table of Contents, Executive Summary, Glossary, Disclaimer,
Comments and Errata were excluded as boilerplate — 2,582 cells nobody checked.
Put back in scope they compare clean everywhere except three run-stamped cells
on Cover, now named as metadata. Images and drawings are invisible to both
layers on every sheet, so a tab full of logos costs nothing to verify. The
option still exists; nothing ships using it.

### Every difference is reported, including cascades

Numeric comparison briefly carried slack of scale × 1e-12, on the reasoning that
Excel keeps 15 significant digits so a smaller gap is rounding rather than
change. True, and the wrong trade: a cell whose stored values differ is a cell
that differs. Judging which differences matter is the reader's job, and
`tolerance` is how they say so, per column, deliberately.

Cascades are likewise reported in full. The root-cause mark stays — it says
where a problem starts — but a report that shows only causes understates the
reach of a change.

### `report.md` restructured

- a two-layer assurance block, replacing the old "cells unchecked" count
- differences grouped by column when a table has more than twelve of them
- one block per sheet, with its own count, in the recalculation and
  unchecked-cells sections
- long sections folded behind `Show`
- values printed at full stored precision, so the two sides of a difference can
  never print as the same string

### Cases can say what they are

`reportType` in a report type's `meta.json` and `label` in a case's `case.json`
now head the case in the run log and title its `report.md`:

```
✗ Validation Report · case_003 · an extra Cat Model Version row shifts the perils block down
    validation_report/case_003
    1 sheet failing
```

`reportType` inherits. **`label` does not** — one written a folder above would
describe every case beneath it identically. A label that only repeats the folder
name is treated as no label, so nothing reads `case_003 · case_003`.

### A pair can be named by folder

```
case_001/
  golden/case_1%1786955263151.xlsx
  current/case_1%1786957329031.xlsx
  results/
```

Which is what a downloader produces when it keeps the source system's name —
that name carries the download's timestamp, and renaming it to `golden.xlsx`
throws it away. Same words as the file prefixes (`golden|baseline|expected|
before`, `current|actual|new|after|report`). Exactly one spreadsheet per folder;
two stops the run rather than one being guessed at.

Re-blessing such a pair writes the new golden under the name the new report
arrived with and removes the file it replaced.

### `cases` selects what runs

```json
{ "cases": ["comparison_report/**", "!comparison_report/case_002"] }
```

In any `meta.json`, relative to that file's own folder, so the root can select
by report type and a type can select its own cases. What is left out is counted
with the results — a case set aside is not a case that passed.

### A folder meant to be a case is never skipped in silence

A malformed case folder used to be reported only when the run found no cases at
all. Beside working cases it vanished, and the total simply read one lower.
Those folders are now always reported, counted, and fail the run:

```
1 folder(s) meant to be cases could not be run:
  validation_report/case_002: a golden/ folder is here with no current/ folder beside it

5 cases, 5 failing, 1 could not be run
```

Folders that were never cases — groupings, stray directories — stay quiet unless
nothing runs at all.

### Fixes

- **Merged cells.** ExcelJS reports the master's value for every slave in a
  merge, so a merged banner was counted once per column it spanned. Slaves now
  read as blank, in the reader as well as the sweep.
- **Trailing blank rows.** Excel's used range outlives its contents: a sheet
  edited down to thirteen rows still reports thirty. Under positional matching
  every phantom row took an ordinal, so a disclaimer identical in both files
  arrived as seventeen removed rows. The trailing run is now trimmed; a blank
  row *between* populated ones is kept, since dropping it would shift everything
  under it out of step.
- **One equality rule.** `differences.xlsx` had its own, stricter than the
  comparison's, and reported cells the report said were fine. Both now call
  `equalValues`.
- **Full precision.** Report values were rounded to twelve significant digits,
  which printed the two sides of a smaller difference as the same string.

### Housekeeping

- the tree is `output_comparison/`, gitignored like the other data folders
- the per-case output folder is `results/`, renamed from `result/`
- `npm run clean` removes a `results/` beside a golden file **or** a `golden/`
  folder
- `npm run compare` defaults to `output_comparison`

### Known gaps

- **Row numbers in config are literal.** `validation_report`'s `Perils` table is
  pinned to `headerRow: 23`, which is right for two of its three cases; the
  third carries an extra `Cat Model Version` row that shifts the block to 24, so
  the key column is not found there and its differences fall through to the
  address sweep. Anchoring a header by content rather than row number would fix
  the class of problem.
- **A misspelled setting inside `sheets` is accepted and never read.** The
  unknown-key check covers the top level of a config file only.
- **`endRow` is ignored for CSV.**
- **Root cause is decided from same-row references only.** A full dependency
  graph exists in `impact.ts` and is not wired into that decision.
