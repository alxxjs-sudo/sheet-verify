# Changelog

## Unreleased

### Tolerance defaults to 0.001

Applied when no config names one, to both layers, and overridden by `tolerance`
anywhere in the tree — `0` for exactness. It absorbs recalculation noise, which
lands about `1.19e-7` from the stored figure on a report in the hundreds of
millions, and stays well under a cent so a change made on purpose still shows.

`0.01` was the first proposal and was measurably too wide: it swallowed a real
`34.45 → 34.46`, and the suite's own "a difference that matters is still caught,
however small" failed against it.

Four existing tests now ask for `tolerance: 0` explicitly. They were written
when exactness was the default, and each is about something else — how a wall of
differences is presented, whether the ledger and the comparer agree — so they
say what they need rather than relying on it.

### What differs gets its own table

```
**Cells that differ**

| total | within tolerance (±0.001) | above tolerance |
| ---: | ---: | ---: |
| 894 | 870 | **24 (2.7%)** |
```

894 reads as a disaster when 870 of those moved by less than a thousandth. It is
the figure the whole report exists to deliver, so it no longer sits as one line
among the file paths, and the tolerance that drew the line is named on the
column that used it. Where a run applies several — tolerance resolves per column
— the header gives the range rather than claiming a single number.

### `differences.xlsx` leaves out within-tolerance rows

It is read as a list of things to fix, and a cell the tolerance already forgave
is not one. `--ledger all` still carries them, and `report.md` lists every one
under "Inside the tolerance you set", so the rule stays auditable.

### Tolerance reaches layer 2

A tolerance used to quiet only the keyed comparison. The address sweep knew
nothing about it, so `cells differing` — the number at the top of every report,
and the one most people read first — went on counting gaps of `1.19e-7` that
the config had already declared immaterial, and the same cells filled the
"nothing checked them" list.

The sweep now reads the per-column tolerances off the tables layer 1 compared,
translating them from column names to the addresses it works in, and takes a
blanket `*` from `defaults` for the cells no compared table covers. Both layers
therefore agree about which gaps matter.

Cells inside tolerance are counted and listed on their own — `within tolerance`
beside the differing count, and a folded section showing each cell with the size
of its gap. Hiding them would make a tolerance set too wide invisible, which is
a worse failure than the noise it was set to remove.

Layer 2 compares the text a cell displays and never sees its type, so text that
reads as a number is treated as one. Documented rather than worked around: the
answer is to set a tolerance to the size of the rounding it absorbs.

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
