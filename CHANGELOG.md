# Changelog

Versions follow the policy in [the README](README.md#versioning): a major is
reserved for changes to configuration keys, CLI flags, exports and artefact
shapes. Detection changes ship as minors, each saying what to recheck.

## 1.3.0 — 2026-08-19

### A block with no header row is read as data, not renamed by one

A key-value block -- a label in column A, its value in column B -- has no header
row. Detection had to pick one anyway, and that cost twice over: the row picked
stopped being data, and the value column took its name from a *value*.

Where that value is the report's own name or id, it differs between any two runs
by construction. The column then pairs with nothing in the other file, and every
row of the block arrives as one column removed and another added -- thirty
findings for a fifteen-row block, with the real change buried among them. The
rows above the chosen header were not compared at all: on a pro-forma report,
`Report ID` and the report name sat outside every table, so an edit to either
was invisible to layer 1.

It went unseen here because a golden and its actual are usually the same report
with something planted in it, so the name matched. Two genuinely different runs
are what expose it.

Detection now tells the two shapes apart by asking what the painting marks out.
A header row is painted to stand out from its data -- bold white on navy across
the table, nothing beneath it. A key-value block is painted down its label
column on *every* row, so the paint describes a column, not a row, and no row
stands out. Where every row is painted alike, the block is read as data from its
first row down, with columns named `Column A`, `Column B` -- names no edit to
the report can change.

Blocks nobody painted are excluded from the rule deliberately. Every row of an
unstyled table is painted identically too -- not at all -- and concluding "no
header" there would throw away the column names of every plain grid.

`headerRow` accepts `0` to say this by hand, for a block detection reads wrong.
It is not a new concept: `headerRow` has always named the row *above* the data,
so `0` is simply the row above row 1 -- which is what a block at the top of a
sheet needs, and what no value could express before.

Measured over the same tree of 34 real cases: **1,092 tables before and after,
14,747 columns before and after, no case changing verdict** -- and **103 more
rows compared**, on 58 tables. Eleven of those rows carried a planted edit that
had never been compared.

Table ranges are more accurate as a side effect. The range said what rectangle a
table covered by starting at `headerRow`, which claimed a row that named nothing
when the header row was blank. It now starts at the header row only when that
row named the columns, correcting **163 ranges** across the tree.

## 1.2.0 — 2026-08-19

Detection finds blocks it used to walk past — tables printed side by side, and
tables under a bold section title. A run can assert what it expects to compare,
and a report type's configuration can be generated from the pairs. Table
numbering shifts where new blocks appear, so anything pinned by number wants
rechecking with `--print-spec`.

### A section title no longer hides the table underneath it

Detection picks a block's header row by looking for the row that names the most
columns, with formatting breaking the tie — these generators paint a header row
deliberately, and nothing else on the sheet is painted that way.

A bold section title sitting alone in column A beat that test. It was the only
painted row in its block, so it won the search; then the block was discarded for
naming a single column, and every row under the title went unseen by layer 1
with nothing said about it. On one report that lost an entire exchange-rate
block — four label/value pairs, one of which had changed.

A row naming fewer than two columns is now passed over during the search, which
is the same rule already applied to the row that wins it. Across a tree of 33
real cases this brought **55 previously invisible blocks into detection** and
44 more tables into layer 1, with no change to any table that was already found
and no case changing verdict.

Table numbers below a newly found block shift by one — `Table 2` means "the
second block detection finds", so anything pinned by number on an affected sheet
needs rechecking with `--print-spec`. None of the sheets configured in the tree
this was measured on were affected.

### The README is a starting point again, not the manual

It had grown to 1,181 lines, most of it reference that only matters once you
are already running. Split into a document per question, with the README down
to 321 lines: what it is, install, quick start, the five things worth doing
once, the one behaviour to know before trusting a result, how it works, and a
table saying where everything else lives.

| | |
| --- | --- |
| `docs/cli.md` | the command: flags, pair naming, case selection, labels, generating a config |
| `docs/configuration.md` | keys, sheets, tables, tolerance, metadata, the options reference |
| `docs/detection-tuning.md` | working out which correction a report needs |
| `docs/reading-a-report.md` | the case folder, each artefact, reading the verdict |
| `docs/api.md` | matchers, the direct API, invariants, extending it |

No prose was dropped -- every line of the original lands in exactly one of
these, checked mechanically -- and every internal link was repointed and
verified to resolve. `docs/` now ships with the package, so those links work
for a consumer too.

### Adding a table to a single-table sheet no longer replaces it

A sheet holding one table carries its header row and key at the sheet level,
with no `tables` block -- that is what makes reports read "Ledger" rather than
"Ledger · Table 1". Declaring a table on such a sheet *replaced* it: one entry
written to check a title block also stopped the sheet's real table being
compared, and nothing said so. The sheet's own table is now kept, filed under
the sheet's name, which is the name reports already give it:

```
| Ledger         | `A1:C4` | 3 columns | 3 rows | key |
| Ledger · Bands | `A5:C7` | 3 columns | 2 rows | key |
```

Only fires when a layer above declares `tables` and the one below has none, so
a sheet-level `keyColumns` written on its own -- the common correction, and the
one in every example -- behaves exactly as before.

### Overriding and adding tables: docs/detection-tuning.md

Two new sections, both written from behaviour that was checked rather than
assumed: which fields survive a partial override (all of them -- merging is per
table and per field), what a table name that matches nothing does (it adds a
phantom at `A1:B1` and fails the run, rather than doing nothing), and the three
things that decide where an added table lands -- `headerRow`, its bound, and
`columns` when it has a neighbour.

### `--write-meta`: a report type's meta.json, from the pairs

Starting a report type meant typing out a config by hand and finding out which
parts were needed from the failures. Now:

    sheet-verify --write-meta output_comparison/natural_cat_srq

It reads the pairs under that folder and writes a `meta.json` holding what it
can show evidence for, and comments for the rest:

- **`reportType`**, from the folder name, with a note that the spelling is
  yours to fix.
- **`metadata`** — run-identity labels found in the files, matched against a
  deliberately short vocabulary and kept only where the value was **observed to
  take more than one value** across the files scanned. A label found in a wide
  header row is passed over: "Program ID" as the 28th column heading of a table
  is not run identity, and listing it would drop that column out of the
  comparison. A report type that spells its own name into the label gets the
  wildcard form, `*Report Name`, which is what covers Facility / Pro-Forma /
  RiskPlay in one entry. Labels found but unvarying are listed as candidates,
  not written.
- **`defaults`** — `requireCachedValues: false` when the files really do carry
  formulas with no stored result, `fillKeyDown: true` when a sheet really does
  write a group heading once and leave it blank beneath. Both quote the
  evidence.
- **the unkeyed tables**, largest first, as a note — the work that is left,
  visible without being pre-decided.

What it deliberately does **not** write is per-sheet header rows, end rows and
keys. Those are re-detected from the files on every run, which is what lets a
report change shape without breaking the config; frozen into a file they go
stale, and a generated entry is indistinguishable from one somebody meant.

It refuses to touch a `meta.json` that already has settings in it.

On a facility report it produced `Report ID`, `*Report Name`, `Creation Date`
and `Elapsed Processing Time` — the hand-written list plus one — and the tree
it wrote gave the same verdicts as the hand-written config.

### `expect`: a table that stops being compared now fails

Detection is remade from the files on every run. That is the point, and it
means coverage can shrink with nothing to show for it but a smaller number in
a summary line nobody was watching. On one report here an entire block went
unread by layer 1 for the life of the tool.

    { "expect": { "Report Info": ["A2:B17", "H2:J22", "A19:B24"], "Cover": 1 } }

Ranges, or a count. Checked after the comparison and reported as an integrity
error, which fails the run and names the sheet:

```
- Report Info: expected 5 table(s), compared 4 — not compared: A99:B120
- Comments: expected 3 table(s) compared by name and key, found 1
```

It is an assertion, never an instruction — it changes nothing about what is
compared, so an entry that is wrong stops the run and says so rather than
quietly comparing the wrong thing, which is what a pinned `headerRow` does when
a report shifts.

`sheet-verify --write-expect` records it from a run, into each case's
`case.json`, beside whatever is already there. Deliberately its own step rather
than part of `--bless`: blessing accepts a change to the *output*, and
accepting a change to what is *checked* is a separate decision.

### Every report says which rectangle each table covered

The question "what are we actually verifying" had no answer in the output. A
table's spec does not hold one either: with no `endRow` it runs to the bottom
of the sheet, and with no `columns` it spans the whole width. So the range is
now read off the model — what was *read*, not what was asked for — and reported
in three places.

`report.md` gains a **What was verified** section listing every table layer 1
read, with its range in each file, its width, its height and whether rows were
matched by key or by position. Each changed table also carries its own range
under the heading, because a finding at `B25` reads differently depending on
whether the table starts at row 2 or row 24:

```markdown
### Currency Info · Table 2

_`A18:B23` — 2 columns × 5 rows, rows matched by position_
```

`compared.xlsx` gains a banner on row 1 of every table's worksheet saying the
same thing — bold, on a filled band the width of the table, over a rule. The
column headers move to row 2 and the frozen pane and autofilter move with them.
The band is light where the header under it is dark, so the two read as two
things rather than one heavy block. `diff.json` gains `range: { base, next }`
on every compared sheet outcome.

The two files are named separately whenever they disagree. A table that grew,
moved or lost its bottom rows between runs explains a wall of differences that
would otherwise read as edits.

### Tables printed side by side are found, and can be bounded by column

A sheet was split into tables by blank rows alone. That cannot separate a
definitions table in `H:J` from the key-value block in `A:B` beside it: both
start on row 1, and the taller one keeps every row of the block non-blank. Read
as one table the two header rows fuse, the shorter table's rows read as blank,
and a key named in one of them can be found in the other. One report had four
tables on such a sheet and two were detected.

Detection now cuts a sheet by rows and by columns in turn, until neither moves —
cutting by rows can expose a column gap that spanned the old region, and the
other way round. A cut needs **two or more** blank columns: one is a spacer, and
cutting on one fragments 186 blocks across a tree of 33 real cases, most of them
wrongly. Two cuts 67, and those are genuinely two tables.

New `columns` option on a sheet or table, written as a range of letters:

```json
{ "tables": { "Definitions": { "columns": "H:J", "keyColumns": ["Return Year"] } } }
```

Detection sets it when a sheet holds tables side by side, and only then — a
stacked sheet gets no bound, since freezing a width that grows would help
nobody. It is what stops a header row reaching across into a neighbour: without
it the reader takes the whole width, finds the neighbour's columns hold data,
and folds them into this table.

Measured over the same tree: **37 fewer tables matched by row position**, and
118 cells moved from layer 1 to layer 2 — caption bands like `Label A / Label B
/ % Change` that were being read as one-row tables and are decoration, not data.
Both layers still cover every cell; no case changed verdict. Table numbers shift
on any sheet where this applies, so anything pinned by number wants rechecking
with `--print-spec`.

### Tuning detection: docs/detection-tuning.md

How to read a run for the tables detection got wrong, recipes for the two shapes
it reliably struggles with — a key-value block with no header row, and a data
table under title blocks — and which parts of the configuration are positional
rather than named, so it is clear which entries survive a report changing shape
and which do not.

## 1.1.0 — 2026-08-18

Comparison gained a tolerance — `0.001` by default, applied in both layers —
with the cells it forgives counted and listed rather than quietly dropped.

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

## 1.0.0 — 2026-08-18

First release, covering everything through `ce721d0`. The changelog starts
here, so the entries below are the last day of that work and everything
earlier is in the git history.

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
