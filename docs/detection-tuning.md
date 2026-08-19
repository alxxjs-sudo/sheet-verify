# Tuning detection for a report type

Detection is a guess made from one file. It is a good guess, and it is the
reason a new report type can be compared the day it arrives with no
configuration at all. But it is a guess about a *shape*, and you know something
it cannot: what kind of report this is, and therefore what identifies a row on
each of its sheets.

This is how to hand it that knowledge, where to put it, and which parts of it
are load-bearing enough to be worth writing down.

## What is actually being decided

Every table on every sheet needs three answers before it can be compared.

| question | detection's answer | what it costs when wrong |
| --- | --- | --- |
| Where does the table start and end? | a run of non-blank rows; the fullest, most heading-like row near its top is the header | a title row read as headers, or two tables read as one |
| What are the columns called? | that header row, with blank cells named `Column A`, `Column B`, … | columns pair by the wrong name, so a whole table reads as changed |
| What identifies a row? | the first column that is complete, distinct and not a measure — or nothing | with nothing, rows pair by **position** instead |

Only the third one is ever guessed conservatively: detection never invents a row
key. If it cannot find one, it says so, and the table is matched by row position
and named in the report as such. That is a real comparison — exact, as long as
both files hold the same rows in the same order — but an inserted row shifts
every row beneath it, so one change reads as many.

Naming the key is what turns those tables from "correct while nothing moves"
into "correct when things move", which is the reason this tool exists.

## Where configuration goes, and what beats what

The spec the run uses is built in this order, each layer merged over the last:

```
detection of the golden file
  └── output_comparison/meta.json              everything, every type
        └── <report type>/meta.json            this type of report
              └── <report type>/case_00n/case.json   this pair only
```

Anything you write beats what detection found for the same key. Merging is
per field — `tolerance` records merge, `ignoreColumns` / `ignoreRows` /
`metadata` accumulate, `tables` merge **per table name**, everything else is
replaced. So a case.json naming one table's `keyColumns` leaves every other
table on that sheet exactly as the report type set it.

**Put it in the report type's `meta.json` when it is true of the report type**
— which sheets carry which dimension, what identifies a row on each. That is
the knowledge you have and detection does not, and it is written once for every
case of that type, present and future.

**Put it in a `case.json` when it is true of that one pair of files** — a row
inserted above a block, an extra column, a table that split differently. Those
are properties of two spreadsheets, not of the report type, and writing them at
the type level breaks every other case.

## The workflow

**1. Run it, and read three sections of `report.md`.**

- `## Matched by row position (n)` — every table with no key, with its row
  count, its column names, and a ready-to-paste JSON snippet. This is the
  worklist.
- `## Differing, and nothing checked them (n)` — cells that changed where layer 1
  never reached. A number above zero here is a coverage hole, and usually the
  first symptom of a mis-detected table.
- `**Comparison integrity**` under a table in `## What changed` — a key column
  that was configured and not found. This means your configuration is landing on
  the wrong table, not that the report is wrong.

**2. See what detection currently thinks.**

```
npm run compare -- --print-spec output_comparison/<type>/case_001
```

It prints the fully merged spec — detection plus every layer above it — with a
comment naming each file that contributed and where to save your edits. Table
names, header rows and keys are all visible there before you change anything.

**3. Edit the layer the fact belongs to, and re-run that one case.**

```
npm run compare -- output_comparison/<type>/case_001
```

## Recipes

### A key-value block with no header row

Cover sheets and report-info blocks are a label in column A and a value in
column B, with nothing above them that is a header. Detection has to pick some
row as the header, so it picks the first one — which deletes that row from the
data, makes a *value* the name of a column, and often leaves the value column as
the only complete-and-distinct one, so the block ends up keyed on the very
figures being tested. A changed value then reads as one row removed and another
added, and never as a change.

Take the blank row above the block as the header row instead:

```json
{
  "sheets": {
    "Report Info": {
      "tables": {
        "Table 1": { "headerRow": 4, "endRow": 21, "keyColumns": ["Column A"] }
      }
    }
  }
}
```

Row 4 is blank, so the columns are named `Column A` and `Column B` — names no
edit to the report can change — every label from row 5 down is data keyed by
itself, and findings read `GCMP USD rate · Column B` instead of `row #13`.

It also makes the `metadata` list work on that sheet: entries like
`Creation Date` are matched as row labels, so they are set aside there the same
way they are everywhere else.

**When the block is small and some rows have no label**, match by position
instead. A five-row exchange-rate block written in a fixed order is exactly
identified by position, and unlike a key it also reaches the row whose label
cell is empty — a key would skip that row, or, with `fillKeyDown` on, collide it
with the row above:

```json
{
  "sheets": {
    "Currency Info": {
      "tables": {
        "Table 2": { "headerRow": 18, "endRow": 23, "keyColumns": [] }
      }
    }
  }
}
```

The rule of thumb: a key when rows can be added, removed or reordered; position
when the block is a fixed form.

### A data table under one or more title blocks

Name it, bound it, and key it. Tables are bounded by the next one's `headerRow`,
so nothing needs recounting as the data grows:

```json
{
  "sheets": {
    "Report Info": {
      "tables": {
        "Table 1": { "headerRow": 4, "endRow": 21, "keyColumns": ["Column A"] },
        "Perils":  { "headerRow": 23, "endRow": 25, "keyColumns": ["Perils Uploaded"] }
      }
    }
  }
}
```

A name detection did not produce — `Perils` here — **adds** a table rather than
replacing one. Give the block it comes out of an explicit `endRow`, or the two
will overlap.

### Two tables printed side by side

Reports put a definitions table in `H:J` beside a key-value block in `A:B`, both
starting on row 1. No blank row separates those, so bound them by column:

```json
{
  "sheets": {
    "Report Info": {
      "tables": {
        "Details":     { "headerRow": 1, "endRow": 17, "columns": "A:B", "keyColumns": ["Field"] },
        "Definitions": { "headerRow": 2, "endRow": 22, "columns": "H:J", "keyColumns": [] }
      }
    }
  }
}
```

Detection writes the bound itself when it sees a gap of two or more blank
columns, so this is usually already right in `--print-spec`. Write it by hand
when a header row would otherwise reach across into a neighbour: without
`columns`, the reader takes the whole width, finds the neighbour's columns hold
data, and folds them into this table under names of their own.

Tables on such a sheet are numbered in reading order — down the sheet, then
across it — so the table to the right of another comes after it.

### Clearing a key that a layer above set

An empty list means "match these rows by position", and it is how a case
switches off a key its report type set:

```json
{ "sheets": { "Geography": { "tables": { "Table 2": { "keyColumns": [] } } } } }
```

Useful when a type-level key belongs to a table that is numbered differently in
this one case — clear it where it does not belong, set it where it does.

### A grouping column written once per group

Reports name a group at the top of it and leave the column blank on every row
beneath. Read literally that column is empty almost everywhere, so it adds
nothing to a key and the rows of one group collide with the next.
`fillKeyDown` reads it the way a person does:

```json
{ "defaults": { "fillKeyDown": true } }
```

Only key building is affected. No cell value is invented, and nothing filled
this way is ever compared.

## Overriding one table

Everything below goes in a `case.json` when it is true of one pair of files, or
in the report type's `meta.json` when it is true of the type. The shape is the
same either way: sheet, then `tables`, then the table's name.

```json
{
  "sheets": {
    "Currency Info": {
      "tables": {
        "Table 3": { "headerRow": 26, "endRow": 30, "keyColumns": ["Currency Name"] }
      }
    }
  }
}
```

### Get the name right

A table's name is whatever detection called it — `Table 3`, or the sheet's own
name when the sheet holds one table. Two places to read it off:

```bash
npx sheet-verify --print-spec output_comparison/<type>/case_003
```

or the `## What was verified` section of that case's `report.md`.

**A name that matches nothing does not do nothing.** It adds a table — see
below — so a misspelling produces a phantom at `A1:B1` with no rows, and the
run fails on it:

```
### Currency Info · Table 9

_`A1:B1` — 2 columns × 0 rows, rows matched by key_

**Comparison integrity** — these make the result untrustworthy
- baseline: key column "Currency Name" not found. Headers: Report ID, 4366
```

If a table you just configured starts reporting an integrity error, check its
name against `--print-spec` before anything else. That is nearly always what it
is.

### Write only what you are changing

Merging is per table **and** per field, so the rest survives. Overriding just a
tolerance leaves the header row, the end row and the key exactly as detection
found them:

```json
{ "sheets": { "Currency Info": { "tables": { "Table 3": { "tolerance": 5 } } } } }
```

```
Table 3 {"headerRow":26,"endRow":28,"keyColumns":["Currency Name"],"tolerance":{"*":5}}
```

Other tables on the sheet are untouched, and so is every other case of the
report type.

### What a table takes

`headerRow` · `endRow` · `columns` (`"H:J"`) · `keyColumns` · `tolerance` ·
`ignoreColumns` · `ignoreRows` · `fillKeyDown` · `matchRowsByPosition`

An empty `keyColumns` clears a key a layer above set, which puts the table back
on positional matching:

```json
{ "sheets": { "Geography": { "tables": { "Table 2": { "keyColumns": [] } } } } }
```

## Adding a table

A name detection did **not** produce is added rather than replacing anything.
That is how a block detection missed entirely gets verified, and it works in a
`case.json` or a report type's `meta.json` alike:

```json
{
  "sheets": {
    "Report Info": {
      "tables": {
        "Perils": { "headerRow": 23, "endRow": 25, "keyColumns": ["Perils Uploaded"] }
      }
    }
  }
}
```

Three things decide whether it lands where you meant:

**`headerRow` is required.** Everything else has a default. Point it at the row
holding the column names — or, for a key-value block with no header of its own,
at the blank row above, which names the columns `Column A`, `Column B`, … and
keeps every labelled row as data.

**Bound it, or it will be bounded for you.** With no `endRow` a table runs
until the next table's header row, counting only tables that share its columns.
That is usually what you want for a table added *below* another. If the block
it comes out of has no explicit `endRow` either, give one of them a bound or
the two overlap and the same cells are compared twice.

**Give it `columns` if it has a neighbour.** Without a column bound the reader
takes the whole width of the sheet, finds the table beside it holds data, and
folds those columns in under names of their own.

```json
{ "Definitions": { "headerRow": 2, "endRow": 22, "columns": "H:J", "keyColumns": [] } }
```

Adding a table to a sheet that holds only one keeps both. The sheet's own table
carries its settings at the sheet level rather than under `tables`, and it is
preserved under the sheet's name:

```json
{ "sheets": { "Ledger": { "tables": { "Bands": { "headerRow": 5, "keyColumns": ["Band"] } } } } }
```

```
| Ledger         | `A1:C4` | 3 columns | 3 rows | key |
| Ledger · Bands | `A5:C7` | 3 columns | 2 rows | key |
```

The original is bounded above the added table rather than running on into it,
so no cell is compared twice.

## Two traps

### Table numbers are positional

`Table 2` means "the second block detection finds on that sheet" — nothing more.
Add a title block above the data and it becomes `Table 3`; merge two blocks and
it becomes `Table 1`.

This is not hypothetical. Across three cases of one report type, the *same*
dimension table — same sheet, same columns, same header row — was `Table 2`,
`Table 3` and `Table 2` again, because the summary block above it split
differently in each file. A single key written at the report-type level is
therefore correct for some of its own cases and reports
`key column "Line of Business" not found` for the rest.

The workaround today is a `case.json` per case that clears the key where it does
not belong and sets it where it does. It works, and it has to be revisited
whenever the source system changes a title block.

### Row numbers are per-case, not per-type

`headerRow` and `endRow` are row numbers in one file. Insert one row above a
block — a new field in a report-info sheet is enough — and every row number
below it is wrong.

Also real: one case of a validation report carries an extra `Program Owners`
row, so its perils table starts at row 24 where the other two cases start at
row 23. The type-level `headerRow: 23` landed on the banner row above, the
configured key was not found, and the table was **not compared at all** — a
silent coverage hole, visible only as three differing cells in
`## Differing, and nothing checked them`.

The fix is a per-case override:

```json
{
  "sheets": {
    "Report Info": {
      "tables": {
        "Table 1": { "endRow": 22 },
        "Perils":  { "headerRow": 24, "endRow": 27 }
      }
    }
  }
}
```

Prefer `endRow` bounds that come from the next table's `headerRow` where you
can, and keep explicit row numbers to the cases that need them.

## When not to configure anything

A positional match is not automatically a defect. Cover sheets, contact pages,
tables of contents, glossaries and disclaimers have no row key because there is
nothing in them that identifies a row — position genuinely *is* the identity,
and layer 2 compares every one of their cells by address regardless. Keying them
would buy nothing.

Spend the effort where a table is large, ordered by something meaningful, and
likely to gain or lose rows between runs. Sort the `## Matched by row position`
list by row count and start at the top: on one report type, six entries out of
thirty-two accounted for almost every row in the list.

## Reading the result

**Start with `## What was verified`.** It lists every table layer 1 read and the
rectangle it read, in both files:

| table | golden | report | columns | rows | rows matched by |
| --- | --- | --- | ---: | ---: | --- |
| Report Info · Table 1 | `A2:B17` | same | 2 | 14 | key |
| Report Info · Table 2 | `H2:J22` | same | 3 | 20 | position |
| Report Info · Table 4 | `A28:BX41` | same | 76 | 13 | key |

This is the cross-reference: hold it against the sheet in Excel and a table
that is missing, or one whose range stops short of where the data does, is
visible immediately. The same caption sits on row 1 of each worksheet in
`compared.xlsx`, and the same ranges are in `diff.json` under each outcome's
`range`.

The range comes from what was *read*, not from what the spec asked for — a spec
with no `endRow` runs to the bottom of the sheet and would tell you nothing. The
golden and report columns are shown separately when they differ, which is worth
knowing on its own: a table that moved between runs explains a wall of
differences that would otherwise read as edits.

The two numbers worth watching after a change:

- `| tables compared | 41, 30 by row position |` in the report header — the
  second number should fall as you add keys.
- `! n differing cell(s) nobody checked` in the run log — cells that changed
  and that layer 1 never reached. This is the number that says a table is
  mis-detected rather than merely unkeyed.

Both are in every `report.md`, so a configuration change can be measured rather
than believed.
