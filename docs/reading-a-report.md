# Reading a run

What a case folder holds after a run, what each artefact is for, and how to
read the verdict.

A case is a folder holding everything about one report: the golden output, the
report under test, and the artefacts describing what the comparison did. It can
be reviewed, archived, or attached to a ticket as a unit.

The CLI reads the two files straight out of the folder and writes its output to
`results/`, so nothing it produces can be mistaken for one of the inputs.

Used from code, the two files need not be in the folder to begin with — pass the
new report from wherever your app wrote it and it is **copied in**, whatever the
outcome, so a CI failure can be opened next to the golden output it was judged
against rather than hunted down in a build artefact:

```ts
import { runCase } from 'sheet-verify';

const result = await runCase(actual, 'cases/monthly-policy-export', spec);
if (!result.ok) {
  console.error(result.summary);
  console.error(`details: ${result.files.report}`);
}
```

## Re-blessing the golden output

Schema changes between releases are expected. When a diff is correct, accept it
explicitly rather than editing the golden file by hand:

```bash
npx sheet-verify --bless               # from the CLI
UPDATE_SHEET_BASELINE=1 npx playwright test   # from a test run
```

Golden outputs are committed to git, so the change is reviewed in the pull
request — which is the point. A silent baseline edit is how a defect becomes
permanent.

On a pair named [by folder](cli.md#two-ways-to-name-a-pair), blessing writes the new
golden under **the name the new report arrived with** and removes the file it
replaced, so the golden's name never claims a download it no longer holds:

```
✓ Comparison Report · case_001
    golden replaced by rep_1786957329031.csv, and rep_1786955263151.csv removed
```

## report.md

The read, and the one file to start from. Ordered so the most actionable thing
is first: what was verified, then integrity problems, then removed sheets, then
each failing table in full, then what layer 1 could not reach, then everything
merely worth reviewing.

```markdown
# geography rebuilt: 851 cells restated across the breakdown

_Comparison Report · case_001_

**Differences found.**

| golden | `golden.xlsx` — 12 sheet(s) |
| report | `actual.xlsx` — 12 sheet(s) |
| tables compared | 24, 3 by row position |
| cells differing | 894 |

**Two-layer verification — both layers ran over every shared sheet.**

- **Layer 1, by name and key** — 12,904 cells, across tables whose columns were
  paired by header name and rows by the values that identify them.
- **Layer 2, by address** — 15,117 cells, every one in both files compared A1
  against A1 … 2,213 cells rest on this layer alone.

## What changed

### Geography · Table 2

**Value changes (851)** — grouped by column, largest group first …
```

Sections that would otherwise run to walls of rows are grouped by column and
folded behind a `Show` toggle; sheet-by-sheet blocks carry their own counts.
Nothing is truncated — a terminal has a reason to elide, a file does not.

## diff.json

The same content, addressable. This is what a CI gate or a dashboard reads.
It holds four things that exist nowhere else, because they are not properties of
any cell pair:

- **the verdict** — `ok` and `reviewOnly`
- **sheets with no cells** — a sheet added (nothing to compare) or removed (a
  defect that no cell can point at)
- **rows dropped before comparison** — `duplicateKeysBase` / `duplicateKeysNext`
- **integrity errors and invariant failures** — properties of the file, not of a
  comparison

## differences.xlsx

For working through failures cell by cell. It arrives as a real Excel table —
filter buttons, banded rows, header frozen, columns already wide enough to read,
status colour-coded — so it opens ready to sort and filter.

| column | meaning |
| --- | --- |
| Sheet, Table, Row key, Column | which cell, in business terms |
| Status | `value-differs`, `formula-differs`, `type-differs`, `within-tolerance`, `ignored-column`, `ignored-row`, `row-added`, `row-removed`, `column-added`, `column-removed`, `match` |
| Root cause | `yes` for a cause, `no` for something downstream of one |
| Golden cell, Actual cell | the addresses, which differ when a column moved |
| Golden value, Actual value, Delta | what changed, and by how much — written as numbers, so sorting works |
| Tolerance | what the cell was judged against |
| Golden formula, Actual formula | the original A1 text |

Only differing cells get a row. A cell that was **ignored** still earns one when
it actually differs — that is the evidence the exclusion is doing work — but not
when it was ignored *and* identical.

### The "Will recalculate" sheet

A second worksheet, present whenever a formula reads something that differs.

**Read this one when a difference you can see in Excel has no row in
`Differences`.** These reports arrive from the generator with formulas and no
stored results — Excel works them out when you open the file — so a formula
whose inputs moved has nothing to compare and cannot appear as a difference.
Open the two files side by side and a whole column of totals plainly differs;
look for it in `Differences` and it is not there. The comparison did not miss
it. It reported the *input* that moved, usually on another sheet, and this
sheet is the bridge between the two.

| column | meaning |
| --- | --- |
| Sheet, Cell, Column | the cell that will come out different |
| How | `reads it directly`, or `through another formula` |
| Driven by sheet, Driven by cell | the change that reaches it |
| Golden / Actual (value or formula) | what that driving cell holds on each side |

Chains are followed, so a cell two steps downstream names the cell above it and
that row names the original difference. A driving cell that is itself a formula
has no stored value to quote, so those two columns are blank and the row is
greyed.

What it does **not** say is what the recalculated number will be. Nothing here
evaluates formulas, so there is no such number to write — the sheet answers
*which cells and driven by what*, not *by how much*. To get the numbers
themselves, open both files in Excel and save them before comparing: that bakes
the computed results in, and they then compare like any other value.

The same list is in `report.md` under **Will recalculate differently**.

**When nothing differs, the file is not written at all**, and one left by an
earlier run is removed. An empty differences file reads as a fault rather than
as the answer, and a stale one is worse: it describes a comparison that no
longer holds.

Its absence is therefore not a pass. Some failures have no cells to point at —
a removed sheet, a failed invariant, a duplicate key — and those appear in
`report.md` and `diff.json` only. **The verdict is the exit code and
`report.md`, never the presence of this file.**

## The run summary

A run covering more than one case also writes, at the root of what was run:

```
output_comparison/!summary/
  run-summary.md      how the run went — for reading and pasting into a message
  run-summary.xlsx    the same, for sorting and filtering
```

The name starts with `!` so the folder sorts to the top of the tree, above every
report type — this is the first thing to read after a run, and `results/` would
land alphabetically somewhere in the middle of them. `_summary` was the obvious
choice and is wrong: Windows orders a leading underscore *after* letters, so it
sorted to the bottom.

A run given `--results <name>` puts its summary in `!summary/<name>/`, so a
plain run and a `--recalc` run keep both instead of one overwriting the other.

### Keeping it honest

The file sits at the tree root and describes *the run*, so a run narrowed to one
report type replaces the last full one. That would leave "3 case(s)" reading as
the size of the tree, so a narrowed run says what it was:

> **Scoped run — `facility_report` only.** Cases elsewhere in the tree were not
> compared and are not represented below.

It is written even for a single case, for the same reason: skipping it would
leave the previous run's summary in place, claiming thirty-four cases when one
was just compared.

To get the whole tree back without re-comparing anything:

```bash
npm run summary
```

That rebuilds `!summary/` from the `diff.json` each case's last run left behind
— the whole tree however narrow the last run was, and instant where re-running
is a minute, or half an hour with `--recalc`. It says so at the top, and a case
with no results on disk is listed as **never run** rather than left out: a case
missing from an overview is indistinguishable from a case that passed.

Every other artefact describes one case, which is the right shape for fixing
something and the wrong shape for the question asked first: **how did the run
go?** That answer used to live only in terminal scrollback, where it is gone as
soon as anyone scrolls and cannot be sent to someone who was not watching.

Both are grouped **by report type**, because that is the unit people work in — a
release breaks a *kind* of report, and eight failures on one type with the rest
clean is a different morning from one failure on each of eight types. A type
that never named itself in a `meta.json` is grouped under its folder path.

The markdown opens with a table of every type and its counts, then a section per
type listing each case with its label and result, failures first. The workbook
has an **Overview** sheet with the same table plus a totals row, and one sheet
per report type with more columns than the markdown carries — sheets compared,
tables failing, tables not compared, unchecked differing cells, whether the run
was recalculated, and the path to each case's own report.

**Watch the "Unchecked differing" column.** It counts cells that differ and
that layer 1 never looked at, and it is the one number that can be non-zero on a
case that *passed* — which is what a table with no row key looks like from the
outside. The markdown calls it out under any type where it is non-zero.

`npm run clean` removes the summary along with the `results/` folders, for the
same reason: one left from the last run reads as current.

## compared.xlsx

The full record of what was checked, in a deliberately plain column set: row key,
column, the cell address on each side, both values, and the verdict. "We compared
this and it was fine" is the claim a golden-file suite is really making, and this
is the file that states it.

It is **split one worksheet per compared table**, because this is the file that
grows with the report:

```
Policies · Info      10 cells      Commissions · Info      10 cells
Policies · Detail    35 cells      Commissions · Detail    18 cells
Premiums · Info      10 cells      Regions · Info          10 cells
Premiums · Detail    50 cells      Regions · Detail        20 cells
```

A five-sheet report becomes ten tabs to scan rather than one sheet of
everything, and each tab stays clear of Excel's million-row ceiling on its own.
A table that would exceed it anyway is truncated with a row saying so.

Anything that is not a plain match is highlighted, in the same colours
`differences.xlsx` uses. Matches are left unpainted — in a tab that is mostly matches,
colour is only useful if it marks the exceptions. Grey marks a cell that was
*excluded* rather than wrong, so `ignoreRows` and `ignoreColumns` show their
effect here.

## Controlling what gets written

```ts
await expect(actual).toMatchCase(dir, {
  ...spec,
  cellLedger: 'all',                  // fold matches into differences.xlsx too
  names: { cells: 'differences.csv' },      // .csv streams instead of building a workbook
  comparedLedger: false,              // skip compared.xlsx entirely
});
```

The ledger format follows the file name — `.xlsx` gets the formatted table,
`.csv` gets streamed text. Use `.csv` with `cellLedger: 'all'` on a large case:
that combination grows with rows × columns, and streaming keeps it off the heap.

## Reading the output

Start with the one-line summary, then `report.md`, then `differences.xlsx` if
you need to work through individual cells.

Three distinctions do most of the work:

**Compared vs not compared.** A table with no row key is skipped, not passed.
Check the header line — `10 tables compared` — against what your report actually
contains, and treat any `not compared` in the summary as work to do rather than
a clean result. See [Detection will not invent a row
key](../README.md#detection-will-not-invent-a-row-key).

**Root cause vs cascade.** One wrong input feeding two formulas is reported as
one cause and two consequences, not three failures. Every difference is
reported either way — a cascade is still a difference, and hiding it would
understate the reach of the change — and both `report.md` and
`differences.xlsx` mark each one `yes` or `no` in *Root cause*. Fix causes;
consequences follow.

**Defect vs review.** A defect fails the run. A review item — an inserted
column, a new sheet, added rows — passes, but is reported so somebody looks.
`reviewOnly: true` means "it changed, nothing is wrong".

## What counts as a failure

| layer | meaning | verdict |
| --- | --- | --- |
| Sheets | worksheets added or moved | review |
| Sheets | a worksheet **removed** | **defect** |
| Schema | columns added, removed, moved | review |
| Row population | keys present in one file only | review |
| Values | cell differences, beyond tolerance | **defect** |
| Formulas | calculation logic, normalised | **defect** |
| Types | same rendering, different type | **defect** |
| Invariants | properties that must hold regardless | **defect** |
| Integrity | duplicate keys, uncached formulas, dialect drift | **defect** |

A renamed column is reported as one added plus one removed, not as a rename —
nothing here guesses that `Cost` became `Annual Cost`.

To tighten it:

```ts
{ strictSchema: true }   // any column change fails
{ strictSheets: true }   // added and unconfigured sheets fail too
```
