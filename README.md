# sheet-verify

Comparison for generated Excel/CSV outputs that change shape between releases.

Columns align by **header name**, rows align by **business key**, and formula
references resolve to **header names** before comparison. An inserted column is
reported once as a schema change instead of as churn across every downstream
cell.

On a report where a `Premium` column was inserted mid-table and two real defects
were planted:

| comparison | differences reported |
| --- | --- |
| positional cell-by-cell | 25 |
| literal A1 formulas | 10 |
| R1C1 normalised | 6 |
| **header-resolved** | **1** — the real defect |

---

**Contents**

- [Install](#install)
- [Quick start](#quick-start) — two files in a folder, one command
- [What you need to do to use it](#what-you-need-to-do-to-use-it) — the five things worth doing once
- [Detection will not invent a row key](#detection-will-not-invent-a-row-key) — read this one
- [How it works](#how-it-works)
- [Where everything else is](#where-everything-else-is)
- [Notes and limits](#notes-and-limits)
- [Versioning](#versioning) — what a major means here

---

## Install

```bash
npm install exceljs csv-parse jszip
```

Then either vendor `src/` into your test project, or build and depend on it:

```bash
npm install        # builds dist/ on the way out, via the prepare script
```

`dist/` is not committed — it is regenerated from `src/` — so `prepare` runs the
build after any install, including a clone or a git install, and before publish.
Nothing depending on this package can end up with entry points that are not
there.

Playwright transpiles TypeScript itself, so `src/` can be imported directly from
test files with no build step. Only the direct-import path needs the build.

## Quick start

Put two files in a folder and run one command. There is nothing to configure —
sheets, header rows and row keys are worked out from the files themselves.

**1. Make a folder and drop the two files in it:**

```
output_comparison/
  global_standard_cat_report/case_001/
    golden.xlsx      the output you trust
    actual.xlsx      the output under test
```

Or let each file keep the name it was downloaded under, and put the role in the
folder instead — see [Two ways to name a pair](docs/cli.md#two-ways-to-name-a-pair):

```
output_comparison/
  global_standard_cat_report/case_001/
    golden/case_1%1786955263151.xlsx
    current/case_1%1786957329031.xlsx
```

**2. Run it:**

```bash
npx sheet-verify
```

```
✗ Global Standard Cat Report · case_001 · a peril column added mid-table
    global_standard_cat_report/case_001
    1 sheet failing, 1 table to review
    …/case_001/results/report.md

1 case, 1 failing
```

**3. Read `results/`:**

```
output_comparison/global_standard_cat_report/case_001/
  golden.xlsx
  actual.xlsx
  results/
    report.md           everything the run found — start here
    diff.json           the same, structured, for scripts and CI
    differences.xlsx    one row per differing cell, plus the cells that will
                        recalculate — absent when there is nothing to report
    compared.xlsx       every cell checked, a worksheet per table
```

A run covering more than one case also writes `results/run-summary.md` and
`run-summary.xlsx` at the root of the tree: how the run went, grouped by report
type, with a sheet per type. See
[the run summary](docs/reading-a-report.md#the-run-summary).

Exit code is `0` when everything matched and `1` when it did not, so it drops
straight into CI.

That is the whole workflow. Everything below is for when the defaults are not
quite right, or when you want this inside a test suite instead.

## What you need to do to use it

The tool detects the layout. What it cannot know is what your reports *are*, so
these are the five things worth doing once, in the order they pay off.

**1. Put your pairs in a tree.** One folder per report type, one folder per case
under it. Names are free — every folder above a case is just grouping — and the
tree can be as flat or as deep as suits you.

**2. Run it and read the summary line, not just the ticks.** `10 tables
compared` against a report you know holds twelve is the finding. A table with no
row key is **not compared**, and that is reported rather than passed. See
[Detection will not invent a row key](#detection-will-not-invent-a-row-key).

**3. Name the keys detection could not find**, in a `meta.json` beside the
report type. This is the only configuration most trees ever need:

```json
{ "sheets": { "Geography": { "tables": { "Table 2": { "keyColumns": ["Portfolio", "Geography Level"] } } } } }
```

**4. Name the cells that identify the run**, so they stop failing every run. A
report id and a creation timestamp differ by construction; a creator name does
not, if the same account generates every report. See [Report
metadata](docs/configuration.md#report-metadata).

```json
{ "metadata": ["*Report Name", "Report ID", "Creation Date", "Cover!A3"] }
```

**5. Label each case**, one sentence saying what it is for. It titles the case's
report and heads it in the run log, which is what turns a wall of `case_002`
into something readable. See [Naming what a case is](docs/cli.md#naming-what-a-case-is).

Then the loop is: run it, read `results/report.md`, fix the report or
[re-bless](docs/reading-a-report.md#re-blessing-the-golden-output) the golden when the change was
intended.

### What gets detected

From the **golden** file, for every sheet:

- **where each table starts and stops** — a sheet is split at blank rows, so an
  "output info" block above the data becomes its own table rather than swallowing
  the data below it
- **the header row** — the first row of each block
- **the row key** — the column, or pair of columns, that identifies a row

A column qualifies as a key only if **every** value is present and **no** value
repeats. A purely numeric column is refused unless its name says it identifies
something: `Amount` is distinct across three rows and useless across three
thousand, while `Invoice No` is kept. Failing a single column, it tries pairs
from the leftmost few — that is how `PolicyId` + `Period` is found.

### Detection will not invent a row key

**This is the one behaviour to know before trusting a result.**

If no column or pair identifies a row, that table is **not compared at all**. It
is never guessed at, because a wrong key is worse than no key: rows would be
paired arbitrarily and the diff would be confident nonsense.

The table is reported instead, so the gap is visible rather than silent. In
`report.md`:

```
## Sheets to review

- not compared: **Summary** — no keyColumns configured for this table
```

and in the one-line summary as `1 sheet not compared`.

**The signal to watch for is a case reporting fewer tables compared than your
report actually has.** That is not a pass — it is the tool saying it does not
know how to identify a row there. Name the key yourself:

```json
{ "sheets": { "Summary": { "keyColumns": ["Region", "Band"] } } }
```

A table with genuinely no key — a totals band, a pivot — cannot be compared by
this tool at all. Either give it a composite key that is unique, or exclude the
sheet with `ignoreSheets` so it stops being reported.

**A narrated walkthrough of exactly this** — a sheet detection cannot key, the
green-but-incomplete run it produces, and the three-line `case.json` that fixes
it — is a single command:

```bash
npm run build && npm run example:case-json
```

## How it works

Three ideas, and the third is the one that is unusual.

**Columns align by header name.** Not by position. A column inserted at position
5 shifts everything after it; matching on the name means nothing downstream
moves.

**Rows align by business key.** Not by row number. `keyColumns` names the column
or columns that identify a row, so a row inserted in the middle is an *added
row*, not a thousand changed cells. This is the one thing you must configure —
row identity cannot be guessed, and without it a diff degrades into exactly the
positional comparison this exists to avoid.

**Formula references resolve to header names.** `=D2*E2` becomes
`[Sum Insured]@row*[Rate]@row` before anything is compared. This is what the
table at the top measures. R1C1 normalisation is the usual advice and it is only
half right: it is invariant when the inserted column falls *outside* the span
between a formula and its operands, and produces false positives when it lands
between them. Resolving to header names stops encoding position at all.

## Where everything else is

This file is the start. The reference lives beside it, one document per
question:

| you want to | read |
| --- | --- |
| run it, name a pair, choose which cases run, generate a config | **[docs/cli.md](docs/cli.md)** |
| set keys, tables, tolerance, metadata — every option and where it goes | **[docs/configuration.md](docs/configuration.md)** |
| work out *which* correction a report needs, and fix a mis-detected table | **[docs/detection-tuning.md](docs/detection-tuning.md)** |
| read `report.md`, the ledgers, and the verdict | **[docs/reading-a-report.md](docs/reading-a-report.md)** |
| call it from a test suite or from code | **[docs/api.md](docs/api.md)** |

The four commands worth knowing now:

```bash
npx sheet-verify                       # every case in ./output_comparison
npx sheet-verify --print-spec <case>   # what detection made of the files
npx sheet-verify --write-meta <type>   # a starting meta.json, from the pairs
npx sheet-verify --bless               # accept the differences as the new golden
```

## Notes and limits

- **Cached values.** Nothing here evaluates formulas. Comparison uses the value
  the generator wrote plus the formula text, so a generator that writes formulas
  with no stored result leaves value comparison nothing to check — and that is
  the normal case, not a corner one: on one real tree **every one of 379,959
  formula cells** arrived without a result. Two answers, neither of them an
  evaluator. [`--recalc`](docs/cli.md#recalculating-before-comparing) has Excel
  work them out first, which is exact but needs Windows. Failing that, the
  **Will recalculate** sheet in `differences.xlsx` names the cells that would
  move and the change driving each — which cells, not by how much.
- **Cross-sheet formulas.** References to other sheets stay in A1 form, since
  this sheet's headers say nothing about another sheet's layout. A column moved
  on a *referenced* sheet will show as a formula difference. Resolving those to
  header names is possible now that the workbook layer loads every sheet, but it
  is not implemented — nothing has needed it.
- **Duplicate keys.** Rows with a repeated key are excluded and reported as an
  integrity error rather than guessed at.
- **A misspelled key inside `sheets` is not caught.** The unknown-key check runs
  at the top level of a config file — writing `keyColumns` where `sheets` was
  meant is refused outright — but one level down, any field is accepted and
  simply never read. `ignoreColumn` instead of `ignoreColumns` does nothing, and
  says nothing. Check with `--print-spec`, which shows what the run will
  actually use.
- **`endRow` is Excel-only.** The CSV reader takes every row from `headerRow` to
  the end of the file and does not consult it. On a `.csv` the setting is inert.
- **Table numbers are positional.** `Table 2` means the second block detection
  finds on that sheet, so a table named in config can shift if the file gains a
  block above it. A run that suddenly reports `key column not found` on a table
  that was fine is usually this — `--print-spec` shows the current numbering.
- **Row numbers in config are literal.** `headerRow` and `endRow` are absolute,
  so a case whose block sits one row lower than its siblings needs its own
  `case.json`. Configuration written for a report type assumes every case of
  that type has the same shape.
- **Memory.** Around 600 MB of heap for 50k rows. Past ~100k rows use a streaming
  reader or the CSV path, and pair `cellLedger: 'all'` with a `.csv` ledger.
- **Shared formulas.** Excel stores a filled-down formula once and points the
  other cells at the master. The reader resolves these through ExcelJS's
  `cell.formula` getter; reading `cell.value.formula` instead returns the
  master's *address* for every filled cell, which silently hides any formula
  defect in the column. There is a regression test pinning this.
- **Computed references defeat impact tracing.** The **Will recalculate** sheet
  finds what a formula reads by parsing its references, so a function that
  *computes* which cell it reads is opaque to it: `OFFSET($A$1, MATCH(...), 0)`
  is recorded as reading `$A$1` while it reads wherever the arithmetic lands,
  and `INDIRECT` is worse. Not a corner case either — **62,406 of those 379,959
  formulas** were `OFFSET`. Changes reaching a cell that way are not listed, and
  cannot be. `--recalc` has no such gap, because Excel does not reason about
  dependencies, it computes them.
- **`--recalc` needs Windows, Excel installed, and Excel closed.** It drives
  Excel through COM, and COM attaches to a running session rather than starting
  its own — so a run that quit it would close your workbooks and discard unsaved
  work. The tool checks and refuses. It refuses on the other two counts as well
  rather than comparing the files as they arrived, because finding less than the
  flag promised would read as a pass.
- **ExcelJS is barely maintained.** The last stable release is 4.4.0, October
  2023; there has been one prerelease since, in December 2024, and nothing after
  it. It is still the most capable option, which is why the reader sits behind
  `SheetReader`.
- **Do not `npm install xlsx`.** The npm copy is stuck at 0.18.5 and carries two
  high-severity advisories whose fixes exist only on SheetJS's own CDN.

## Versioning

Semantic versioning, with one line drawn deliberately. Releases are tagged,
and [CHANGELOG.md](CHANGELOG.md) says what each one changed.

| bump | what it covers |
| --- | --- |
| **major** | a config key or CLI flag removed or renamed, an export changed, an artefact restructured — a tree that worked stops working |
| **minor** | new options, new features, and **detection changes** |
| **patch** | fixes that change no interface |

**Detection changes ship as minors, and that is the deliberate part.**
Detection is re-derived from the files on every run, and improving it is the
point of the tool — a release that finds a table it used to walk past is doing
its job. Treating each of those as a breaking change would put the major
number somewhere in the double digits within a year and leave it meaning
nothing.

The cost is real, though: finding a new block **renumbers the ones below it**,
and `Table 2` in a config file means "the second block detection finds". So a
minor can cost you a pass over your configuration. That trade is only honest
if the release admits it, so every entry that shifts numbering says so and
says what to recheck:

```bash
npx sheet-verify --print-spec <case>   # the current numbering, per sheet
```

Read the changelog before taking a minor if your config names tables by
number. Configuration that names only sheets and columns is unaffected — which
is the better reason to prefer it. See [Table numbers are
positional](#notes-and-limits).

## Development

```bash
npm install         # builds dist/ on the way out, via the prepare script
npm run check       # typecheck + doc links + 285 Playwright tests
```

`check` is the gate: run it before pushing and it says yes or no once.

Every script, and what it does:

| script | what it does |
| --- | --- |
| `check` | the gate — typecheck, documentation links, and the full test suite |
| `build` | compile `src/` to `dist/`. `npm install` does this for you |
| `typecheck` | types only, no output written |
| `test` | the Playwright suite |
| `links` | every internal documentation link and heading anchor resolves |
| | |
| `compare -- <case>` | compare a case, a report type, or the whole tree |
| `recalc -- <case>` | the same, with Excel working the formulas out first |
| `spec -- <case>` | print what detection made of the files, and stop |
| `bless -- <case>` | accept the new report as the golden output |
| | |
| `write:meta -- <type>` | generate a starting `meta.json` from the pairs |
| `write:expect -- <case>` | record what a run verified, as a guard against losing it |
| | |
| `clean` | delete every `results/` folder, and the run summary |
| `clean:dry` | list what `clean` would delete, and delete nothing |
| `fidelity -- <file>` | what a round-trip through the reader loses |
| | |
| `example` | the six example cases, end to end |
| `example:case-json` | a narrated walkthrough of a sheet needing a `case.json` |

Anything taking an argument needs the `--` separator, which is npm's way of
saying the rest belongs to the script rather than to npm:

```bash
npm run compare -- output_comparison/quarterly_report
npm run recalc -- output_comparison/quarterly_report/case_001
```

Everything that runs the CLI rebuilds `dist/` first. `dist/` is generated and
not committed, so a stale one answers with yesterday's code and says nothing
about having done so — the same silent-staleness hazard `clean` exists for, one
level up. The extra couple of seconds is worth not wondering.

`clean` is not needed for a correct comparison — a run overwrites its own
artefacts. It is needed for an honest one: a case that stops failing leaves its
old `differences.xlsx` behind, and a renamed or deleted case leaves a whole
`results/` folder that still reads as current. Clearing first means everything
present came from the run you just did.

```bash
npm run clean -- --dry                     # list what would go, delete nothing
npm run clean -- output_comparison/validation_report
```

It only removes a folder named `results` that sits beside a golden — a golden
file, or a `golden/` folder holding one — so the two inputs, any `case.json`,
and any unrelated folder of that name are safe.

**Run `clean` and `compare` as separate commands.** Chaining them has been seen
to interleave on Windows, with the clean still deleting while the comparison
writes: the run reports every case as finished and a good part of the tree ends
up with no `results/` at all. The log looked perfect while doing it, which is
the dangerous kind of wrong.
