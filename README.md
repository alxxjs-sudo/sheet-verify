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
- [The command](#the-command)
  - [Two ways to name a pair](#two-ways-to-name-a-pair) — by file name, or by folder
  - [Choosing which cases run](#choosing-which-cases-run)
  - [Naming what a case is](#naming-what-a-case-is)
- [Detection will not invent a row key](#detection-will-not-invent-a-row-key) — read this one
- [How it works](#how-it-works)
- [Cases](#cases) — the folder, the artefacts, re-blessing
- [Describing your report](#describing-your-report) — keys, sheets, tables, tolerance
- [Reading the output](#reading-the-output)
- [What counts as a failure](#what-counts-as-a-failure)
- [Invariants](#invariants)
- [Other entry points](#other-entry-points) — single sheet, workbook, CSV, direct API
- [Template fidelity gate](#template-fidelity-gate)
- [Options reference](#options-reference)
- [Extending it](#extending-it)
- [Notes and limits](#notes-and-limits)

---

## Install

```bash
npm install exceljs csv-parse jszip
```

Then either vendor `src/` into your test project, or build and depend on it:

```bash
npm install && npm run build
```

Playwright transpiles TypeScript itself, so `src/` can be imported directly from
test files with no build step. Only the direct-import path needs `npm run build`.

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
folder instead — see [Two ways to name a pair](#two-ways-to-name-a-pair):

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
    differences.xlsx    one row per differing cell — absent when nothing differed
    compared.xlsx       every cell checked, a worksheet per table
```

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
metadata](#report-metadata).

```json
{ "metadata": ["*Report Name", "Report ID", "Creation Date", "Cover!A3"] }
```

**5. Label each case**, one sentence saying what it is for. It titles the case's
report and heads it in the run log, which is what turns a wall of `case_002`
into something readable. See [Naming what a case is](#naming-what-a-case-is).

Then the loop is: run it, read `results/report.md`, fix the report or
[re-bless](#re-blessing-the-golden-output) the golden when the change was
intended.

## The command

```bash
npx sheet-verify                      # every case in ./output_comparison
npx sheet-verify path/to/case_007     # just one case
npx sheet-verify --bless              # accept the differences as the new golden
npx sheet-verify --print-spec         # show what was detected, as JSON
npx sheet-verify --ledger all         # record matching cells too
npx sheet-verify --help
```

**A case is any folder holding a golden file and the report to compare**, at any
depth. Everything above a case is just grouping, so file them by kind:

```
output_comparison/
  meta.json                          applies to every case below
  global_standard_cat_report/
    meta.json                        applies to this report type
    case_001/  case_002/  …
  pro-forma/
    meta.json
    case_001/
      case.json                      only this case
  analyses/
    marginal/
      case_001/  case_002/  …
```

**Configuration is inherited down the tree.** Every `meta.json` from the root
down applies in order, and the case's own `case.json` wins. So the settings a
whole report type shares are written once, and only a case that genuinely
differs — an extra sheet or two — needs a file of its own.

Cases are named by their path, since `case_001` will exist under every type:

```
✗ Global Standard Cat Report · case_002 · three columns inserted into Geocoding
    global_standard_cat_report/case_002
    1 sheet failing
```

Targeting a subfolder runs only what is under it, but still applies the
configuration above it — so `sheet-verify pro-forma` gives the same verdicts for
those cases as a full run.

### Two ways to name a pair

**By file name.** Matched by prefix, so `golden.xlsx` / `actual.xlsx` is the
convention but not the only option:

| role | any of |
| --- | --- |
| the trusted output | `golden` · `baseline` · `expected` · `before` |
| the output under test | `actual` · `new` · `current` · `after` · `report` |

The prefix must end at a non-word character, so `golden-2026.xlsx` and
`golden 2026.xlsx` are recognised and `golden_2026.xlsx` is not — an underscore
is a word character.

**By folder.** Put the role in the folder and each file keeps whatever name it
arrived with, which is what a downloader produces when it preserves the source
system's name — that name carries the download's timestamp, and renaming it to
`golden.xlsx` throws that away:

```
case_001/
  golden/case_1%1786955263151.xlsx      one spreadsheet, any name
  current/case_1%1786957329031.xlsx     the folder says which side it is
  results/
```

Folder names are the same words as the file prefixes. **Exactly one spreadsheet
per folder** — two stops the run rather than one being picked, because with
names like those a guess is a coin toss:

```
comparison_report/case_001: current/ holds 2 spreadsheets [rep_178699.csv, rep_178657.csv] — it must hold exactly one
```

A half-built case — `golden/` written and `current/` not yet — is reported and
fails the run, rather than quietly dropping out of the count:

```
1 folder(s) meant to be cases could not be run:
  validation_report/case_002: a golden/ folder is here with no current/ folder beside it

5 cases, 5 failing, 1 could not be run
```

`.xlsx`, `.xlsm` and `.csv` all work, and both files must be the same kind.

### Choosing which cases run

A run can be narrowed from the command line by targeting a folder, or written
down in any `meta.json` so a full run does what the tree says:

```json
{ "cases": ["comparison_report/**", "!comparison_report/case_002"] }
```

Paths are relative to the file that carries them, so the root can select by
report type while a type selects its own cases. `*` stays within one path
segment and `**` crosses them; naming a folder takes everything inside it; a
leading `!` excludes. Every file carrying a list narrows further — a case has to
be selected by all of them.

What is left out is counted with the results, never dropped in silence:

```
2 cases, 2 failing — 4 not selected by "cases"
```

### Naming what a case is

Two keys change nothing about the comparison and a great deal about reading it.
`reportType` goes in a report type's `meta.json`; `label` goes in a case's
`case.json` and is one sentence saying what that case is for:

```json
{ "label": "an extra Cat Model Version row shifts the whole perils block down" }
```

```
✗ Validation Report · case_003 · an extra Cat Model Version row shifts the whole perils block down
    validation_report/case_003
    1 sheet failing
```

The label also titles that case's `report.md`, with the folder name kept
underneath so the file is still findable from what it says:

```markdown
# an extra Cat Model Version row shifts the whole perils block down

_Validation Report · case_003_
```

`reportType` is inherited like any other setting. **`label` deliberately is
not** — one written a folder above would head every case beneath it with the
same sentence and distinguish none of them. A label that merely repeats the
folder name is treated as no label at all, so nothing reads `case_003 ·
case_003` before you have written a real one.

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

### Correcting anything else

`--print-spec` shows exactly what was worked out from your files:

```bash
npx sheet-verify --print-spec > output_comparison/comparison_report/case_001/case.json
```

Edit that file and it is layered over the detection on the next run. Everything
in it is optional — you only write the parts you want to change, and the rest of
the detection is kept. The most common entry by far silences a timestamp that is
rewritten on every run:

```json
{
  "defaults": { "ignoreRows": ["Generated At"], "ignoreColumns": ["Run Id"] }
}
```

**Put it at the right level.** The same file format works as `meta.json` at any
folder above a case, so a setting that applies to a whole report type belongs
there rather than copied into every case:

| goes in | applies to |
| --- | --- |
| `output_comparison/meta.json` | every case — timestamps, run ids |
| `reports/<type>/meta.json` | that report type — its keys, its decoration sheets |
| `<case>/case.json` | one case — an extra sheet it happens to have |

Lists accumulate down the layers rather than replacing each other, so a type
excluding its glossary and a case excluding one more sheet ends up excluding
both. `--print-spec` lists every layer that was applied, so it is always
answerable which rules a result came from.

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

## Cases

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

### Re-blessing the golden output

Schema changes between releases are expected. When a diff is correct, accept it
explicitly rather than editing the golden file by hand:

```bash
npx sheet-verify --bless               # from the CLI
UPDATE_SHEET_BASELINE=1 npx playwright test   # from a test run
```

Golden outputs are committed to git, so the change is reviewed in the pull
request — which is the point. A silent baseline edit is how a defect becomes
permanent.

On a pair named [by folder](#two-ways-to-name-a-pair), blessing writes the new
golden under **the name the new report arrived with** and removes the file it
replaced, so the golden's name never claims a download it no longer holds:

```
✓ Comparison Report · case_001
    golden replaced by rep_1786957329031.csv, and rep_1786955263151.csv removed
```

### report.md

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

### diff.json

The same content, addressable. This is what a CI gate or a dashboard reads.
It holds four things that exist nowhere else, because they are not properties of
any cell pair:

- **the verdict** — `ok` and `reviewOnly`
- **sheets with no cells** — a sheet added (nothing to compare) or removed (a
  defect that no cell can point at)
- **rows dropped before comparison** — `duplicateKeysBase` / `duplicateKeysNext`
- **integrity errors and invariant failures** — properties of the file, not of a
  comparison

### differences.xlsx

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

**When nothing differs, the file is not written at all**, and one left by an
earlier run is removed. An empty differences file reads as a fault rather than
as the answer, and a stale one is worse: it describes a comparison that no
longer holds.

Its absence is therefore not a pass. Some failures have no cells to point at —
a removed sheet, a failed invariant, a duplicate key — and those appear in
`report.md` and `diff.json` only. **The verdict is the exit code and
`report.md`, never the presence of this file.**

### compared.xlsx

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

### Controlling what gets written

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

## Describing your report

Everything here is detected automatically by the CLI — reach for it when the
detection needs correcting, or when you are calling the API directly.

### Keys

`keyColumns` is the only required option when calling the API. One column, or
several forming a composite key:

```ts
{ keyColumns: ['PolicyId'] }
{ keyColumns: ['PolicyId', 'Period'] }
```

Rows whose key is blank are skipped. Rows with a **repeated** key are excluded
and reported as an integrity error rather than guessed at.

### Multi-sheet workbooks

Each sheet gets its own spec; `defaults` covers what they share. The whole
workbook is parsed once, not once per sheet.

```ts
{
  defaults: { tolerance: { '*': 0.01 }, ignoreColumns: ['Generated At'] },
  sheets: {
    Policies: { keyColumns: ['PolicyId'] },
    Premiums: { keyColumns: ['PolicyId', 'Period'] },
    Regions:  { keyColumns: ['Region'], headerRow: 2 },
  },
  ignoreSheets: ['Scratch'],
}
```

`defaults` merges **per field**: `tolerance` records merge, `ignoreColumns`,
`ignoreRows` and `invariants` accumulate, everything else is replaced by the
per-sheet value. A blanket override would be a footgun — a per-sheet `tolerance`
replacing the default record would silently drop the `*` fallback.

Sheet names match case-insensitively. A name in `sheets` matching no sheet in
either file is reported as an error rather than ignored: it is almost always a
typo, and the alternative is a sheet silently going unchecked forever.

**The golden output's sheet list is the contract:**

| sheet is in | verdict |
| --- | --- |
| both | compared, all layers |
| the new output only | **noted, not compared** — there is nothing to compare it against |
| the golden output only | **defect** — output that used to be produced is gone |
| both, but no `keyColumns` | **not compared** — see [above](#detection-will-not-invent-a-row-key); the gap is reported, never guessed at |

The asymmetry is deliberate. A new sheet is additive and harmless; a sheet that
has vanished means consumers stopped receiving something they expect.

### Several tables on one sheet

Most generated sheets carry a small "output info" block above the data — report
name, creator, release, a generation timestamp. That is two tables on one sheet,
and a single `headerRow` cannot describe it: reading from row 1 runs to the
bottom of the sheet and swallows the data table below.

Declare each table instead. They are bounded by the next table's `headerRow`, so
nothing needs re-counting as the data grows:

```ts
const info = {
  headerRow: 1,
  keyColumns: ['Field'],          // a key-value block is keyed by field name
  ignoreRows: ['Generated At'],   // rewritten every run
};

{
  defaults: { headerRow: 8 },
  sheets: {
    Policies: { tables: { Info: info, Detail: { keyColumns: ['PolicyId'] } } },
    Premiums: { tables: { Info: info, Detail: { keyColumns: ['PolicyId', 'Period'] } } },
  },
}
```

Each table is compared independently and reported under its own name, so a
release bump in the info block never mixes with a defect in the data:

```
SHEET "Premiums · Info" — 1 value
    Release · Value @B5: "4.2.0" → "4.3.0"

SHEET "Premiums · Detail" — 1 value
    P-1003 / 2026-08 · Gross @C14: 6000 → 6900  (Δ 900)
```

### Tolerance and exclusions

```ts
{
  keyColumns: ['PolicyId'],
  tolerance: { '*': 0, 'Annual Cost': 0.01, Commission: 0.01 },
  ignoreColumns: ['Generated At', 'Run Id'],
  ignoreRows: ['Generated At'],
}
```

`tolerance` takes a single number for every column, or a record with `*` as the
fallback. `ignoreColumns` excludes a column; `ignoreRows` excludes a row **by
key** — the row-wise counterpart, for key-value blocks where the per-run value
is a row and no column exclusion can reach it.

### Report metadata

Some cells identify the run rather than describe it. A report's name, its id and
its creation timestamp are minted fresh every time, so comparing them produces a
difference on every single run — which makes a clean run impossible and teaches
whoever reads the report to skip the first section. Name them once, at the top
of the spec:

```json
{
  "metadata": [
    "*Report Name",
    "Report ID",
    "Creation Date",
    "Cover!A3"
  ]
}
```

**The test is whether the value differs _by construction_, not whether it
sounds like metadata.** A generated id does. A creator name does not, if the
same account generates every report — there it is expected to stay the same,
and a change means the wrong account ran it. That is worth stopping on, so
leave it out of this list. The same goes for anything the figures depend on:
view of risk, currency, model version, the as-at date of the data. Those look
like header furniture and are not — if one of them moved, the numbers
underneath it should have moved too.

An entry is a **label**, matched against a cell's text and taking the value
beside it — `"Report ID"` covers `A1 "Report ID"` with `B1 4542`, and a fused
`="Report ID: " & id` as well. Or it is a **cell reference**, `"Cover!A3"`, for
a bare date with no label of its own. A `*` stands for a run of text, which is
what lets one pattern cover `Facility Report Name`, `RiskPlay Report Name` and
`Pro-Forma Report Name`. Either form takes a sheet qualifier —
`"Report Info!Report ID"` — for a word that means run identity in a header
block and a column heading somewhere else.

Matching cells are skipped by **both** layers, and nothing downstream of them is
chased either. They are not hidden: `report.md` lists every one under **Not
verified, on purpose**, with both values, so an id that moved when you did not
expect it is still there to be seen. It simply does not fail the run.

## Reading the output

Start with the one-line summary, then `report.md`, then `differences.xlsx` if
you need to work through individual cells.

Three distinctions do most of the work:

**Compared vs not compared.** A table with no row key is skipped, not passed.
Check the header line — `10 tables compared` — against what your report actually
contains, and treat any `not compared` in the summary as work to do rather than
a clean result. See [Detection will not invent a row
key](#detection-will-not-invent-a-row-key).

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

## Invariants

A golden file proves the new output matches the old one, never that either is
*correct*. A wrong value present in both passes forever. Invariants close that
gap by asserting properties of the output alone:

```ts
import { invariants as inv } from 'sheet-verify';

{
  keyColumns: ['PolicyId'],
  invariants: [
    inv.noErrorValues(),                  // #REF!, #DIV/0! anywhere
    inv.notBlank('PolicyId', 'Holder'),
    inv.unique('PolicyId'),
    inv.inRange('Rate', 0, 1),
    inv.derived('Commission', (_row, num) => num('Annual Cost') * 0.1, 0.005),
  ],
}
```

`derived()` recomputes a column independently, so it catches calculation bugs
present in the golden output too. The `invariant-catch` example demonstrates
exactly this: a rate of 1.4 in *both* files, where the comparison passes and only
`inRange` catches it.

## Other entry points

The CLI is the shortest path. These are for wiring the comparison into a test
suite or a script.

**Inside Playwright**, so a report is checked as part of the suite:

```ts
import { test } from '@playwright/test';
import { expect } from 'sheet-verify/matcher';

test('monthly policy export', async () => {
  const actual = await app.generateReport('2026-08');

  await expect(actual).toMatchCase('output_comparison/comparison_report/case_001', {
    sheets: { Policies: { keyColumns: ['PolicyId'] } },
  });
});
```

The matcher takes an explicit spec rather than detecting one — a test should
say what it checks. To detect instead, call `detectSpec()` and pass the result.

**Detection on its own**, if you want the spec without running a comparison:

```ts
import { detectSpec, detectWorkbook } from 'sheet-verify';

const spec = await detectSpec('golden.xlsx');       // ready to pass to runCase
const layout = await detectWorkbook('golden.xlsx'); // tables, bounds, candidate keys
```

The narrower matchers below are still supported.

**One sheet against one baseline file:**

```ts
await expect(actual).toMatchSheetBaseline('baselines/policies.xlsx', {
  keyColumns: ['PolicyId'],
  sheet: 'Policies',           // name or 0-based index; default 0
});
```

**A whole workbook, without a case folder:**

```ts
await expect(actual).toMatchWorkbookBaseline('baselines/policies.xlsx', {
  sheets: { Policies: { keyColumns: ['PolicyId'] } },
});
```

Both attach the diff to the test result as `sheet-diff.txt` and
`sheet-diff.json`, so it lands in the Playwright HTML report and in Allure.

**CSV** works everywhere a workbook does. It is one table, so it presents itself
as a single sheet named `CSV`:

```ts
await expect('out.csv').toMatchCase('output_comparison/comparison_report/case_002', {
  sheets: { CSV: { keyColumns: ['PolicyId'] } },
  defaults: { csv: { strictDialect: true } },
});
```

`csv.strictDialect` fails on BOM, delimiter or line-ending drift — invisible
changes that break downstream consumers.

**Direct, with no test framework:**

```ts
import { verifySheet, verifyWorkbook, formatReport, formatWorkbookReport } from 'sheet-verify';

const diff = await verifySheet('golden.xlsx', 'actual.xlsx', { keyColumns: ['PolicyId'] });
console.log(formatReport(diff));

const wb = await verifyWorkbook('golden.xlsx', 'actual.xlsx', {
  sheets: { Policies: { keyColumns: ['PolicyId'] } },
});
console.log(formatWorkbookReport(wb));
```

## Template fidelity gate

Editing a template means reading and re-saving it, and libraries silently drop
parts of the format when they do. Check real templates before trusting any of
them:

```bash
npx sheet-fidelity templates/*.xlsx
```

Exits non-zero if a template loses content. Measured against ExcelJS 4.4.0,
conditional formatting, data validation, merged cells, images, auto-filters and
frozen panes all survive; `calcChain.xml` is dropped harmlessly and printer
settings are lost. That was measured on constructed fixtures — verify against
*your* templates rather than trusting it.

If a template does fail the gate, do not rewrite it. Write only into a data sheet
the template's own formulas read from, so the fragile parts are never touched.

## Options reference

**Case options** — `toMatchCase` / `runCase`, on top of the workbook options:

| option | default | meaning |
| --- | --- | --- |
| `cellLedger` | `'differences'` | scope of `differences.xlsx`: `'differences'` \| `'all'` \| `'none'` |
| `comparedLedger` | `true` | write `compared.xlsx` at all |
| `names` | see [Cases](#cases) | file names within the case folder |
| `updateGolden` | `false` | overwrite the golden output and pass |
| `createMissingGolden` | `true` | create it on first run rather than failing |

**Workbook options** — `toMatchWorkbookBaseline` / `verifyWorkbook`:

| option | default | meaning |
| --- | --- | --- |
| `sheets` | `{}` | per-sheet spec, keyed by worksheet name |
| `defaults` | `{}` | applied to every sheet, then overridden per sheet |
| `ignoreSheets` | `[]` | worksheets excluded entirely |
| `metadata` | `[]` | run identity — read, listed, never judged |
| `strictSheets` | `false` | treat added and unconfigured sheets as failures |
| `matchUnkeyedRowsByPosition` | `true` | compare a table with no key by row order instead of leaving it unchecked |

**Tree options** — read from `meta.json` / `case.json` by the CLI, and ignored
by the API, which is handed one case at a time:

| option | scope | meaning |
| --- | --- | --- |
| `cases` | any `meta.json` | which cases beneath it run — globs, `!` excludes |
| `reportType` | inherited | the kind of report, shown in the log and the report |
| `label` | the case's own `case.json` only | one sentence saying what the case is for |
| `source` | any | provenance, carried and never read |
| `//…` | any | a note. JSON has no comments, so any key starting `//` is ignored |

**Per-sheet options**, and the whole spec for a single-sheet comparison:

| option | default | meaning |
| --- | --- | --- |
| `keyColumns` | *required* | column(s) identifying a row |
| `tables` | – | several tables on one sheet, keyed by name |
| `sheet` | `0` | worksheet name or index — single-sheet API only |
| `headerRow` | `1` | 1-based row holding headers |
| `endRow` | *last row* | 1-based last row; set for you when a sheet declares `tables` |
| `tolerance` | `0` | number, or per-column record with `*` fallback |
| `ignoreColumns` | `[]` | columns excluded from comparison |
| `ignoreRows` | `[]` | rows excluded from comparison, by key |
| `keySeparator` | `␟` | joins composite key parts |
| `formulaMode` | `'header'` | `'header'` \| `'r1c1'` \| `'a1'` |
| `compareFormulas` | `true` | compare formula logic at all |
| `requireCachedValues` | `true` | fail if formula cells have no cached result |
| `trimStrings` | `true` | trim whitespace on string values |
| `looseHeaders` | `true` | match headers ignoring case and extra spaces |
| `strictSchema` | `false` | treat schema changes as failures |
| `invariants` | `[]` | baseline-free assertions |
| `csv.delimiter` | detected | field delimiter |
| `csv.numeric` | `'auto'` | how CSV text becomes numbers |
| `csv.strictDialect` | `false` | fail on BOM/delimiter/line-ending drift |

**Report options** — accepted alongside any of the above:

| option | default | meaning |
| --- | --- | --- |
| `limit` | `20` | rows listed per section before truncating |
| `showCascades` | `false` | list cascaded differences, not just count them |

## Extending it

The reader sits behind an interface, so ExcelJS can be replaced without touching
any comparison logic:

```ts
import { registerReader } from 'sheet-verify';

const unregister = registerReader(myReader);   // returns a disposer
```

A reader implementing `SheetReader` covers `verifySheet`. To take part in
workbook and case comparisons it must implement `readWorkbook` too — that is the
method that parses a file once and returns a model per requested table.

## Notes and limits

- **Cached values.** Nothing here evaluates formulas. Comparison uses the value
  the generator wrote plus the formula text. If the generator writes formulas
  with no cached result, value comparison has nothing to check — hence
  `requireCachedValues`, on by default.
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
- **ExcelJS is dormant.** No release since October 2023. It is still the most
  capable option, which is why the reader sits behind `SheetReader`.
- **Do not `npm install xlsx`.** The npm copy is stuck at 0.18.5 and carries two
  high-severity advisories whose fixes exist only on SheetJS's own CDN.

## Development

```bash
npm install
npm run build       # tsc -> dist/
npm test            # 257 Playwright tests
npm run typecheck
npm run example     # run the six example cases
npm run compare     # the CLI, from source
npm run clean       # delete every results/ folder under output_comparison
```

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
