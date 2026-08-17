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
- [Quick start](#quick-start)
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

Pick one report, give it a case folder, and name the column that identifies a
row. Everything else has a working default.

```ts
import { test } from '@playwright/test';
import { expect } from 'sheet-verify/matcher';

test('monthly policy export', async () => {
  const actual = await app.generateReport('2026-08');

  await expect(actual).toMatchCase('cases/monthly-policy-export', {
    sheets: {
      Policies: { keyColumns: ['PolicyId'] },
      Premiums: { keyColumns: ['PolicyId', 'Period'] },
    },
  });
});
```

**On the first run** there is no golden output, so the report becomes one and the
test passes. Open `cases/monthly-policy-export/golden.xlsx`, check it is actually
correct, and commit it. From then on it is the contract.

**On later runs** the report is compared against it. If they differ, the test
fails and the case folder holds everything needed to work out why.

To see all of this working before wiring up your own report:

```bash
npm run build && npm run example
```

That runs six cases covering the outcomes you will meet in practice — see
[examples/](examples/).

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

A case is a folder holding everything about one report type. Everything for one
report sits together, so it can be reviewed, archived, or attached to a ticket
as a unit.

```
cases/monthly-policy-export/
  golden.xlsx     committed; the contract
  actual.xlsx     the latest run, copied in
  diff.txt        the differences, human-readable
  diff.json       the same, structured
  cells.xlsx      one row per differing cell, formatted for working through
  compared.xlsx   every cell checked, one worksheet per compared table
```

The new report is copied into the folder **whatever the outcome**, so a CI
failure can be opened next to the golden output it was judged against rather
than hunted down in a build artefact.

Outside Playwright, `runCase()` does the same and returns the result:

```ts
import { runCase } from 'sheet-verify';

const result = await runCase(actual, 'cases/monthly-policy-export', spec);
if (!result.ok) {
  console.error(result.summary);
  console.error(`details: ${result.files.diffText}`);
}
```

### Re-blessing the golden output

Schema changes between releases are expected. When a diff is correct, accept it
explicitly rather than editing the golden file by hand:

```bash
UPDATE_SHEET_BASELINE=1 npx playwright test
```

Golden outputs are committed to git, so the change is reviewed in the pull
request — which is the point. A silent baseline edit is how a defect becomes
permanent.

### diff.txt

The read. Ordered so the most actionable thing is first: integrity problems,
then removed sheets, then each failing table with its full detail, then
everything merely worth reviewing collapsed to a line each.

```
baseline  golden.xlsx   5 sheet(s)
actual    actual.xlsx   6 sheet(s)
          10 tables compared · 1 added

SHEET "Premiums · Detail" — 1 value
  VALUE CHANGES (1 root cause, 2 cascaded)
    P-1003 / 2026-08 · Gross @C14: 6000 → 6900  (Δ 900)

SHEETS TO REVIEW
  ~ sheet "Policies · Detail" — 3 schema
  + sheet "Premium Detail" — new, not compared
```

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

### cells.xlsx

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
`cells.xlsx` uses. Matches are left unpainted — in a tab that is mostly matches,
colour is only useful if it marks the exceptions. Grey marks a cell that was
*excluded* rather than wrong, so `ignoreRows` and `ignoreColumns` show their
effect here.

### Controlling what gets written

```ts
await expect(actual).toMatchCase(dir, {
  ...spec,
  cellLedger: 'all',                  // fold matches into cells.xlsx too
  names: { cells: 'cells.csv' },      // .csv streams instead of building a workbook
  comparedLedger: false,              // skip compared.xlsx entirely
});
```

The ledger format follows the file name — `.xlsx` gets the formatted table,
`.csv` gets streamed text. Use `.csv` with `cellLedger: 'all'` on a large case:
that combination grows with rows × columns, and streaming keeps it off the heap.

## Describing your report

### Keys

`keyColumns` is the only required option. One column, or several forming a
composite key:

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
| both, but no `keyColumns` | listed as *not compared*, so the gap stays visible |

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

## Reading the output

Start with the one-line summary, then `diff.txt`, then `cells.xlsx` if you need
to work through individual cells.

Two distinctions do most of the work:

**Root cause vs cascade.** One wrong input feeding two formulas is reported as
one cause and two consequences, not three failures. `diff.txt` shows the causes
and counts the cascades; `cells.xlsx` marks each row `yes` or `no` in *Root
cause*. Fix causes; consequences follow.

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

Cases are the recommended path. These are narrower and still supported.

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

**CSV** has no worksheets, so it uses the single-sheet API:

```ts
await expect('out.csv').toMatchSheetBaseline('baselines/out.csv', {
  keyColumns: ['PolicyId'],
  csv: { numeric: 'auto', strictDialect: true },
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
| `cellLedger` | `'differences'` | scope of `cells.xlsx`: `'differences'` \| `'all'` \| `'none'` |
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
| `strictSheets` | `false` | treat added and unconfigured sheets as failures |

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
npm test            # 122 Playwright tests
npm run typecheck
npm run example     # run the six example cases
```
