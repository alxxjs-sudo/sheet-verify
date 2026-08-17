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

## Install

```bash
npm install exceljs csv-parse jszip
```

Then copy `src/` into your test project, or build and depend on it:

```bash
npm install && npm run build
```

Playwright transpiles TypeScript itself, so `src/` can be imported directly from
test files without a build step.

## Use in Playwright

```ts
import { test } from '@playwright/test';
import { expect } from 'sheet-verify/matcher';

test('monthly policy export matches baseline', async () => {
  const actual = await app.generateReport('2026-08');

  await expect(actual).toMatchSheetBaseline('baselines/policies-2026-08.xlsx', {
    keyColumns: ['PolicyId'],
    tolerance: { '*': 0, 'Annual Cost': 0.01, Commission: 0.01 },
    ignoreColumns: ['Generated At'],
  });
});
```

The full diff is attached to the test result as `sheet-diff.txt` and
`sheet-diff.json`, so it lands in the Playwright HTML report and in Allure.

## Cases

A case is a folder holding everything about one report type: the golden output
the new report is judged against, the report from the latest run, and the
artefacts describing what the comparison did. Everything for one report sits
together, so it can be reviewed, archived or attached to a ticket as a unit.

```
cases/monthly-policy-export/
  golden.xlsx     committed; the contract
  actual.xlsx     the latest run, copied in
  diff.txt        the differences, human-readable
  diff.json       the same, structured
  cells.xlsx      one row per differing cell, formatted for working through
  compared.xlsx   every cell checked, one worksheet per compared table
```

```ts
await expect(actual).toMatchCase('cases/monthly-policy-export', {
  sheets: {
    Policies: { keyColumns: ['PolicyId'] },
    Premiums: { keyColumns: ['PolicyId', 'Period'] },
  },
});
```

The new report is copied into the folder whatever the outcome, so a CI failure
can be opened next to the golden output it was judged against rather than
hunting for it in a build artefact. A missing `golden.xlsx` is created from the
new report and the run passes, the way a snapshot test behaves on first use —
commit it and review the contents.

Outside Playwright, `runCase()` does the same and returns the result:

```ts
import { runCase } from 'sheet-verify';

const result = await runCase(actual, 'cases/monthly-policy-export', spec);
if (!result.ok) console.error(result.summary, result.files.diffText);
```

### cells.xlsx

`diff.txt` reads top-down; `cells.xlsx` is for working through cell by cell. It
arrives as a real Excel table — filter buttons, banded rows, header frozen,
columns already wide enough to read, and the status column colour-coded — so it
opens ready to sort and filter rather than needing to be set up first.

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

The format follows the file name — `.xlsx` gets the formatted table, `.csv` gets
streamed text. `cellLedger: 'none'` skips the file entirely, and `'all'` folds
every matching cell into this file too rather than leaving them to
`compared.xlsx`.

### compared.xlsx

The full record of what was checked, in a deliberately plain column set — row
key, column, the cell address on each side, both values, and the verdict.
"We compared this and it was fine" is the claim a golden-file suite is really
making, and this is the file that states it.

It is **split one worksheet per compared table**, because this is the file that
grows with the report:

```
Policies · Info      10 cells
Policies · Detail    35 cells
Premiums · Info      10 cells
Premiums · Detail    50 cells
…
```

A five-sheet report becomes ten tabs to scan rather than one sheet of
everything, and each tab stays clear of Excel's million-row ceiling on its own.
A table that would exceed it anyway is truncated with a row saying so.

Turn it off for a case where it is not worth the write time:

```ts
await expect(actual).toMatchCase(dir, { ...spec, comparedLedger: false });
```

### Multi-sheet workbooks

Most generated workbooks have several sheets with little in common, so each one
gets its own spec and `defaults` covers what they share. The whole workbook is
compared in a single pass — the file is parsed once, not once per sheet.

```ts
await expect(actual).toMatchWorkbookBaseline('baselines/policies-2026-08.xlsx', {
  defaults: { tolerance: { '*': 0.01 }, ignoreColumns: ['Generated At'] },
  sheets: {
    Policies: { keyColumns: ['PolicyId'] },
    Premiums: { keyColumns: ['PolicyId', 'Period'] },
    Regions:  { keyColumns: ['Region'], headerRow: 2 },
  },
  ignoreSheets: ['Scratch'],
});
```

`defaults` merges per field: `tolerance` records merge, `ignoreColumns` and
`invariants` accumulate, everything else is replaced by the per-sheet value.
Sheet names match case-insensitively, and a name in `sheets` that matches no
sheet in either file is reported as an error rather than ignored — it is
almost always a typo.

**The baseline's sheet list is the contract:**

| sheet is in | verdict |
| --- | --- |
| both | compared, all four layers |
| the new output only | **noted, not compared** — there is nothing to compare it against |
| the baseline only | **defect** — output that used to be produced is gone |
| both, but no `keyColumns` | listed as *not compared*, so the gap is visible |

The asymmetry is deliberate. A new sheet is additive and harmless; a sheet that
has vanished means consumers stopped receiving something they expect. Set
`strictSheets: true` to fail on added and unconfigured sheets too.

One consolidated report covers the workbook: sheets that failed appear first
with their full diff, and everything merely worth reviewing collapses to a line
each at the end.

```
baseline  policies-2026-07.xlsx  4 sheet(s)
actual    policies-2026-08.xlsx  4 sheet(s)
          2 compared · 1 added · 1 removed · 1 not compared

SHEETS REMOVED (1) — output that is no longer produced
  ✗ − sheet "Legacy"

SHEET "Premiums" — 1 value
  VALUE CHANGES (1 root cause, 1 cascaded)
    P-1003 / 2026-08 · Amount @C7: 6000 → 9999  (Δ 3999)

SHEETS TO REVIEW
  ~ sheet "Policies" — 3 schema
  ? sheet "Regions" — no keyColumns configured for this sheet
  + sheet "Premium Detail" — new, not compared
```

CSV has no worksheets, so it stays on `toMatchSheetBaseline`.

### Several tables on one sheet

Most generated sheets carry a small "output info" block above the data — report
name, creator, release, a generation timestamp. That is two tables on one sheet,
and a single `headerRow` cannot describe it: reading from row 1 runs to the
bottom and swallows the data table below.

Declare each table instead. They are bounded by the next table's `headerRow`, so
nothing needs re-counting as the data grows:

```ts
const info = {
  headerRow: 1,
  keyColumns: ['Field'],          // a key-value block is keyed by field name
  ignoreRows: ['Generated At'],   // rewritten every run
};

await expect(actual).toMatchWorkbookBaseline(baseline, {
  defaults: { headerRow: 8 },
  sheets: {
    Policies: { tables: { Info: info, Detail: { keyColumns: ['PolicyId'] } } },
    Premiums: { tables: { Info: info, Detail: { keyColumns: ['PolicyId', 'Period'] } } },
  },
});
```

Each table is compared independently and reported under its own name, so a
release bump in the info block never mixes with a defect in the data:

```
SHEET "Premiums · Info" — 1 value
    Release · Value @B5: "4.2.0" → "4.3.0"

SHEET "Premiums · Detail" — 1 value
    P-1003 / 2026-08 · Gross @C14: 6000 → 6900  (Δ 900)
```

`ignoreRows` is the row-wise counterpart of `ignoreColumns`. In a key-value
block the per-run timestamp is a *row*, so no column exclusion can reach it.

A runnable version of exactly this — a five-sheet report with an info block and
a formula table on every sheet, and a following release that inserts a column,
adds a sheet, moves rows and plants two real defects — is in [examples/](examples/).
It runs as a case and leaves the folder behind to inspect:

```bash
npm run build && npm run example
```

### Re-blessing a baseline

Schema changes between releases are expected. When a diff is correct, accept it
explicitly rather than editing the baseline by hand:

```bash
UPDATE_SHEET_BASELINE=1 npx playwright test
```

Baselines are committed to git, so the change is reviewed in the pull request.
A missing baseline is created on first run, the way snapshot tests behave.

## Use directly

```ts
import { verifySheet, verifyWorkbook, formatReport, formatWorkbookReport } from 'sheet-verify';

const diff = await verifySheet('baseline.xlsx', 'actual.xlsx', {
  keyColumns: ['PolicyId'],
});
console.log(formatReport(diff));

const wb = await verifyWorkbook('baseline.xlsx', 'actual.xlsx', {
  sheets: { Policies: { keyColumns: ['PolicyId'] } },
});
console.log(formatWorkbookReport(wb));
```

## What it reports

Independent layers, so a schema change is never confused with a defect:

| layer | meaning | verdict |
| --- | --- | --- |
| Sheets | worksheets added, removed, moved | review, except removed |
| Schema | columns added, removed, moved | review |
| Row population | keys present in one file only | review |
| Values | cell differences, with per-column tolerance | defect |
| Formulas | calculation logic, normalised | defect |

A renamed column is reported as one added plus one removed, not as a rename —
nothing here guesses that `Cost` became `Annual Cost`.

Value differences are split into **root causes** and **cascades**: one wrong
input that feeds two formulas is reported as one cause, not three failures.

## Spec options

Case-level options for `toMatchCase` / `runCase`, on top of the workbook options:

| option | default | meaning |
| --- | --- | --- |
| `cellLedger` | `'differences'` | scope of `cells.xlsx`: `'differences'` \| `'all'` \| `'none'` |
| `comparedLedger` | `true` | write `compared.xlsx` at all |
| `names` | see above | file names within the case folder |
| `updateGolden` | `false` | overwrite the golden output and pass |
| `createMissingGolden` | `true` | create it on first run rather than failing |

Workbook-level options for `toMatchWorkbookBaseline` / `verifyWorkbook`:

| option | default | meaning |
| --- | --- | --- |
| `sheets` | `{}` | per-sheet spec, keyed by worksheet name |
| `defaults` | `{}` | applied to every sheet, then overridden per sheet |
| `ignoreSheets` | `[]` | worksheets excluded entirely |
| `strictSheets` | `false` | treat added and unconfigured sheets as failures |

Per-sheet options, and the whole spec for a single-sheet comparison:

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
| `formulaMode` | `'header'` | `'header'` \| `'r1c1'` \| `'a1'` |
| `compareFormulas` | `true` | compare formula logic at all |
| `requireCachedValues` | `true` | fail if formula cells have no cached result |
| `looseHeaders` | `true` | match headers ignoring case and extra spaces |
| `strictSchema` | `false` | treat schema changes as failures |
| `invariants` | `[]` | baseline-free assertions |
| `csv.numeric` | `'auto'` | how CSV text becomes numbers |
| `csv.strictDialect` | `false` | fail on BOM/delimiter/line-ending drift |

## Invariants

A baseline proves the new output matches the old one, never that either is
correct. Invariants close that gap by asserting properties of the output alone:

```ts
import { invariants as inv } from 'sheet-verify';

await expect(actual).toMatchSheetBaseline(baseline, {
  keyColumns: ['PolicyId'],
  invariants: [
    inv.noErrorValues(),                  // #REF!, #DIV/0! anywhere
    inv.notBlank('PolicyId', 'Holder'),
    inv.unique('PolicyId'),
    inv.inRange('Rate', 0, 1),
    inv.derived('Commission', (_row, num) => num('Annual Cost') * 0.1, 0.005),
  ],
});
```

`derived()` recomputes a column independently, so it catches calculation bugs
that are present in the baseline too.

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
settings are lost. Verify against *your* templates rather than trusting this.

If a template does fail the gate, do not rewrite it. Write only into a data
sheet the template's own formulas read from, so the fragile parts are never
touched.

## Notes and limits

- **Cached values.** Nothing here evaluates formulas. Comparison uses the value
  the generator wrote plus the formula text. If the generator writes formulas
  with no cached result, value comparison has nothing to check — hence
  `requireCachedValues`, on by default.
- **Cross-sheet formulas.** References to other sheets stay in A1 form, since
  this sheet's headers say nothing about another sheet's layout. A column moved
  on a *referenced* sheet will show as a formula difference. Resolving those to
  header names is possible now that the workbook layer loads every sheet, but
  it is not implemented — nothing needed it yet.
- **Duplicate keys.** Rows with a repeated key are excluded and reported as an
  integrity error rather than guessed at.
- **Memory.** Around 600 MB of heap for 50k rows. Past ~100k rows use a
  streaming reader or the CSV path.
- **ExcelJS is dormant.** No release since October 2023. It is still the most
  capable option, which is why the reader sits behind the `SheetReader`
  interface — `registerReader()` swaps it without touching comparison logic.
- **Do not `npm install xlsx`.** The npm copy is stuck at 0.18.5 and carries two
  high-severity advisories whose fixes exist only on SheetJS's own CDN.
