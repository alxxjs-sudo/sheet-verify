# Using it as a library

The CLI is one caller of this. Everything below is the programmatic surface:
matchers for a test suite, the direct API, invariants, and how to extend it.

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
