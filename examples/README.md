# Examples

Six cases covering the outcomes you will actually meet, each one runnable:

```bash
npm run build && npm run example
```

Every case generates its own golden output and report, runs the comparison, and
leaves a folder behind in `examples/cases/<name>/` to open and inspect. Nothing
here is committed — it is all regenerated, so the generators are the source of
truth rather than stale binaries in the history.

## The cases

| case | outcome | what it shows |
| --- | --- | --- |
| `monthly-policy-export` | FAIL | the full picture: a column inserted, a sheet added, rows churned, and two real defects underneath |
| `clean-release` | PASS | nothing changed — `identical`, and no `differences.xlsx` written |
| `schema-drift-only` | PASS | a `Premium` column inserted and nothing else: reported once, not as churn across every row |
| `sheet-removed` | FAIL | the `Regions` sheet stops being produced — a defect with **no cells to point at** |
| `uncached-formulas` | FAIL | the generator writes formulas with no cached result, so value comparison has nothing to check |
| `invariant-catch` | FAIL | a rate of 1.4 in **both** files: the comparison passes, and only an invariant catches it |

Two of them are worth dwelling on.

**`sheet-removed`** fails without writing a `differences.xlsx` at all. A sheet
that is no longer produced has no cells to compare, so the failure appears in
`diff.txt` and `diff.json` and nowhere else. It is the clearest demonstration of
why those two files are not redundant with the cell ledgers — and why a missing
`differences.xlsx` never means a case passed.

**`invariant-catch`** is the case a golden-file suite cannot catch on its own.
The rate is wrong in the golden output *and* the new report, so the comparison is
perfectly happy. Only `inv.inRange('Rate', 0, 1)` fails it. Any suite that relies
on a baseline alone has this blind spot.

## The report being generated

Five sheets, each carrying **two tables** — an output-info block on rows 1–6 and
the data table from row 8 down. That shape is why `tables` exists: a single
`headerRow` cannot describe it, and reading from row 1 would swallow the data
table below.

| sheet | key | formula |
| --- | --- | --- |
| Policies | `PolicyId` | Annual Cost = Sum Insured × Rate |
| Premiums | `PolicyId` + `Period` | Tax, Net |
| Claims | `ClaimId` | Net Claim = Gross − Excess |
| Commissions | `AgentId` | Commission = Volume × Rate |
| Regions | `Region` | Loss Ratio = Claims ÷ Premiums |

Each info block holds a `Generated At` timestamp that changes on every run. It is
excluded with `ignoreRows: ['Generated At']` — the row-wise counterpart of
`ignoreColumns`, since in a key-value block the per-run value is a row.

## When you need a case.json

A separate, narrated walkthrough:

```bash
npm run build && npm run example:case-json
```

It builds a report with a `Summary` sheet whose rows are identified by **three**
columns — `Region` + `Band` + `Quarter`. Detection tries single columns and then
pairs, so it finds no key there, and a defect is planted on that sheet. Five
steps, driving the real CLI at each one:

1. two files in a case folder
2. run it — **passes**, reporting `1 sheet not compared`, defect unreported
3. `--print-spec` shows `Summary · Table 2` has a `headerRow` and no `keyColumns`
4. write a three-line `case.json` naming the key
5. run again — `12 tables compared`, and the defect surfaces:
   `Varna / B / Q2 · Premium @D16: 88000 → 91500  (Δ 3500)`

The point of it is step 2. A green run with `not compared` in the summary is not
a clean bill of health, and this is the cheapest way to see why.

It also shows a subtlety worth knowing: the `case.json` names the key on the
**sheet**, not on `Table 2`. A per-table value beats a sheet-level one, so the
info block keeps the `Field` key detection gave it, and only the table that had
none picks the new key up.

The folder is left in `examples/cli-cases/` with its `case.json` to read.

## The files

- **`generate-report.ts`** — builds the workbook. Takes a `Variant` describing
  what to vary: insert a column, plant a defect, drop a sheet, omit cached
  values, and so on. Also runnable on its own:

  ```bash
  node --experimental-strip-types examples/generate-report.ts out.xlsx
  node --experimental-strip-types examples/generate-report.ts out.xlsx --next
  ```

- **`cases.ts`** — the six scenarios, each pairing a golden variant with a report
  variant and the expected verdict.

- **`run-cases.ts`** — runs them all and prints what each produced. Exits
  non-zero if any case stops behaving as documented, so the examples cannot rot
  silently.

## Adapting one to your own report

Copy `cases.ts` as a starting point and change three things:

1. **`sheets`** — one entry per worksheet, with the `keyColumns` that identify a
   row on it.
2. **`tables`** — only if a sheet holds more than one table. Drop it entirely if
   your sheets have a single header row.
3. **`ignoreRows` / `ignoreColumns`** — whatever your generator rewrites on every
   run. Timestamps and run ids are the usual ones.

Then point `runCase` at a real generated file instead of `generate()`.
