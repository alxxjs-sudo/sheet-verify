# Describing your report

Every option, and where each one goes. Reach for this when detection needs
correcting, or when you are calling the API directly. For working out *which*
correction a report needs, see [detection-tuning.md](detection-tuning.md).

## Keys

`keyColumns` is the only required option when calling the API. One column, or
several forming a composite key:

```ts
{ keyColumns: ['PolicyId'] }
{ keyColumns: ['PolicyId', 'Period'] }
```

Rows whose key is blank are skipped. Rows with a **repeated** key are excluded
and reported as an integrity error rather than guessed at.

## Multi-sheet workbooks

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
| both, but no `keyColumns` | **not compared** — see [Detection will not invent a row key](../README.md#detection-will-not-invent-a-row-key); the gap is reported, never guessed at |

The asymmetry is deliberate. A new sheet is additive and harmless; a sheet that
has vanished means consumers stopped receiving something they expect.

## Several tables on one sheet

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

**Overriding one table** is per table and per field, so you write only what you
are changing and detection's header row, end row and key all survive:

```json
{ "sheets": { "Currency Info": { "tables": { "Table 3": { "tolerance": 5 } } } } }
```

**Adding a table** is the same thing with a name detection did not produce — it
is added rather than replacing anything, and works in a `case.json` or a report
type's `meta.json` alike. Adding one to a sheet that holds a single table keeps
both: the sheet's own table stays, under the sheet's name.

Which means **a name that matches nothing is not a no-op**. It adds a phantom
table at `A1:B1` with no rows and the run fails on it — the usual cause of an
integrity error appearing right after a config edit. Check the name against
`--print-spec` first.

[docs/detection-tuning.md](detection-tuning.md) covers both in full,
including how an added table is bounded and when it needs `columns`.

## Tolerance and exclusions

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

**The default is `0.001`**, applied when no config names one. It absorbs what
recalculation leaves behind — a total rebuilt in a different order lands about
`1.19e-7` away from the one stored last month on a figure in the hundreds of
millions — while staying well under a cent, so anything written down on purpose
is still reported. Set `"tolerance": 0` in a `meta.json` for exactness, or a
larger number where a column deserves one.

It is **absolute, not a percentage**. On a figure like 164,488,104.6, a
tolerance of `0.01` means "ignore anything under a cent". Bear the scale of the
column in mind: the same `0.01` on a rate column holding `0.05` is a fifth of
the value.

**Both layers apply it.** Layer 1 looks the tolerance up by column name; layer 2
works in addresses, so it reads the same per-column tolerances off the tables
layer 1 compared, and falls back to the `*` entry from `defaults` for cells no
compared table covers — title blocks, unkeyed tables, anywhere else it reaches.
Without that, a tolerance would quiet the keyed comparison while the headline
`cells differing` count went on reporting the same noise.

Cells inside tolerance are **counted separately and still listed**, never
dropped:

```markdown
**Cells that differ**

| total | within tolerance (±0.001) | above tolerance |
| ---: | ---: | ---: |
| 894 | 870 | **24 (2.7%)** |

## Inside the tolerance you set (870)
```

That way a tolerance set too wide is visible rather than silent — you can read
what it swallowed and see the size of each gap it forgave. In
`differences.xlsx` the same cells carry the status `within-tolerance`, so the
Status column filters them out or in.

One caveat, stated because it is easy to trip over: layer 2 never sees a cell's
type, only the text it displays, so text that reads as a number counts as one
there. A version written `4.20` would be compared numerically. Set a tolerance
to the size of the rounding it is meant to absorb rather than to a round figure,
and this never bites.

## Report metadata

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

## Options reference

**Case options** — `toMatchCase` / `runCase`, on top of the workbook options:

| option | default | meaning |
| --- | --- | --- |
| `cellLedger` | `'differences'` | scope of `differences.xlsx`: `'differences'` \| `'all'` \| `'none'` |
| `comparedLedger` | `true` | write `compared.xlsx` at all |
| `names` | see [Cases](reading-a-report.md) | file names within the case folder |
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
| `expect` | any `meta.json` or `case.json` | tables each sheet should have compared: ranges, or a count |
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
| `headerRow` | `1` | 1-based row holding headers; really "the row above the data", so `0` means there is no header row and the table starts at row 1 |
| `endRow` | *last row* | 1-based last row; set for you when a sheet declares `tables` |
| `columns` | *every column* | column range this table occupies, `"H:J"`; set for you when a sheet holds tables side by side |
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
