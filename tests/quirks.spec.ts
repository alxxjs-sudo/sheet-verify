import { test, expect } from '@playwright/test';
import { mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import {
  detectWorkbook, formatMarkdownReport, ledgerRows, runWorkbook, sweep,
} from '../src/index.js';
import type { WorkbookSpec } from '../src/index.js';
import { DIR } from './fixtures.js';

/**
 * Shapes real reports turn out to have, each of which made the comparison
 * report something that had not changed. Every case here was found on an
 * actual report before it was written down as a test.
 */

interface CellSpec {
  formula?: string;
  value?: string | number;
}
type Row = (string | number | null | CellSpec)[];

/** Builds a one-sheet workbook from a grid, with optional merged ranges. */
async function sheet(
  name: string,
  rows: Row[],
  opts: { merges?: string[]; sheetName?: string; headerRow?: number } = {},
): Promise<string> {
  await mkdir(DIR, { recursive: true });
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(opts.sheetName ?? 'Data');
  for (const row of rows) {
    ws.addRow(row.map((c) =>
      c !== null && typeof c === 'object' ? (c.formula ? { formula: c.formula } : c.value) : c));
  }
  for (const m of opts.merges ?? []) ws.mergeCells(m);
  // Paint a row the way these generators paint a heading: bold, on navy.
  if (opts.headerRow) {
    ws.getRow(opts.headerRow).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000F47' } };
    });
  }
  const full = join(DIR, name);
  await wb.xlsx.writeFile(full);
  return full;
}

const headersOf = async (path: string) =>
  (await detectWorkbook(path))[0]!.tables.map((t) => ({ row: t.headerRow, headers: t.headers.filter(Boolean) }));

test.describe('detection: which row is the header', () => {
  test('a merged banner is one cell wide, not nineteen headers', async () => {
    // Row 1 is a category band merged across the table. ExcelJS reports the
    // master's value from every covered cell, so before this it looked like
    // the widest row on the sheet and every column was named "PERSONAL LINES".
    const path = await sheet('quirk-banner.xlsx', [
      ['PERSONAL LINES', null, null, null],
      ['Region', 'Premium', 'Claims', 'Ratio'],
      ['North', 100, 40, 0.4],
      ['South', 200, 90, 0.45],
    ], { merges: ['A1:D1'] });

    const [table] = await headersOf(path);

    expect(table!.row).toBe(2);
    expect(table!.headers).toEqual(['Region', 'Premium', 'Claims', 'Ratio']);
  });

  test('a row of numbers is data however wide it is', async () => {
    // Without this the widest row wins and the columns get named after
    // whatever numbers happened to be in it -- "8190608765.87" as a column.
    const path = await sheet('quirk-numeric.xlsx', [
      ['Portfolio Totals', null, null],
      ['Region', 'Premium', 'Claims'],
      [8190608765.87, 420030849, 10547647459.72],
      [8190608765.87, 420030849, 10547647459.72],
    ]);

    const [table] = await headersOf(path);

    expect(table!.row).toBe(2);
    expect(table!.headers).toEqual(['Region', 'Premium', 'Claims']);
  });

  test('a painted header beats a fuller data row beneath it', async () => {
    // The info-block shape: the heading holds two words, the row under it
    // holds three. On cell counts alone the data row wins, and the columns
    // end up named "All Lines HU" and "Label A" -- values, not headings.
    const path = await sheet('quirk-styled.xlsx', [
      ['Portfolio', 'Model'],
      ['All Lines HU', 'Label A', 'Verisk Touchstone 8.0'],
      ['Label B', 'RMS RiskLink 18.1', 'EQ'],
    ], { headerRow: 1 });

    const [table] = await headersOf(path);

    expect(table!.row).toBe(1);
    expect(table!.headers).toEqual(['Portfolio', 'Model']);
  });

  test('but paint alone does not make a row of numbers a header', async () => {
    const path = await sheet('quirk-styled-numeric.xlsx', [
      ['Region', 'Premium', 'Claims'],
      [100, 200, 300],
      [400, 500, 600],
    ], { headerRow: 2 });

    const [table] = await headersOf(path);

    expect(table!.row).toBe(1);
    expect(table!.headers).toEqual(['Region', 'Premium', 'Claims']);
  });

  test('a key is never taken from a column of formulas', async () => {
    // Detection reads a formula cell as the text it would produce, so a column
    // of them looks full. The comparison only ever sees stored values, and
    // there are none -- so a key taken from that column gives every row a
    // blank key and the whole table is dropped. One real sheet detected such a
    // key and then compared none of its 888 rows.
    const path = await sheet('quirk-formulakey.xlsx', [
      ['Ref', 'Region', 'Amount'],
      [{ formula: 'CONCATENATE("R-",1)' }, 'North', 10],
      [{ formula: 'CONCATENATE("R-",2)' }, 'South', 20],
      [{ formula: 'CONCATENATE("R-",3)' }, 'East', 30],
    ], { headerRow: 1 });

    const [table] = await headersOf(path);
    expect(table!.headers).toEqual(['Ref', 'Region', 'Amount']);

    const detected = (await detectWorkbook(path))[0]!.tables[0]!;
    // "Ref" is identifier-shaped and looks unique, but holds no stored value.
    expect(detected.keyColumns).not.toContain('Ref');
  });

  test('a totals row of uncomputed formulas is not a row of twenty-five names', async () => {
    // The header row is row 1. Row 2 totals it with formulas the generator
    // never computed -- which is every formula in these reports. Those used to
    // contribute their own text as a name, so row 2 looked like three names
    // against row 1's three, won the search on paint, and the columns ended up
    // called SUM(B3:B4). That name carries row numbers, so inserting a row
    // renames every column and the whole table reports as columns added and
    // removed.
    await mkdir(DIR, { recursive: true });
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Loss');
    ws.addRow(['Company', 'Single Limit', 'Aggregate Limit']);
    ws.getRow(1).font = { bold: true };
    ws.addRow(['Total', { formula: 'SUM(B3:B4)' }, { formula: 'SUM(C3:C4)' }]);
    ws.getRow(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDDDDD' } };
    ws.addRow(['Farmers', 100, 200]);
    ws.addRow(['Hartford', 300, 400]);

    const path = join(DIR, 'quirk-formula-totals.xlsx');
    await wb.xlsx.writeFile(path);
    const [table] = await headersOf(path);

    expect(table!.row).toBe(1);
    expect(table!.headers).toEqual(['Company', 'Single Limit', 'Aggregate Limit']);
  });

  test('a block of nothing but numbers has no header row, and keeps every row', async () => {
    // This used to take row 1 as the header on the grounds that a wrong header
    // beats refusing to compare. It does -- but naming the columns 1, 2, 3
    // costs the first row of data and gives names that move the moment a
    // figure changes, which is how a whole table came to be reported as
    // columns added and removed. Positional names cost neither.
    const path = await sheet('quirk-allnumeric.xlsx', [
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9],
    ]);

    const [table] = await headersOf(path);

    expect(table!.row).toBe(0);
    expect(table!.headers).toEqual(['Column A', 'Column B', 'Column C']);
  });
});

test.describe('comparison: differences that are not differences', () => {
  const SPEC: WorkbookSpec = {
    defaults: { requireCachedValues: false },
    sheets: { Data: { keyColumns: ['Region'] } },
  };
  const diffOf = async (a: string, b: string, spec: WorkbookSpec = SPEC) =>
    (await runWorkbook(a, b, spec)).diff.sheets.find((s) => s.sheet === 'Data')!.diff!;

  test('whitespace around a formula is not a calculation change', async () => {
    // Excel strips it on save, so a generator that wrote " IF(...)" comes back
    // as "IF(...)" from any file a person has opened.
    const golden = await sheet('quirk-ws-g.xlsx', [
      ['Region', 'Premium', 'Doubled'],
      ['North', 100, { formula: ' B2*2' }],
    ]);
    const actual = await sheet('quirk-ws-a.xlsx', [
      ['Region', 'Premium', 'Doubled'],
      ['North', 100, { formula: 'B2*2' }],
    ]);

    const d = await diffOf(golden, actual);

    expect(d.formulas).toHaveLength(0);
    expect(d.ok).toBe(true);
  });

  test('a real formula change is still caught', async () => {
    const golden = await sheet('quirk-ws-g2.xlsx', [
      ['Region', 'Premium', 'Doubled'],
      ['North', 100, { formula: 'B2*2' }],
    ]);
    const actual = await sheet('quirk-ws-a2.xlsx', [
      ['Region', 'Premium', 'Doubled'],
      ['North', 100, { formula: 'B2*3' }],
    ]);

    expect((await diffOf(golden, actual)).formulas).toHaveLength(1);
  });

  test('with tolerance off, the fifteenth digit is still a difference', async () => {
    // There is no hidden slack in the comparison: asked for exactness, it
    // compares exactly, however far down the gap sits. What absorbs a gap is
    // `tolerance` and nothing else -- and since it now defaults to 0.001, a
    // test about exactness has to say so rather than assume it.
    const golden = await sheet('quirk-float-g.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.45781166297475],
    ]);
    const actual = await sheet('quirk-float-a.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.45781166297476],
    ]);

    const d = (await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false, tolerance: 0 },
      sheets: { Data: { keyColumns: ['Region'] } },
    })).diff.sheets.find((x) => x.sheet === 'Data')!.diff!;

    expect(d.values).toHaveLength(1);
    expect(d.ok).toBe(false);
  });

  test('and the default tolerance absorbs exactly that, visibly', async () => {
    // The same pair with nothing configured. The gap is 1e-14, far inside the
    // 0.001 default, so it is not a difference -- and the run says how many
    // cells it forgave rather than quietly dropping them.
    const golden = await sheet('quirk-float-g5.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.45781166297475],
    ]);
    const actual = await sheet('quirk-float-a5.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.45781166297476],
    ]);

    const d = await diffOf(golden, actual);

    expect(d.values).toHaveLength(0);
    expect(d.ok).toBe(true);
  });

  test('and a tolerance is how you say it does not matter', async () => {
    const golden = await sheet('quirk-float-g4.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.45781166297475],
    ]);
    const actual = await sheet('quirk-float-a4.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.45781166297476],
    ]);

    const d = (await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false, tolerance: { Loss: 1e-9 } },
      sheets: { Data: { keyColumns: ['Region'] } },
    })).diff.sheets.find((x) => x.sheet === 'Data')!.diff!;

    expect(d.values).toHaveLength(0);
    expect(d.ok).toBe(true);
  });

  test('a difference that matters is still caught, however small', async () => {
    const golden = await sheet('quirk-float-g2.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.45],
    ]);
    const actual = await sheet('quirk-float-a2.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.46],
    ]);

    expect((await diffOf(golden, actual)).values).toHaveLength(1);
  });
});

test.describe('a grouping column written once per group', () => {
  // The shape every one of these reports uses: the portfolio name sits on the
  // first row of its group and the rows beneath leave it blank.
  const GROUPED: Row[] = [
    ['Portfolio', 'Event ID', 'Loss'],
    ['Alpha', 'E-1', 10],
    [null, 'E-2', 20],
    [null, 'E-3', 30],
    ['Beta', 'E-1', 40],
    [null, 'E-2', 50],
    [null, 'E-3', 60],
  ];

  const spec = (fillKeyDown: boolean): WorkbookSpec => ({
    defaults: { requireCachedValues: false },
    sheets: { Data: { keyColumns: ['Portfolio', 'Event ID'], fillKeyDown } },
  });

  test('collides across groups when the heading is read literally', async () => {
    const golden = await sheet('quirk-group-g.xlsx', GROUPED);
    const actual = await sheet('quirk-group-a.xlsx', GROUPED);

    const d = (await runWorkbook(golden, actual, spec(false))).diff.sheets
      .find((s) => s.sheet === 'Data')!.diff!;

    // "|E-1" for Alpha's row and Beta's row alike: the same key twice.
    expect(d.rows.duplicateKeysBase.length).toBeGreaterThan(0);
  });

  test('and identifies every row once the heading is carried down', async () => {
    const golden = await sheet('quirk-group-g2.xlsx', GROUPED);
    const actual = await sheet('quirk-group-a2.xlsx', GROUPED);

    const d = (await runWorkbook(golden, actual, spec(true))).diff.sheets
      .find((s) => s.sheet === 'Data')!.diff!;

    expect(d.rows.duplicateKeysBase).toEqual([]);
    expect(d.rows.compared).toBe(6);
  });

  test('the carried value is used for the key only, never compared as data', async () => {
    const golden = await sheet('quirk-group-g3.xlsx', GROUPED);
    // Beta's second row gains a value; nothing about the grouping changed.
    const edited: Row[] = GROUPED.map((r, i) => (i === 5 ? ['Beta', 'E-2', 55] : r));
    const actual = await sheet('quirk-group-a3.xlsx', edited);

    const d = (await runWorkbook(golden, actual, spec(true))).diff.sheets
      .find((s) => s.sheet === 'Data')!.diff!;

    // The Loss change is reported. Portfolio is NOT, even though the golden
    // holds a blank there and the actual holds "Beta" -- filling is for keys.
    expect(d.values.map((v) => v.column)).toEqual(['Portfolio', 'Loss']);
    expect(d.values.find((v) => v.column === 'Loss')!.next).toBe(55);
  });

  test('a spacer row does not inherit the group above it', async () => {
    const withGap: Row[] = [...GROUPED.slice(0, 4), [null, null, null], ...GROUPED.slice(4)];
    const golden = await sheet('quirk-group-g4.xlsx', withGap);
    const actual = await sheet('quirk-group-a4.xlsx', withGap);

    const d = (await runWorkbook(golden, actual, spec(true))).diff.sheets
      .find((s) => s.sheet === 'Data')!.diff!;

    // Six real rows; the blank one is skipped rather than taking Alpha's key.
    expect(d.rows.compared).toBe(6);
    expect(d.rows.duplicateKeysBase).toEqual([]);
  });
});

test.describe('comparison: repeated column names', () => {
  const SPEC: WorkbookSpec = {
    defaults: { requireCachedValues: false },
    sheets: { Data: { keyColumns: ['Program'] } },
  };

  test('a repeated column group is a normal report shape, not a defect', async () => {
    // Currency blocks list Name and Abbreviation over and over. Both files lay
    // them out identically, so pairing them by position is exact.
    const rows: Row[] = [
      ['Program', 'Currency Name', 'Currency Abbr', 'Currency Name', 'Currency Abbr'],
      ['Alpha', 'Euro', 'EUR', 'Dollar', 'USD'],
      ['Beta', 'Lev', 'BGN', 'Pound', 'GBP'],
    ];
    const golden = await sheet('quirk-dup-g.xlsx', rows);
    const actual = await sheet('quirk-dup-a.xlsx', rows);

    const d = (await runWorkbook(golden, actual, SPEC)).diff;

    expect(d.errors).toEqual([]);
    expect(d.sheets.find((s) => s.sheet === 'Data')!.diff!.errors).toEqual([]);
    expect(d.ok).toBe(true);
  });

  test('and the second of a repeated pair is still compared on its own', async () => {
    const golden = await sheet('quirk-dup-g2.xlsx', [
      ['Program', 'Currency Name', 'Currency Abbr', 'Currency Name', 'Currency Abbr'],
      ['Alpha', 'Euro', 'EUR', 'Dollar', 'USD'],
    ]);
    const actual = await sheet('quirk-dup-a2.xlsx', [
      ['Program', 'Currency Name', 'Currency Abbr', 'Currency Name', 'Currency Abbr'],
      ['Alpha', 'Euro', 'EUR', 'Lev', 'BGN'],
    ]);

    const d = (await runWorkbook(golden, actual, SPEC)).diff.sheets
      .find((s) => s.sheet === 'Data')!.diff!;

    // The change is in the *second* "Currency Name", which only works because
    // disambiguation is positional and consistent across the two files.
    expect(d.values).toHaveLength(2);
    expect(d.values.map((v) => v.column).sort()).toEqual(['Currency Abbr (#2)', 'Currency Name (#2)']);
  });

  test('but differing layouts make the pairing a guess, and that is an error', async () => {
    const golden = await sheet('quirk-dup-g3.xlsx', [
      ['Program', 'Currency Name', 'Currency Name'],
      ['Alpha', 'Euro', 'Dollar'],
    ]);
    const actual = await sheet('quirk-dup-a3.xlsx', [
      ['Program', 'Currency Name', 'Rate', 'Currency Name'],
      ['Alpha', 'Euro', 1.1, 'Dollar'],
    ]);

    const d = (await runWorkbook(golden, actual, SPEC)).diff.sheets
      .find((s) => s.sheet === 'Data')!.diff!;

    expect(d.errors.join(' ')).toContain('order their columns differently');
    expect(d.ok).toBe(false);
  });
});

test.describe('artefacts of the file rather than the report', () => {
  const POSITIONAL: WorkbookSpec = {
    defaults: { requireCachedValues: false },
    matchUnkeyedRowsByPosition: true,
  };

  test('trailing blank rows do not read as removed rows', async () => {
    // Excel's used range outlives its contents. A disclaimer page edited down
    // to four rows still reports thirty if something once occupied them, and
    // the two files have no reason to agree on that number -- so an identical
    // page arrived as seventeen "row-removed" entries.
    const golden = await sheet('quirk-tail-g.xlsx', [
      ['Notice'],
      ['These figures are indicative.'],
      ['Prepared for internal review.'],
      [null], [null], [null], [null], [null],
    ]);
    const actual = await sheet('quirk-tail-a.xlsx', [
      ['Notice'],
      ['These figures are indicative.'],
      ['Prepared for internal review.'],
    ]);

    const d = (await runWorkbook(golden, actual, POSITIONAL)).diff.sheets
      .find((s) => s.sheet === 'Data')!.diff!;

    expect(d.rows.removed).toEqual([]);
    expect(d.ok).toBe(true);
  });

  test('but a blank row between two populated ones still holds its place', async () => {
    // The reason trailing-only is the rule: an interior blank is what keeps
    // the rows beneath it lined up with the other side.
    const golden = await sheet('quirk-gap-g.xlsx', [
      ['Notice'], ['first'], [null], ['second'],
    ]);
    const actual = await sheet('quirk-gap-a.xlsx', [
      ['Notice'], ['first'], [null], ['CHANGED'],
    ]);

    const d = (await runWorkbook(golden, actual, POSITIONAL)).diff.sheets
      .find((s) => s.sheet === 'Data')!.diff!;

    expect(d.rows.removed).toEqual([]);
    expect(d.values).toHaveLength(1);
    expect(d.values[0]!.address).toBe('A4');
  });

  test('the cell ledger and the comparer agree, whatever the rule is', async () => {
    // The bug: the ledger carried its own equality rule, so report.md and
    // differences.xlsx disagreed about the same cell. What matters is that
    // they agree -- not which way they land, which `tolerance` decides.
    const golden = await sheet('quirk-led-g.xlsx', [
      ['Region', 'Total'],
      ['North', 84.76743932],
    ]);
    const actual = await sheet('quirk-led-a.xlsx', [
      ['Region', 'Total'],
      ['North', 84.76743932 + 1.42109e-14],
    ]);

    // Tolerance off, so the gap is a difference and there is something for the
    // two to agree about; the default would absorb it in both at once, which
    // is the same agreement seen from the other side.
    const run = await runWorkbook(golden, actual, {
      ...POSITIONAL,
      defaults: { requireCachedValues: false, tolerance: 0 },
    });
    const d = run.diff.sheets.find((s) => s.sheet === 'Data')!.diff!;
    const led = [...ledgerRows(run.compared, 'differences')]
      .filter((r) => r.status === 'value-differs');

    // Reported by both, at full precision, so the two sides read differently.
    expect(d.values).toHaveLength(1);
    expect(led).toHaveLength(1);
    expect(String(d.values[0]!.base)).not.toBe(String(d.values[0]!.next));

    // And a tolerance silences it in both at once.
    const quiet = await runWorkbook(golden, actual, {
      ...POSITIONAL,
      defaults: { requireCachedValues: false, tolerance: 1e-9 },
    });
    expect(quiet.diff.sheets.find((s) => s.sheet === 'Data')!.diff!.values).toHaveLength(0);
    expect([...ledgerRows(quiet.compared, 'differences')]
      .filter((r) => r.status === 'value-differs')).toHaveLength(0);
  });
});

test.describe('report shape', () => {
  test('what differs gets its own table, naming the tolerance that drew the line', async () => {
    // "894" reads as a disaster when 870 of those moved by less than a
    // thousandth, so the count is split into parts and given a table of its
    // own -- it is the figure the whole report exists to deliver.
    const golden = await sheet('quirk-parts-g.xlsx', [
      ['Region', 'Loss'],
      ['North', 100],
      ['South', 200],
      ['East', 300],
    ]);
    const actual = await sheet('quirk-parts-a.xlsx', [
      ['Region', 'Loss'],
      ['North', 100 + 1e-9],
      ['South', 200 + 1e-9],
      ['East', 305],
    ]);

    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      sheets: { Data: { keyColumns: ['Region'] } },
    };
    const run = await runWorkbook(golden, actual, spec);
    const swept = await sweep(golden, actual, run.compared);
    const md = formatMarkdownReport(run.diff, swept, { name: 'parts' });

    const lines = md.split(/\r?\n/);
    const head = lines.findIndex((l) => l.includes('**Cells that differ**'));
    expect(head).toBeGreaterThan(-1);

    // The tolerance is named on the column that used it, so nobody has to go
    // looking for which number drew the line.
    expect(lines[head + 2]).toBe('| total | within tolerance (±0.001) | above tolerance |');
    expect(lines[head + 4]).toBe('| 3 | 2 | **1 (33.3%)** |');

    // Three columns in the header and three in the row: a count rendered into
    // a broken table is a count nobody reads.
    const cols = (l: string) => l.split('|').slice(1, -1).length;
    expect(cols(lines[head + 2]!)).toBe(3);
    expect(cols(lines[head + 4]!)).toBe(3);
  });

  test('a changed table carries its own counts, and only when tolerance applied', async () => {
    // "Value changes (1)" reads very differently depending on whether the rest
    // of the table held still or drifted a hair each way.
    const golden = await sheet('quirk-tblcount-g.xlsx', [
      ['Region', 'Loss', 'Note'],
      ['North', 100, 'a'],
      ['South', 200, 'b'],
      ['East', 300, 'c'],
    ]);
    const actual = await sheet('quirk-tblcount-a.xlsx', [
      ['Region', 'Loss', 'Note'],
      ['North', 100 + 1e-9, 'a'],
      ['South', 200 + 1e-9, 'b'],
      ['East', 305, 'c'],
    ]);

    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      sheets: { Data: { keyColumns: ['Region'] } },
    };
    const run = await runWorkbook(golden, actual, spec);
    const swept = await sweep(golden, actual, run.compared);
    const md = formatMarkdownReport(run.diff, swept, { name: 'per-table' });

    const lines = md.split(/\r?\n/);
    const heading = lines.findIndex((l) => l.startsWith('### Data'));
    expect(heading).toBeGreaterThan(-1);

    // Where the table is, then its own counts, then the findings. The counts
    // are the table's own -- the one above the report is the whole file's.
    expect(lines[heading + 2]).toBe('_`A1:C4` — 3 columns × 3 rows, rows matched by key_');
    expect(lines[heading + 4]).toContain('within tolerance (±0.001)');
    expect(lines[heading + 6]).toBe('| 3 | 2 | **1 (33.3%)** |');
    expect(lines.slice(heading).find((l) => l.startsWith('**Value changes')))
      .toBe('**Value changes (1)**');
  });

  test('a table where nothing was tolerated still gets its counts', async () => {
    const golden = await sheet('quirk-tblnone-g.xlsx', [
      ['Region', 'Loss'],
      ['North', 100],
    ]);
    const actual = await sheet('quirk-tblnone-a.xlsx', [
      ['Region', 'Loss'],
      ['North', 105],
    ]);

    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      sheets: { Data: { keyColumns: ['Region'] } },
    };
    const run = await runWorkbook(golden, actual, spec);
    const swept = await sweep(golden, actual, run.compared);
    const md = formatMarkdownReport(run.diff, swept, { name: 'per-table-none' });

    const lines = md.split(/\r?\n/);
    const heading = lines.findIndex((l) => l.startsWith('### Data'));

    // The same three columns in the same place, so two tables can be compared
    // at a glance without working out whether a missing block means no drift
    // or no data. With nothing forgiven, the column claims no number.
    expect(lines[heading + 4]).toBe('| total | within tolerance | above tolerance |');
    expect(lines[heading + 6]).toBe('| 1 | 0 | **1 (100%)** |');
  });

  test('a table whose comparison never ran claims no counts', async () => {
    // Its key column is missing, so nothing was compared. Zeros here would say
    // the cells were checked and matched, which is the opposite of the
    // integrity error printed underneath.
    const golden = await sheet('quirk-tblnorun-g.xlsx', [
      ['Region', 'Loss'],
      ['North', 100],
    ]);
    const actual = await sheet('quirk-tblnorun-a.xlsx', [
      ['Region', 'Loss'],
      ['North', 105],
    ]);

    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      sheets: { Data: { keyColumns: ['Nope'] } },
    };
    const run = await runWorkbook(golden, actual, spec);
    const swept = await sweep(golden, actual, run.compared);
    const md = formatMarkdownReport(run.diff, swept, { name: 'per-table-norun' });

    const lines = md.split(/\r?\n/);
    const heading = lines.findIndex((l) => l.startsWith('### Data'));
    expect(lines[heading + 4]).toContain('Comparison integrity');
    expect(md).not.toContain('| 0 | 0 | **0 (0%)** |');
  });

  test('a clean verdict over a large sweep count says why it is not a contradiction', async () => {
    // The same rows in a different order. Layer 1 pairs by key and finds
    // nothing; layer 2 compares by address and finds everything. Seen on a real
    // 8,476-row CSV: verdict "Identical", and a table reading 50,274 differing,
    // 100% above tolerance, directly underneath it.
    const rows: Row[] = [['Key', 'A']];
    for (let i = 1; i <= 12; i++) rows.push([`K${i}`, i * 10]);
    const golden = await sheet('quirk-order-g.xlsx', rows);
    const shuffled: Row[] = [rows[0]!, ...rows.slice(1).reverse()];
    const actual = await sheet('quirk-order-a.xlsx', shuffled);

    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      sheets: { Data: { keyColumns: ['Key'] } },
    };
    const { diff } = await runWorkbook(golden, actual, spec);
    const swept = await sweep(golden, actual, [], { spec });
    const md = formatMarkdownReport(diff, swept, { name: 'order' });

    expect(diff.ok).toBe(true);
    expect(md).toContain('**Identical.**');
    // The count is still shown -- it is true -- but it no longer reads as the
    // report disagreeing with itself.
    expect(md).toContain('what moved *position*, not what changed');
    expect(md).toContain('Layer 2 never decides the verdict');
  });

  test('many failing tables are ranked, so the worst is found without scrolling', async () => {
    // Two sheets, one badly wrong and one barely. Flat, the reader learns
    // which is which by scrolling to the end of the first.
    const build = async (name: string, bump: (i: number, c: number) => number) => {
      await mkdir(DIR, { recursive: true });
      const wb = new ExcelJS.Workbook();
      for (const [sheetName, factor] of [['Big', 1], ['Small', 0]] as const) {
        const ws = wb.addWorksheet(sheetName);
        ws.addRow(['Key', 'A', 'B', 'C']);
        for (let i = 0; i < 8; i++) {
          ws.addRow([`K${i}`, ...[0, 1, 2].map((c) => (factor ? bump(i, c) : (i === 0 && c === 0 ? bump(i, c) : i)))]);
        }
      }
      const file = join(DIR, name);
      await wb.xlsx.writeFile(file);
      return file;
    };
    // Big differs in all 24 cells, Small in one.
    const golden = await build('quirk-rank-g.xlsx', (i) => i);
    const actual = await build('quirk-rank-a.xlsx', (i) => i + 1);

    const { diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false, tolerance: 0 },
      sheets: { Big: { keyColumns: ['Key'] }, Small: { keyColumns: ['Key'] } },
    });
    const md = formatMarkdownReport(diff, null, { name: 'rank' });

    expect(md).toContain('## Where the differences are');
    // The worst table is named before any of the detail, and above the lesser
    // one -- which is the whole point of the section.
    const summary = md.slice(md.indexOf('## Where the differences are'), md.indexOf('## What changed'));
    expect(summary.indexOf('Big')).toBeGreaterThan(-1);
    expect(summary.indexOf('Big')).toBeLessThan(summary.indexOf('Small'));
  });

  test('a wall of differences is summarised by column, capped, and listed in full on request', async () => {
    // A recalculated report drifts in the last digit of every total, so one
    // sheet can carry hundreds of differences of which two matter. Flat, the
    // two are unfindable.
    const rows: Row[] = [['Region', 'Gross', 'Net']];
    for (let i = 1; i <= 20; i++) rows.push([`R${i}`, 1000 + i, 500 + i]);
    const golden = await sheet('quirk-wall-g.xlsx', rows);

    const drifted: Row[] = [['Region', 'Gross', 'Net']];
    for (let i = 1; i <= 20; i++) {
      // Every Gross drifts by a hair; one of them moves by a million.
      const gross = i === 7 ? 1000 + i + 1_000_000 : 1000 + i + 1e-9;
      drifted.push([`R${i}`, gross, 500 + i]);
    }
    const actual = await sheet('quirk-wall-a.xlsx', drifted);

    // Exactness on purpose: this is a test about how a wall of differences is
    // presented, so the wall has to exist. Left to the default tolerance the
    // 1e-9 drifts would be absorbed and there would be one difference to show.
    const { diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false, tolerance: 0 },
      sheets: { Data: { keyColumns: ['Region'] } },
    });
    const md = formatMarkdownReport(diff, null, { name: 'wall' });

    // The summary names the column, its count, and the largest gap in it --
    // which is how the million is found without reading twenty rows.
    expect(md).toContain('By column');
    expect(md).toMatch(/\| Gross \| 20 \|/);
    expect(md).toContain('1000000');

    // By default the rows are capped: enough to see the shape, and a count of
    // what was left out so the reader knows what they are not being shown.
    expect(md).toContain('and 10 more of these');
    expect(md).toContain('differences.xlsx');
    for (let i = 1; i <= 10; i++) expect(md).toContain(`R${i}`);

    // The per-column tally is never capped -- it is the part that finds the
    // million, and it is a handful of lines however many cells there are.
    expect(md).toMatch(/\| Gross \| 20 \|/);

    // And nothing is lost: asking for it prints every row, as it always did.
    const full = formatMarkdownReport(diff, null, { name: 'wall', detail: 'full' });
    expect(full).toContain('<details><summary>All 20 cells</summary>');
    for (let i = 1; i <= 20; i++) expect(full).toContain(`R${i}`);
    expect(full).not.toContain('and 10 more of these');
  });

  test('cells that will recalculate are named by column, not just by sheet', async () => {
    // A list of addresses says where to look and nothing about what moves.
    // The column header is the part a reader recognises.
    const rows: Row[] = [['Region', 'Gross', 'Share']];
    for (let i = 1; i <= 12; i++) {
      rows.push([`R${i}`, 100 + i, { formula: `B${i + 1}/1000` }]);
    }
    const golden = await sheet('quirk-rec-g.xlsx', rows);

    const drifted: Row[] = [['Region', 'Gross', 'Share']];
    for (let i = 1; i <= 12; i++) {
      drifted.push([`R${i}`, i === 3 ? 999 : 100 + i, { formula: `B${i + 1}/1000` }]);
    }
    const actual = await sheet('quirk-rec-a.xlsx', drifted);

    const { compared, diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false },
      sheets: { Data: { keyColumns: ['Region'] } },
    });
    const swept = await sweep(golden, actual, compared);
    const md = formatMarkdownReport(diff, swept, { name: 'recalc' });

    // Share reads Gross, so it recalculates -- and is reported under its own
    // header rather than as a bare address on the sheet.
    expect(swept.affected.some((a) => a.column === 'Share')).toBe(true);
    expect(md).toContain('| Column | Cells | Where |');
    expect(md).toMatch(/\| Share \| \d+ \|/);
  });

  test('each sheet heads its own table, so the sheets are visibly separate', async () => {
    // Markdown has no per-row border, so whitespace is the only separator
    // available. Once the sheet heads its own block the column that carried it
    // is redundant, and the table gets narrower for it.
    const rows: Row[] = [['Region', 'Gross', 'Share', 'Pct']];
    for (let i = 1; i <= 4; i++) {
      rows.push([`R${i}`, 100 + i, { formula: `B${i + 1}/1000` }, { formula: `B${i + 1}/100` }]);
    }
    const golden = await sheet('quirk-merge-g.xlsx', rows);

    const drifted: Row[] = [['Region', 'Gross', 'Share', 'Pct']];
    for (let i = 1; i <= 4; i++) {
      drifted.push([
        `R${i}`, i === 2 ? 999 : 100 + i,
        { formula: `B${i + 1}/1000` }, { formula: `B${i + 1}/100` },
      ]);
    }
    const actual = await sheet('quirk-merge-a.xlsx', drifted);

    const { compared, diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false },
      sheets: { Data: { keyColumns: ['Region'] } },
    });
    const md = formatMarkdownReport(diff, await sweep(golden, actual, compared), { name: 'merge' });

    // The sheet is a heading carrying its own total, and the table beneath it
    // no longer spends a column repeating that name on every row.
    expect(md).toMatch(/\*\*Data\*\* — \d+ cell\(s\)/);
    expect(md).toContain('| Column | Cells | Where |');
    expect(md).not.toContain('| Sheet | Column | Cells | Where |');
  });

  test('unchecked differences group by sheet, one block each', async () => {
    // The same treatment "Will recalculate differently" gets: a block per
    // sheet, and no column spent repeating the sheet name down every row.
    const golden = await sheet('quirk-gap-sheet-g.xlsx', [
      ['Region', 'Gross'], ['North', 100], ['South', 200],
    ], { sheetName: 'Alpha' });
    const actual = await sheet('quirk-gap-sheet-a.xlsx', [
      ['Region', 'Gross'], ['North', 100], ['South', 999],
    ], { sheetName: 'Alpha' });

    // No key and no positional fallback, so layer 1 reaches nothing and every
    // difference lands in the gap list.
    const { compared, diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false },
      matchUnkeyedRowsByPosition: false,
    });
    const swept = await sweep(golden, actual, compared);
    const md = formatMarkdownReport(diff, swept, { name: 'gaps' });

    expect(md).toContain('## Differing, outside the keyed comparison');
    expect(md).toMatch(/\*\*Alpha\*\* — \d+ cell\(s\)/);
    expect(md).toContain('| Cell | Golden | Actual | Why |');
    expect(md).not.toContain('| Sheet | Cell | Golden | Actual | Why |');
  });

  test('the positional-matching caveat stays visible, its tables fold away', async () => {
    // The warning is the part that has to be read. One of these tables can
    // list sixty column names, which buries everything under it.
    const golden = await sheet('quirk-pos-g.xlsx', [
      ['Region', 'Gross'], ['North', 100], ['South', 200],
    ]);
    const actual = await sheet('quirk-pos-a.xlsx', [
      ['Region', 'Gross'], ['North', 111], ['South', 200],
    ]);

    const { diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false },
      matchUnkeyedRowsByPosition: true,
    });
    const md = formatMarkdownReport(diff, null, { name: 'positional' });
    const at = md.indexOf('## Matched by row position');
    expect(at).toBeGreaterThan(-1);

    const section = md.slice(at);
    // Caveat above the fold, table names below it.
    expect(section.indexOf('paired by')).toBeLessThan(section.indexOf('<details>'));
    expect(section).toContain('<details><summary>Show the tables</summary>');
    expect(section).toContain('keyColumns');
  });

  test('a handful of differences stays flat, with no summary in the way', async () => {
    const golden = await sheet('quirk-few-g.xlsx', [
      ['Region', 'Gross'], ['North', 100], ['South', 200],
    ]);
    const actual = await sheet('quirk-few-a.xlsx', [
      ['Region', 'Gross'], ['North', 111], ['South', 222],
    ]);

    const { diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false },
      sheets: { Data: { keyColumns: ['Region'] } },
    });
    const md = formatMarkdownReport(diff, null, { name: 'few' });

    expect(md).not.toContain('By column');
    expect(md).toContain('North');
    // The only folded block here is the inventory of what was verified; the
    // findings themselves are short enough to stay flat.
    expect(md.split('<details>')).toHaveLength(2);
    expect(md).toContain('## What was verified');
  });
});

/**
 * Formulas are compared as the tool resolves them, not as Excel writes them.
 *
 * A reference becomes `[column name]@row±n`, which is what lets a formula
 * survive its table moving down the sheet -- and it means two formulas can
 * differ while their A1 text is character-for-character identical, because the
 * column one of them points at was renamed. On these reports that happens
 * whenever a header cell holds a date.
 */
test.describe('a formula difference the A1 text cannot show', () => {
  test('the resolved form is shown where the written form is identical', async () => {
    // Column B is named for the run date, so it is renamed between runs. The
    // formula in column C reads it, and its text -- "B2" -- does not change.
    const golden = await sheet('quirk-formula-resolved-g.xlsx', [
      ['Region', '2026-09-02', 'Note'],
      ['North', 100, { formula: 'B2' }],
    ]);
    const actual = await sheet('quirk-formula-resolved-a.xlsx', [
      ['Region', '2026-09-03', 'Note'],
      ['North', 100, { formula: 'B2' }],
    ]);

    const { compared, diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false },
    });
    const swept = await sweep(golden, actual, compared);
    const md = formatMarkdownReport(diff, swept, { name: 'formula-resolved' });

    expect(diff.sheets[0]!.diff!.formulas).toHaveLength(1);
    const f = diff.sheets[0]!.diff!.formulas[0]!;
    // The premise: written identically, resolved differently.
    expect(f.baseA1).toBe(f.nextA1);
    expect(f.base).not.toBe(f.next);

    // So the report must not print the written form twice and call it a
    // difference -- which is what it used to do.
    expect(md).toContain('**Formula changes (1)**');
    expect(md).toContain('[2026-09-02]@row');
    expect(md).toContain('[2026-09-03]@row');
    expect(md).toContain('shown as the');
  });

  test('an ordinary formula change still shows the formula as written', async () => {
    const golden = await sheet('quirk-formula-written-g.xlsx', [
      ['Region', 'Gross', 'Note'],
      ['North', 100, { formula: 'B2*2' }],
    ]);
    const actual = await sheet('quirk-formula-written-a.xlsx', [
      ['Region', 'Gross', 'Note'],
      ['North', 100, { formula: 'B2*3' }],
    ]);

    const { compared, diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false },
    });
    const swept = await sweep(golden, actual, compared);
    const md = formatMarkdownReport(diff, swept, { name: 'formula-written' });

    expect(md).toContain('B2*2');
    expect(md).toContain('B2*3');
    // No footnote, because nothing needed the resolved form.
    expect(md).not.toContain('shown as the');
  });
});

/**
 * A table that changed without failing.
 *
 * Rows arrive or go, a column moves, and every value the two files share still
 * agrees. The verdict says exactly that -- "Something changed — review below"
 * -- and the section listing changed tables is filtered to the ones that
 * failed, so there was nothing below to review.
 *
 * Found on a report that gained five return periods: it invited a review,
 * listed the recalculating cells and the coverage gaps, and never named the
 * five rows. They were in `diff.json` and nowhere else.
 */
test.describe('a change that is not a defect', () => {
  const SPEC: WorkbookSpec = {
    defaults: { requireCachedValues: false },
    sheets: { Data: { keyColumns: ['Region'] } },
  };

  test('the added rows are named, not just counted in the verdict', async () => {
    const golden = await sheet('quirk-review-g.xlsx', [
      ['Region', 'Gross'], ['North', 100], ['South', 200],
    ]);
    const actual = await sheet('quirk-review-a.xlsx', [
      ['Region', 'Gross'], ['North', 100], ['South', 200], ['East', 300],
    ]);

    const { compared, diff } = await runWorkbook(golden, actual, SPEC);
    const swept = await sweep(golden, actual, compared);
    const md = formatMarkdownReport(diff, swept, { name: 'review' });

    // The premise: no defect, but something moved.
    expect(diff.ok).toBe(true);
    expect(diff.reviewOnly).toBe(true);

    expect(md).toContain('Something changed — review below');
    expect(md).toContain('## Changed, and not a defect (1)');
    // ...and the review has something in it.
    expect(md).toContain('**Row population**');
    expect(md).toContain('`East`');
  });

  test('a clean case gets no such section', async () => {
    const golden = await sheet('quirk-review-clean-g.xlsx', [
      ['Region', 'Gross'], ['North', 100],
    ]);
    const actual = await sheet('quirk-review-clean-a.xlsx', [
      ['Region', 'Gross'], ['North', 100],
    ]);

    const { compared, diff } = await runWorkbook(golden, actual, SPEC);
    const swept = await sweep(golden, actual, compared);
    const md = formatMarkdownReport(diff, swept, { name: 'review-clean' });

    expect(md).not.toContain('Changed, and not a defect');
  });
});
