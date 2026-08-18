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

  test('falls back to the widest row when nothing reads as a header', async () => {
    // A wrong header row still beats refusing to compare the block, and the
    // names it produces are visible in the report.
    const path = await sheet('quirk-allnumeric.xlsx', [
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9],
    ]);

    const [table] = await headersOf(path);

    expect(table!.row).toBe(1);
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

  test('a difference in the fifteenth digit is still a difference', async () => {
    // This once asserted the opposite: a gap that small was called rounding
    // from a different order of operations and dropped. It is dropped no
    // longer. A cell whose stored values differ is a cell that differs, and
    // which differences matter is the reader's call -- `tolerance` is how they
    // make it, per column, on purpose.
    const golden = await sheet('quirk-float-g.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.45781166297475],
    ]);
    const actual = await sheet('quirk-float-a.xlsx', [
      ['Region', 'Loss'],
      ['North', 34.45781166297476],
    ]);

    const d = await diffOf(golden, actual);

    expect(d.values).toHaveLength(1);
    expect(d.ok).toBe(false);
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

    const run = await runWorkbook(golden, actual, POSITIONAL);
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
  test('a wall of differences is summarised by column, and still listed in full', async () => {
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

    const { diff } = await runWorkbook(golden, actual, {
      defaults: { requireCachedValues: false },
      sheets: { Data: { keyColumns: ['Region'] } },
    });
    const md = formatMarkdownReport(diff, null, { name: 'wall' });

    // The summary names the column, its count, and the largest gap in it --
    // which is how the million is found without reading twenty rows.
    expect(md).toContain('By column');
    expect(md).toMatch(/\| Gross \| 20 \|/);
    expect(md).toContain('1000000');

    // And nothing was dropped to achieve that.
    expect(md).toContain('<details><summary>All 20 cells</summary>');
    for (let i = 1; i <= 20; i++) expect(md).toContain(`R${i}`);
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

    expect(md).toContain('## Differing, and nothing checked them');
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
    expect(md).not.toContain('<details>');
    expect(md).toContain('North');
  });
});
