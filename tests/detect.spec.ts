import { test, expect } from '@playwright/test';
import { join } from 'node:path';
import { writeFile, mkdir } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import {
  detectWorkbook, detectSpec, detectKeyColumns, specFromDetection,
} from '../src/index.js';
import { buildMultiSheet, buildTwoTableSheet, writeCsv, DIR } from './fixtures.js';

/**
 * Detection is what makes a comparison possible with no hand-written spec.
 * It has to find where each table starts and stops, and which columns identify
 * a row -- and say so plainly when it cannot, rather than inventing a key.
 */

test.describe('detectKeyColumns', () => {
  const rows = [
    ['P-1', 'Ivanov', 'Sofia'],
    ['P-2', 'Petrov', 'Sofia'],
    ['P-3', 'Georgiev', 'Varna'],
  ];

  test('picks the unique column', () => {
    expect(detectKeyColumns(['PolicyId', 'Holder', 'Region'], rows)).toEqual(['PolicyId']);
  });

  test('prefers an identifier-shaped name over one that is merely unique', () => {
    // Holder is unique here too, but only by accident of a small sample.
    expect(detectKeyColumns(['Holder', 'PolicyId', 'Region'], [
      ['Ivanov', 'P-1', 'Sofia'],
      ['Petrov', 'P-2', 'Sofia'],
    ])).toEqual(['PolicyId']);
  });

  test('falls back to a composite key when no single column is unique', () => {
    const key = detectKeyColumns(['PolicyId', 'Period', 'Amount'], [
      ['P-1', '2026-07', '10'],
      ['P-1', '2026-08', '11'],
      ['P-2', '2026-07', '12'],
    ]);
    expect(key).toEqual(['PolicyId', 'Period']);
  });

  test('returns null rather than guessing when nothing identifies a row', () => {
    expect(detectKeyColumns(['Region', 'Band'], [
      ['Sofia', 'A'],
      ['Sofia', 'A'],
    ])).toBeNull();
  });

  test('a column with a blank is passed over, however identifier-shaped its name', () => {
    // Ref looks like the obvious key and is even unique, but a blank means it
    // cannot identify every row.
    expect(detectKeyColumns(['Ref', 'Name'], [
      ['R-1', 'a'],
      ['', 'b'],
    ])).toEqual(['Name']);
  });

  test('a numeric measure that happens to be distinct is not mistaken for a key', () => {
    expect(detectKeyColumns(['Region', 'Amount'], [
      ['Sofia', '10'],
      ['Sofia', '11'],
    ])).toBeNull();
  });

  test('a numeric column still counts when its name says it identifies', () => {
    expect(detectKeyColumns(['Region', 'Invoice No'], [
      ['Sofia', '10'],
      ['Sofia', '11'],
    ])).toEqual(['Invoice No']);
  });
});

test.describe('detectWorkbook', () => {
  test('finds one table per sheet and names it after the sheet', async () => {
    const path = await buildMultiSheet('det-single.xlsx');
    const found = await detectWorkbook(path);

    expect(found.map((s) => s.sheet)).toEqual(['Policies', 'Premiums', 'Regions']);

    const policies = found.find((s) => s.sheet === 'Policies')!;
    expect(policies.tables).toHaveLength(1);
    expect(policies.tables[0]!.name).toBe('Policies');
    expect(policies.tables[0]!.headerRow).toBe(1);
    expect(policies.tables[0]!.keyColumns).toEqual(['PolicyId']);

    const premiums = found.find((s) => s.sheet === 'Premiums')!;
    expect(premiums.tables[0]!.keyColumns).toEqual(['PolicyId', 'Period']);
  });

  test('splits a sheet at the blank row, so an info block does not swallow the data', async () => {
    const path = await buildTwoTableSheet('det-two.xlsx');
    const [sheet] = await detectWorkbook(path);

    expect(sheet!.tables).toHaveLength(2);
    const [info, detail] = sheet!.tables;

    expect(info!.headerRow).toBe(1);
    // Bounded at the block's own last row, not merely above the next table.
    expect(info!.endRow).toBe(4);
    expect(info!.keyColumns).toEqual(['Field']);
    expect(info!.rows).toBe(3);

    expect(detail!.headerRow).toBe(7);
    expect(detail!.endRow).toBe(0);    // runs to the bottom
    expect(detail!.keyColumns).toEqual(['PolicyId']);
    expect(detail!.rows).toBe(5);
  });

  test('a bold section title does not swallow the block underneath it', async () => {
    // The shape that used to disappear: a one-cell heading, painted the way
    // these generators paint a heading, over a two-column key-value block that
    // is not painted at all. The title won the header search on formatting,
    // and the block was then discarded for naming only one column -- so its
    // rows were never compared and nothing said they had not been.
    await mkdir(DIR, { recursive: true });
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Currency Info');

    ws.addRow(['Exchange Rate Information']);
    ws.getCell('A1').font = { bold: true };
    ws.addRow(['Report Currency', 'USD']);
    ws.addRow(['Date of FX', '2025-07']);
    ws.addRow(['# Programs', 8]);
    ws.addRow([]);
    ws.addRow(['Currency Name', 'Unit Per USD']);
    ws.getRow(6).font = { bold: true };
    ws.addRow(['Euro', 1.1771]);
    ws.addRow(['Japanese Yen', 0.00682547]);

    const path = join(DIR, 'det-titled-block.xlsx');
    await wb.xlsx.writeFile(path);
    const [sheet] = await detectWorkbook(path);

    expect(sheet!.tables).toHaveLength(2);
    // The header is the first row that names two columns, not the title above.
    expect(sheet!.tables[0]!.headerRow).toBe(2);
    expect(sheet!.tables[1]!.headerRow).toBe(6);
  });

  test('two tables printed side by side are two tables', async () => {
    // A blank row cannot separate these: both start on row 1 and the taller
    // one keeps every row of the block non-blank. Read as one table, the two
    // header rows fuse, the shorter table's rows read as blank, and a key
    // named in one of them can be found in the other.
    await mkdir(DIR, { recursive: true });
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Report Info');

    const put = (addr: string, v: unknown) => { ws.getCell(addr).value = v as never; };
    put('A1', 'Report ID'); put('B1', 4542);
    put('A2', 'Report Creator'); put('B2', 'S. Arya');
    put('A3', 'Edison Version'); put('B3', '5.0.2');

    put('H1', 'Return Year'); put('I1', 'Start'); put('J1', 'End');
    for (let i = 0; i < 5; i++) {
      put(`H${2 + i}`, 2 + i); put(`I${2 + i}`, i + 1.5); put(`J${2 + i}`, i + 2.5);
    }

    const path = join(DIR, 'det-side-by-side.xlsx');
    await wb.xlsx.writeFile(path);
    const [sheet] = await detectWorkbook(path);

    expect(sheet!.tables).toHaveLength(2);
    // Reading order: down the sheet, then across it.
    const [left, right] = sheet!.tables;
    expect(left!.columns).toBe('A:B');
    expect(right!.columns).toBe('H:J');
    expect(right!.headers).toEqual(['Return Year', 'Start', 'End']);
    expect(right!.rows).toBe(5);

    // The bound reaches the spec, or the reader would take the header row
    // across the whole width and pull one table's columns into the other.
    const spec = specFromDetection(sheet ? [sheet] : []);
    expect(spec.sheets!['Report Info']!.tables!['Table 2']!.columns).toBe('H:J');
  });

  test('a single blank column is a spacer, not a separation', async () => {
    // Dimension breakdowns carry an unnamed column between labels and
    // measures. Cutting there would fragment the table it decorates.
    await mkdir(DIR, { recursive: true });
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Breakdown');
    ws.addRow(['Region', null, 'Gross', 'Net']);
    ws.addRow(['Sofia', null, 10, 8]);
    ws.addRow(['Varna', null, 20, 16]);

    const path = join(DIR, 'det-spacer.xlsx');
    await wb.xlsx.writeFile(path);
    const [sheet] = await detectWorkbook(path);

    expect(sheet!.tables).toHaveLength(1);
    expect(sheet!.tables[0]!.keyColumns).toEqual(['Region']);
  });

  test('a CSV is one table on one pseudo-sheet', async () => {
    const path = await writeCsv('det.csv');
    const found = await detectWorkbook(path);

    expect(found).toHaveLength(1);
    expect(found[0]!.sheet).toBe('CSV');
    expect(found[0]!.tables[0]!.keyColumns).toEqual(['PolicyId']);
  });
});

test.describe('detectSpec', () => {
  test('produces a spec the comparison can run unchanged', async () => {
    const path = await buildTwoTableSheet('det-spec.xlsx');
    const spec = await detectSpec(path);

    expect(spec.sheets!['Policies']!.tables).toBeDefined();
    const tables = spec.sheets!['Policies']!.tables!;
    expect(Object.keys(tables)).toEqual(['Table 1', 'Table 2']);
    expect(tables['Table 1']!.endRow).toBe(4);
    expect(tables['Table 2']!.keyColumns).toEqual(['PolicyId']);
    // The last table carries no endRow, so growth needs no re-detection.
    expect(tables['Table 2']!.endRow).toBeUndefined();
  });

  test('a single-table sheet needs no tables block at all', async () => {
    const path = await buildMultiSheet('det-flat.xlsx');
    const spec = await detectSpec(path);

    expect(spec.sheets!['Policies']!.tables).toBeUndefined();
    expect(spec.sheets!['Policies']!.keyColumns).toEqual(['PolicyId']);
  });

  test('a table whose rows cannot be keyed is left without one, not invented', async () => {
    const path = join(DIR, 'det-nokey.csv');
    await mkdir(DIR, { recursive: true });
    // Every column repeats, so nothing identifies a row.
    await writeFile(path, 'Region,Band\nSofia,A\nSofia,A\nVarna,B\n', 'utf8');

    const spec = await detectSpec(path);
    // The guarantee: no key is invented. A wrong key pairs rows arbitrarily
    // and produces a confident wrong answer, which is worse than no answer.
    expect(spec.sheets!['CSV']!.keyColumns).toBeUndefined();

    const { resolveTables } = await import('../src/index.js');

    // The table is still compared, by row position, and says so.
    const outcome = resolveTables(spec, 'CSV')[0]!;
    expect(outcome.spec!.keyColumns).toEqual([]);
    expect(outcome.spec!.matchRowsByPosition).toBe(true);
    expect(outcome.reason).toContain('position');

    // Turning the fallback off goes back to not comparing it at all.
    const strict = resolveTables({ ...spec, matchUnkeyedRowsByPosition: false }, 'CSV')[0]!;
    expect(strict.spec).toBeNull();
    expect(strict.reason).toContain('keyColumns');
  });
});

/**
 * A key-value block has no header row. Every row is a label and its value, so
 * picking one costs twice over: the row picked stops being data, and the value
 * column takes its name from a *value*. Where that value is the report's own
 * name or id it differs between any two runs, so the column pairs with nothing
 * in the other file and the whole block reports as one column removed and
 * another added.
 */
test.describe('a block with no header row', () => {
  /** The report info block every one of these generators opens a sheet with. */
  async function buildInfoBlock(name: string, reportName: string): Promise<string> {
    await mkdir(DIR, { recursive: true });
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Report Info');

    const pairs: [string, string | number][] = [
      ['Report ID', 4537],
      ['Pro-Forma Report Name', reportName],
      ['View Of Risk', 'RMS_v18.1'],
      ['Report Creator', 'Sumedha Arya'],
      ['Edison Version', '5.0.2'],
    ];
    for (const [label, value] of pairs) ws.addRow([label, value]);
    // Painted down the label column on every row, which is how these reports
    // mark a label -- not how they mark a header.
    for (let r = 1; r <= pairs.length; r++) {
      ws.getRow(r).getCell(1).font = { bold: true };
    }

    const path = join(DIR, name);
    await wb.xlsx.writeFile(path);
    return path;
  }

  test('is read as data with columns named after themselves', async () => {
    const path = await buildInfoBlock('det-kv-block.xlsx', 'Report_5.0.2_0807');
    const [sheet] = await detectWorkbook(path);

    const table = sheet!.tables[0]!;
    // headerRow names the row *above* the data, so a block starting at the top
    // of the sheet has none: row 0 does not exist.
    expect(table.headerRow).toBe(0);
    expect(table.rows).toBe(5); // every row is data, including the first two
    expect(table.headers).toEqual(['Column A', 'Column B']);
    expect(table.keyColumns).toEqual(['Column A']);
  });

  test('does not report the value column as removed and added when a run id changes', async () => {
    // The failure this exists to stop. Two genuinely different runs differ in
    // the report name, and that value used to be the column's *name*.
    const golden = await buildInfoBlock('det-kv-golden.xlsx', 'Report_5.0.2_0807');
    const actual = await buildInfoBlock('det-kv-actual.xlsx', 'Report_5.1_0818');

    const { runWorkbook } = await import('../src/index.js');
    const { diff } = await runWorkbook(golden, actual, await detectSpec(golden));

    const sheet = diff.sheets.find((s) => s.sheet === 'Report Info')!;
    expect(sheet.diff!.schema.added).toEqual([]);
    expect(sheet.diff!.schema.removed).toEqual([]);
    // Five rows, and the one that really changed is reported as a value.
    expect(sheet.diff!.rows.compared).toBe(5);
  });

  test('an unpainted block keeps its header row', async () => {
    // The guard. Every row of a plain table is painted identically too -- not
    // at all -- so without excluding those, any report arriving as an unstyled
    // grid would lose its column names.
    await mkdir(DIR, { recursive: true });
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Plain');
    ws.addRow(['PolicyId', 'Region']);
    ws.addRow(['P-1', 'Sofia']);
    ws.addRow(['P-2', 'Varna']);

    const path = join(DIR, 'det-plain-grid.xlsx');
    await wb.xlsx.writeFile(path);
    const [sheet] = await detectWorkbook(path);

    expect(sheet!.tables[0]!.headerRow).toBe(1);
    expect(sheet!.tables[0]!.headers).toEqual(['PolicyId', 'Region']);
  });

  test('headerRow 0 can be written by hand for a block detection got wrong', async () => {
    const path = await buildInfoBlock('det-kv-manual.xlsx', 'Report_5.0.2_0807');
    const { runWorkbook } = await import('../src/index.js');

    const { compared } = await runWorkbook(path, path, {
      sheets: { 'Report Info': { headerRow: 0, keyColumns: ['Column A'] } },
    });
    const table = compared[0]!;
    expect(table.base.headers).toEqual(['Column A', 'Column B']);
    // The range covers the data alone -- there is no row 0 to name.
    expect(table.diff.rows.compared).toBe(5);
  });
});
