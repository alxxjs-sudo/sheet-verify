import { test, expect } from '@playwright/test';
import { readFile, rm, access } from 'node:fs/promises';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { runCase, formatLedger, runWorkbook } from '../src/index.js';
import type { CaseOptions } from '../src/index.js';
import { buildMultiSheet, buildTwoTableSheet, DIR } from './fixtures.js';

const SPEC: CaseOptions = {
  sheets: {
    Policies: { keyColumns: ['PolicyId'] },
    Premiums: { keyColumns: ['PolicyId', 'Period'] },
    Regions: { keyColumns: ['Region'] },
  },
};

/** A fresh case folder per test, so runs cannot contaminate each other. */
const caseDir = async (name: string) => {
  const dir = join(DIR, 'cases', name);
  await rm(dir, { recursive: true, force: true });
  return dir;
};

const exists = (p: string) => access(p).then(() => true, () => false);
const rows = (csv: string) => csv.trim().split('\n');

/** Parses the CSV ledger by header name, so a new column cannot break a test. */
const parseLedger = (csv: string): Record<string, string>[] => {
  const [header, ...body] = rows(csv);
  const names = header!.split(',');
  return body.map((line) =>
    Object.fromEntries(line.split(',').map((v, i) => [names[i]!, v])),
  );
};

/** Reads the xlsx ledger back as headers plus keyed rows. */
async function readLedger(path: string): Promise<{
  headers: string[];
  rows: Record<string, string | number | null>[];
}> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  const ws = wb.getWorksheet('Cells')!;

  const headers: string[] = [];
  ws.getRow(1).eachCell((cell, c) => { headers[c - 1] = String(cell.value ?? ''); });

  const out: Record<string, string | number | null>[] = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const rec: Record<string, string | number | null> = {};
    headers.forEach((h, i) => {
      const v = ws.getRow(r).getCell(i + 1).value;
      rec[h] = v === undefined || v === '' ? null : (v as string | number);
    });
    out.push(rec);
  }
  return { headers, rows: out };
}

test.describe('runCase', () => {
  test('creates the golden output on first run and passes', async () => {
    const dir = await caseDir('first-run');
    const actual = await buildMultiSheet('case-first.xlsx');

    const r = await runCase(actual, dir, SPEC);

    expect(r.blessed).toBe(true);
    expect(r.ok).toBe(true);
    expect(r.diff).toBeNull();
    expect(await exists(r.files.golden)).toBe(true);
    // The new report is kept even on the blessing run.
    expect(await exists(r.files.actual)).toBe(true);
  });

  test('copies the new report into the case folder and writes all three artefacts', async () => {
    const dir = await caseDir('artefacts');
    const golden = await buildMultiSheet('case-art-golden.xlsx');
    const actual = await buildMultiSheet('case-art-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);          // establishes the golden output
    const r = await runCase(actual, dir, SPEC); // the run under test

    expect(r.ok).toBe(false);
    for (const f of [r.files.actual, r.files.diffText, r.files.diffJson, r.files.cells]) {
      expect(await exists(f)).toBe(true);
    }

    const text = await readFile(r.files.diffText, 'utf8');
    expect(text).toContain('artefacts');       // the case name heads the file
    expect(text).toContain('P-1003');

    const json = JSON.parse(await readFile(r.files.diffJson, 'utf8'));
    expect(json.ok).toBe(false);
    // The drifted Amount plus the Tax formula downstream of it: one cause,
    // one consequence.
    const values = json.sheets.find((s: any) => s.sheet === 'Premiums').diff.values;
    expect(values).toHaveLength(2);
    expect(values.filter((v: any) => v.rootCause)).toHaveLength(1);
  });

  test('the ledger holds only differences by default — a matching cell earns no row', async () => {
    const dir = await caseDir('ledger-default');
    const golden = await buildMultiSheet('case-led-golden.xlsx');
    const actual = await buildMultiSheet('case-led-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    expect(r.files.cells.endsWith('cells.xlsx')).toBe(true);
    const ws = await readLedger(r.files.cells);

    expect(ws.headers).toEqual([
      'Sheet', 'Table', 'Row key', 'Column', 'Status', 'Root cause',
      'Golden cell', 'Actual cell', 'Golden value', 'Actual value',
      'Delta', 'Tolerance', 'Golden formula', 'Actual formula',
    ]);
    expect(ws.rows.some((r) => r['Status'] === 'match')).toBe(false);

    // The drifted Amount, and the Tax formula that reads it.
    const differing = ws.rows.filter((r) => r['Status'] === 'value-differs');
    expect(differing).toHaveLength(2);

    const cause = differing.filter((r) => r['Root cause'] === 'yes');
    expect(cause).toHaveLength(1);
    expect(cause[0]!['Row key']).toBe('P-1003 / 2026-08');
    expect(cause[0]!['Column']).toBe('Amount');
    expect(cause[0]!['Actual value']).toBe(9999);

    const consequence = differing.filter((r) => r['Root cause'] === 'no');
    expect(consequence).toHaveLength(1);
    expect(consequence[0]!['Column']).toBe('Tax');
  });

  test('numbers are written as numbers, so the sheet sorts and filters correctly', async () => {
    const dir = await caseDir('ledger-types');
    const golden = await buildMultiSheet('case-typ-golden.xlsx');
    const actual = await buildMultiSheet('case-typ-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    const ws = await readLedger(r.files.cells);
    const cause = ws.rows.find((r) => r['Root cause'] === 'yes')!;

    expect(typeof cause['Golden value']).toBe('number');
    expect(typeof cause['Actual value']).toBe('number');
    expect(cause['Delta']).toBe(3999);
  });

  test('the ledger arrives as a real Excel table, with the header frozen', async () => {
    const dir = await caseDir('ledger-style');
    const golden = await buildMultiSheet('case-sty-golden.xlsx');
    const actual = await buildMultiSheet('case-sty-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(r.files.cells);
    const ws = wb.getWorksheet('Cells')!;

    expect(ws.views[0]).toMatchObject({ state: 'frozen', ySplit: 1 });
    expect(ws.getRow(1).font?.bold).toBe(true);
    // Every column is widened past the default.
    for (let c = 1; c <= 14; c++) expect(ws.getColumn(c).width).toBeGreaterThan(9);
    // The status cell is colour-coded rather than left plain.
    const status = ws.getCell(2, 5);
    expect((status.fill as any)?.fgColor?.argb).toBeTruthy();
  });

  test('header text contrasts with the header fill rather than vanishing into it', async () => {
    const dir = await caseDir('ledger-contrast');
    const golden = await buildMultiSheet('case-con-golden.xlsx');
    const actual = await buildMultiSheet('case-con-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(r.files.cells);

    for (const path of [r.files.cells, r.files.compared]) {
      const book = new ExcelJS.Workbook();
      await book.xlsx.readFile(path);
      for (const ws of book.worksheets) {
        for (let c = 1; c <= 7; c++) {
          const cell = ws.getCell(1, c);
          if (!cell.value) continue;
          const text = (cell.font?.color as any)?.argb;
          const fill = (cell.fill as any)?.fgColor?.argb;
          expect(fill).toBeTruthy();
          expect(text).toBeTruthy();
          // Light on dark: the two must not be the same shade.
          expect(String(text).toUpperCase()).not.toBe(String(fill).toUpperCase());
          expect(String(text).toUpperCase()).toBe('FFFFFFFF');
        }
      }
    }
  });

  test('compared.xlsx holds every compared cell, one worksheet per table', async () => {
    const dir = await caseDir('compared');
    const golden = await buildMultiSheet('case-cmp-golden.xlsx');
    const actual = await buildMultiSheet('case-cmp-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(r.files.compared);

    expect(wb.worksheets.map((w) => w.name)).toEqual(['Policies', 'Premiums', 'Regions']);

    const ws = wb.getWorksheet('Premiums')!;
    expect(ws.getRow(1).values).toEqual([
      undefined, 'Row key', 'Column', 'Golden cell', 'Actual cell',
      'Golden value', 'Actual value', 'Status',
    ]);
    // 10 rows x 4 columns, matches included -- the point of this file.
    expect(ws.rowCount - 1).toBe(40);
    expect(ws.views[0]).toMatchObject({ state: 'frozen', ySplit: 1 });

    const statuses = new Set<string>();
    for (let r2 = 2; r2 <= ws.rowCount; r2++) {
      statuses.add(String(ws.getRow(r2).getCell(7).value));
    }
    expect(statuses).toContain('match');
    expect(statuses).toContain('value-differs');
  });

  test('a sheet holding two tables gets a worksheet each, named for the table', async () => {
    const dir = await caseDir('compared-tables');
    const golden = await buildTwoTableSheet('cmp-tt-golden.xlsx');
    const actual = await buildTwoTableSheet('cmp-tt-actual.xlsx', { release: '4.3.0' });

    const spec: CaseOptions = {
      sheets: {
        Policies: {
          tables: {
            Info: { headerRow: 1, keyColumns: ['Field'] },
            Detail: { headerRow: 7, keyColumns: ['PolicyId'] },
          },
        },
      },
    };

    await runCase(golden, dir, spec);
    const r = await runCase(actual, dir, spec);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(r.files.compared);
    expect(wb.worksheets.map((w) => w.name)).toEqual(['Policies · Info', 'Policies · Detail']);

    // The info block is 3 rows x 2 columns; the data table 5 x 5.
    expect(wb.getWorksheet('Policies · Info')!.rowCount - 1).toBe(6);
    expect(wb.getWorksheet('Policies · Detail')!.rowCount - 1).toBe(25);
  });

  test('the compared workbook can be turned off', async () => {
    const dir = await caseDir('compared-off');
    const golden = await buildMultiSheet('case-off-golden.xlsx');
    const actual = await buildMultiSheet('case-off-actual.xlsx');

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, { ...SPEC, comparedLedger: false });

    expect(await exists(r.files.compared)).toBe(false);
    expect(await exists(r.files.cells)).toBe(true);
  });

  test('a clean run still produces a readable, empty ledger', async () => {
    const dir = await caseDir('ledger-clean');
    const golden = await buildMultiSheet('case-cln-golden.xlsx');
    const actual = await buildMultiSheet('case-cln-actual.xlsx');

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    expect(r.ok).toBe(true);
    const ws = await readLedger(r.files.cells);
    expect(ws.headers[0]).toBe('Sheet');
    expect(ws.rows).toHaveLength(0);
  });

  test('naming the ledger .csv streams text instead, for large cases', async () => {
    const dir = await caseDir('ledger-csv');
    const golden = await buildMultiSheet('case-csv-golden.xlsx');
    const actual = await buildMultiSheet('case-csv-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, {
      ...SPEC, names: { cells: 'cells.csv' }, cellLedger: 'all',
    });

    const lines = rows(await readFile(r.files.cells, 'utf8'));
    expect(lines[0]).toBe(
      'Sheet,Table,Row key,Column,Status,Root cause,Golden cell,Actual cell,' +
      'Golden value,Actual value,Delta,Tolerance,Golden formula,Actual formula',
    );
    // 'all' keeps the matches, which is what makes the CSV path worth having.
    expect(lines.filter((l) => l.includes(',match,')).length).toBeGreaterThan(50);
  });

  test('re-blessing replaces the golden output and the next run is clean', async () => {
    const dir = await caseDir('rebless');
    const golden = await buildMultiSheet('case-reb-golden.xlsx');
    const actual = await buildMultiSheet('case-reb-actual.xlsx', {
      premiumDrift: { 'P-1001|2026-07': 7 },
    });

    await runCase(golden, dir, SPEC);
    expect((await runCase(actual, dir, SPEC)).ok).toBe(false);

    const blessed = await runCase(actual, dir, { ...SPEC, updateGolden: true });
    expect(blessed.blessed).toBe(true);

    expect((await runCase(actual, dir, SPEC)).ok).toBe(true);
  });

  test('refuses to invent a golden output when createMissingGolden is off', async () => {
    const dir = await caseDir('no-invent');
    const actual = await buildMultiSheet('case-noinv.xlsx');

    await expect(
      runCase(actual, dir, { ...SPEC, createMissingGolden: false }),
    ).rejects.toThrow(/golden output not found/);
  });

  test('reports a missing new report clearly', async () => {
    const dir = await caseDir('missing');
    await expect(runCase(join(DIR, 'nope.xlsx'), dir, SPEC)).rejects.toThrow(/not found/);
  });

  test('file names inside the folder can be overridden', async () => {
    const dir = await caseDir('renamed');
    const actual = await buildMultiSheet('case-rn.xlsx');

    const r = await runCase(actual, dir, {
      ...SPEC,
      names: { golden: 'baseline.xlsx', cells: 'audit.csv' },
    });

    expect(r.files.golden.endsWith('baseline.xlsx')).toBe(true);
    expect(r.files.cells.endsWith('audit.csv')).toBe(true);
  });
});

test.describe('cell ledger', () => {
  test('classifies ignored columns, added rows and formula changes distinctly', async () => {
    const golden = await buildMultiSheet('led-golden.xlsx');
    const actual = await buildMultiSheet('led-actual.xlsx', { insertPremium: true });

    const { compared } = await runWorkbook(golden, actual, {
      sheets: {
        Policies: { keyColumns: ['PolicyId'], ignoreColumns: ['Holder'] },
      },
    });
    const lines = rows(formatLedger(compared, 'all'));

    const statuses = new Set(lines.slice(1).map((l) => l.split(',')[4]));
    expect(statuses).toContain('match');
    expect(statuses).toContain('ignored-column');
    expect(statuses).toContain('column-added');

    // The inserted column shifts Annual Cost's references, but header
    // resolution means the formula still matches.
    expect(statuses).not.toContain('formula-differs');
  });

  test('records the tolerance a numeric cell was judged against', async () => {
    const golden = await buildMultiSheet('led-tol-golden.xlsx');
    const actual = await buildMultiSheet('led-tol-actual.xlsx', {
      premiumDrift: { 'P-1001|2026-07': 2520.5 },
    });

    const { compared } = await runWorkbook(golden, actual, {
      sheets: {
        Premiums: { keyColumns: ['PolicyId', 'Period'], tolerance: { Amount: 1 } },
      },
    });
    const within = parseLedger(formatLedger(compared, 'differences'))
      .filter((r) => r['Status'] === 'within-tolerance');

    expect(within).toHaveLength(1);
    expect(within[0]!['Tolerance']).toBe('1');
    expect(within[0]!['Delta']).toBe('0.5');
    expect(within[0]!['Column']).toBe('Amount');
  });

  test('none writes nothing at all', async () => {
    const golden = await buildMultiSheet('led-none-golden.xlsx');
    const actual = await buildMultiSheet('led-none-actual.xlsx');
    const { compared } = await runWorkbook(golden, actual, {
      sheets: { Policies: { keyColumns: ['PolicyId'] } },
    });

    expect(formatLedger(compared, 'none')).toBe('');
  });
});
