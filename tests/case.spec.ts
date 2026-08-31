import { test, expect } from '@playwright/test';
import { readFile, rm, access } from 'node:fs/promises';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { runCase, formatLedger, runWorkbook } from '../src/index.js';
import type { CaseOptions } from '../src/index.js';
import { buildMultiSheet, buildTwoTableSheet, buildSweepWorkbook, DIR } from './fixtures.js';

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
  const ws = wb.getWorksheet("Differences")!;

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

test.describe('compared.xlsx value columns', () => {
  test('a formula with no stored result shows the formula, not an empty cell', async () => {
    // A file straight from the generator stores no computed results, so these
    // cells came back blank -- indistinguishable from a genuinely empty cell,
    // which is the opposite of what this file is for.
    const dir = await caseDir('formula-display');
    const SPEC = {
      defaults: { requireCachedValues: false },
      sheets: { Policies: { keyColumns: ['PolicyId'] } },
    } as CaseOptions;

    const golden = await buildSweepWorkbook('cmp-fx-golden.xlsx');
    await runCase(golden, dir, SPEC);
    const actual = await buildSweepWorkbook('cmp-fx-actual.xlsx', { sumDrift: { 'P-1002': 90000 } });
    const r = await runCase(actual, dir, SPEC);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(r.files.compared);
    const ws = wb.worksheets.find((w) => w.name.includes('Policies'))!;

    const headers: string[] = [];
    ws.getRow(2).eachCell((c, i) => { headers[i - 1] = String(c.value ?? ''); });
    const col = headers.indexOf('Golden value') + 1;

    const annual: string[] = [];
    for (let row = 3; row <= ws.rowCount; row++) {
      if (String(ws.getRow(row).getCell(2).value ?? '') !== 'Annual Cost') continue;
      annual.push(String(ws.getRow(row).getCell(col).value ?? ''));
    }

    expect(annual.length).toBeGreaterThan(0);
    expect(annual.every((v) => v.startsWith('fx '))).toBe(true);
    expect(annual[0]).toContain('C2*D2');
    // Not written as "=..." -- Excel would try to evaluate it, and those
    // references mean nothing in this workbook.
    expect(annual.every((v) => !v.startsWith('='))).toBe(true);
  });

  test('a cell that is genuinely empty stays empty', async () => {
    const dir = await caseDir('formula-display-blank');
    const SPEC = {
      defaults: { requireCachedValues: false },
      sheets: { Policies: { keyColumns: ['PolicyId'] }, Notes: { headerRow: 2 } },
    } as CaseOptions;

    const golden = await buildSweepWorkbook('cmp-blank-golden.xlsx');
    await runCase(golden, dir, SPEC);
    const r = await runCase(await buildSweepWorkbook('cmp-blank-actual.xlsx'), dir, SPEC);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(r.files.compared);
    const ws = wb.worksheets.find((w) => w.name.includes('Notes'))!;

    // Nothing on Notes is a formula, so nothing there should be marked as one.
    let seen = 0;
    for (let row = 2; row <= ws.rowCount; row++) {
      const v = String(ws.getRow(row).getCell(5).value ?? '');
      expect(v.startsWith('fx ')).toBe(false);
      seen++;
    }
    expect(seen).toBeGreaterThan(0);
  });
});

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
    for (const f of [r.files.actual, r.files.report, r.files.diffJson, r.files.differences]) {
      expect(await exists(f)).toBe(true);
    }

    const text = await readFile(r.files.report, 'utf8');
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

  test('a recalculated run says so at the top of its report', async () => {
    const dir = await caseDir('recalc-notice');
    const golden = await buildSweepWorkbook('notice-golden.xlsx');
    const actual = await buildSweepWorkbook('notice-actual.xlsx', {
      sumDrift: { 'P-1001': 999999 },
    });
    const spec: CaseOptions = {
      sheets: { Policies: { keyColumns: ['PolicyId'] } },
      recalculated: true,
    };

    await runCase(golden, dir, spec);
    const r = await runCase(actual, dir, spec);
    const report = await readFile(r.files.report, 'utf8');

    // What a comparison could see changes the meaning of everything under it,
    // so it is stated rather than left to be inferred from the findings.
    expect(report).toContain('Recalculated before comparison');
    expect(report).toContain('untouched');

    // And not claimed when it did not happen.
    const plainDir = await caseDir('recalc-none');
    const plainSpec: CaseOptions = { sheets: { Policies: { keyColumns: ['PolicyId'] } } };
    await runCase(golden, plainDir, plainSpec);
    const plain = await runCase(actual, plainDir, plainSpec);
    expect(await readFile(plain.files.report, 'utf8')).not.toContain('Recalculated');
  });

  test('differences.xlsx names the cells that will recalculate, and what drives each', async () => {
    const dir = await caseDir('ledger-recalc');
    // Formulas with no cached result, as the real reports arrive: the drifted
    // Sum Insured differs outright, and the Annual Cost that reads it has
    // nothing stored to compare, so it can only be reported as recalculating.
    const golden = await buildSweepWorkbook('case-recalc-golden.xlsx');
    const actual = await buildSweepWorkbook('case-recalc-actual.xlsx', {
      sumDrift: { 'P-1001': 999999 },
    });
    const spec: CaseOptions = { sheets: { Policies: { keyColumns: ['PolicyId'] } } };

    await runCase(golden, dir, spec);
    const r = await runCase(actual, dir, spec);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(r.files.differences);

    // Beside the differences, not instead of them.
    expect(wb.worksheets.map((w) => w.name)).toContain('Differences');
    const ws = wb.getWorksheet('Will recalculate')!;
    expect(ws).toBeTruthy();

    const headers: string[] = [];
    ws.getRow(1).eachCell((cell, c) => { headers[c - 1] = String(cell.value ?? ''); });
    expect(headers).toEqual([
      'Sheet', 'Cell', 'Column', 'How',
      'Driven by sheet', 'Driven by cell',
      'Golden (value or formula)', 'Actual (value or formula)',
    ]);

    const rows: Record<string, string>[] = [];
    for (let n = 2; n <= ws.rowCount; n++) {
      const rec: Record<string, string> = {};
      headers.forEach((h, i) => { rec[h] = String(ws.getRow(n).getCell(i + 1).value ?? ''); });
      rows.push(rec);
    }

    // The Annual Cost cell on the drifted row, named with the change that
    // drives it -- which is the whole reason the sheet exists. Without it a
    // reader who saw this cell move in Excel finds no row for it anywhere.
    const driven = rows.find((x) => x['How'] === 'reads it directly')!;
    expect(driven).toBeTruthy();
    expect(driven['Sheet']).toBe('Policies');
    expect(driven['Column']).toBe('Annual Cost');
    expect(driven['Driven by sheet']).toBe('Policies');
    expect(driven['Driven by cell']).not.toBe('');
    expect(driven['Actual (value or formula)']).toContain('999999');
  });

  test('the ledger holds only differences by default — a matching cell earns no row', async () => {
    const dir = await caseDir('ledger-default');
    const golden = await buildMultiSheet('case-led-golden.xlsx');
    const actual = await buildMultiSheet('case-led-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    expect(r.files.differences.endsWith('differences.xlsx')).toBe(true);
    const ws = await readLedger(r.files.differences);

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

    const ws = await readLedger(r.files.differences);
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
    await wb.xlsx.readFile(r.files.differences);
    const ws = wb.getWorksheet("Differences")!;

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
    await wb.xlsx.readFile(r.files.differences);

    // compared.xlsx carries a caption on row 1, so its header row is row 2.
    for (const [path, headerRow] of [
      [r.files.differences, 1] as const,
      [r.files.compared, 2] as const,
    ]) {
      const book = new ExcelJS.Workbook();
      await book.xlsx.readFile(path);
      for (const ws of book.worksheets) {
        for (let c = 1; c <= 7; c++) {
          const cell = ws.getCell(headerRow, c);
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
    // Row 1 says which rectangle of the source sheet these rows came from,
    // and is painted as a banner rather than as a footnote -- it is the first
    // question anyone opening the tab has.
    const banner = ws.getRow(1);
    expect(String(banner.getCell(1).value)).toContain('A1:D11');
    expect(String(banner.getCell(1).value)).toContain('rows matched by key');
    expect(banner.getCell(1).font).toMatchObject({ bold: true, size: 12 });
    expect(banner.height).toBeGreaterThan(20);
    // Filled the full width of the table, not just under the text.
    for (let c = 1; c <= 7; c++) {
      expect((ws.getCell(1, c).fill as any)?.fgColor?.argb).toBeTruthy();
    }
    // Light banner over dark header: the two must not read as one block.
    const bannerFill = (ws.getCell(1, 1).fill as any).fgColor.argb;
    const headerFill = (ws.getCell(2, 1).fill as any).fgColor.argb;
    expect(String(bannerFill).toUpperCase()).not.toBe(String(headerFill).toUpperCase());
    expect(ws.getRow(2).values).toEqual([
      undefined, 'Row key', 'Column', 'Golden cell', 'Actual cell',
      'Golden value', 'Actual value', 'Status',
    ]);
    // 10 rows x 4 columns, matches included -- the point of this file.
    expect(ws.rowCount - 2).toBe(40);
    expect(ws.views[0]).toMatchObject({ state: 'frozen', ySplit: 2 });

    const statuses = new Set<string>();
    for (let r2 = 3; r2 <= ws.rowCount; r2++) {
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

    // The info block is 3 rows x 2 columns; the data table 5 x 5. Two rows of
    // preamble now: the range caption, then the column headers.
    expect(wb.getWorksheet('Policies · Info')!.rowCount - 2).toBe(6);
    expect(wb.getWorksheet('Policies · Detail')!.rowCount - 2).toBe(25);
  });

  test('non-matching cells are highlighted in compared.xlsx, matches left plain', async () => {
    const dir = await caseDir('compared-highlight');
    const golden = await buildMultiSheet('case-hl-golden.xlsx');
    const actual = await buildMultiSheet('case-hl-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(r.files.compared);
    const ws = wb.getWorksheet('Premiums')!;

    let painted = 0;
    let plain = 0;
    for (let n = 3; n <= ws.rowCount; n++) {
      const row = ws.getRow(n);
      const status = String(row.getCell(7).value);
      const fill = (row.getCell(7).fill as any)?.fgColor?.argb;

      if (status === 'match') {
        expect(fill).toBeFalsy();
        plain++;
      } else {
        expect(fill).toBeTruthy();
        // The values that differ are emphasised alongside the verdict.
        expect(row.getCell(5).font?.bold).toBe(true);
        expect(row.getCell(6).font?.bold).toBe(true);
        painted++;
      }
    }

    expect(painted).toBe(2);          // the drifted Amount and its Tax cascade
    expect(plain).toBe(38);
  });

  test('the compared workbook can be turned off', async () => {
    const dir = await caseDir('compared-off');
    const golden = await buildMultiSheet('case-off-golden.xlsx');
    const actual = await buildMultiSheet('case-off-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, { ...SPEC, comparedLedger: false });

    expect(await exists(r.files.compared)).toBe(false);
    // The differences file is unaffected -- there are differences to record.
    expect(await exists(r.files.differences)).toBe(true);
  });

  test('a clean run writes no differences file at all', async () => {
    const dir = await caseDir('ledger-clean');
    const golden = await buildMultiSheet('case-cln-golden.xlsx');
    const actual = await buildMultiSheet('case-cln-actual.xlsx');

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    expect(r.ok).toBe(true);
    // An empty differences file reads as a fault rather than as the answer.
    expect(await exists(r.files.differences)).toBe(false);
    // The record of what *was* checked is still there.
    expect(await exists(r.files.compared)).toBe(true);
  });

  test('a differences file left by an earlier run is cleared once the case is clean', async () => {
    const dir = await caseDir('ledger-stale');
    const golden = await buildMultiSheet('case-stl-golden.xlsx');
    const drifted = await buildMultiSheet('case-stl-drifted.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const bad = await runCase(drifted, dir, SPEC);
    expect(await exists(bad.files.differences)).toBe(true);

    // The defect is fixed and the report matches again. A stale file here
    // would describe a comparison that no longer holds.
    const fixed = await runCase(golden, dir, SPEC);
    expect(fixed.ok).toBe(true);
    expect(await exists(fixed.files.differences)).toBe(false);
  });

  test('a clean run clears a stale CSV ledger too', async () => {
    const dir = await caseDir('ledger-stale-csv');
    const golden = await buildMultiSheet('case-stc-golden.xlsx');
    const drifted = await buildMultiSheet('case-stc-drifted.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });
    const names = { differences: 'differences.csv' };

    await runCase(golden, dir, { ...SPEC, names });
    const bad = await runCase(drifted, dir, { ...SPEC, names });
    expect(await exists(bad.files.differences)).toBe(true);

    const fixed = await runCase(golden, dir, { ...SPEC, names });
    expect(await exists(fixed.files.differences)).toBe(false);
  });

  test('cellLedger none writes no differences file either', async () => {
    const dir = await caseDir('ledger-off');
    const golden = await buildMultiSheet('case-lof-golden.xlsx');
    const actual = await buildMultiSheet('case-lof-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, { ...SPEC, cellLedger: 'none' });

    expect(r.ok).toBe(false);
    expect(await exists(r.files.differences)).toBe(false);
    expect(await exists(r.files.report)).toBe(true);
  });

  test('naming the ledger .csv streams text instead, for large cases', async () => {
    const dir = await caseDir('ledger-csv');
    const golden = await buildMultiSheet('case-csv-golden.xlsx');
    const actual = await buildMultiSheet('case-csv-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, {
      ...SPEC, names: { differences: 'differences.csv' }, cellLedger: 'all',
    });

    const lines = rows(await readFile(r.files.differences, 'utf8'));
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
      names: { golden: 'baseline.xlsx', differences: 'audit.csv' },
    });

    expect(r.files.golden.endsWith('baseline.xlsx')).toBe(true);
    expect(r.files.differences.endsWith('audit.csv')).toBe(true);
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
    // Kept out of the default ledger: differences.xlsx is read as a list of
    // things to fix, and a cell the tolerance already forgave is not one. It
    // is counted and listed in report.md instead, and the full ledger still
    // carries it with the tolerance it was judged against.
    expect(parseLedger(formatLedger(compared, 'differences'))
      .filter((r) => r['Status'] === 'within-tolerance')).toHaveLength(0);

    const within = parseLedger(formatLedger(compared, 'all'))
      .filter((r) => r['Status'] === 'within-tolerance');

    expect(within).toHaveLength(1);
    expect(within[0]!['Tolerance']).toBe('1');
    expect(within[0]!['Delta']).toBe('0.5');
    expect(within[0]!['Column']).toBe('Amount');
  });

  test('records the allowance a cell was judged by, not the one its column was given', async () => {
    const golden = await buildMultiSheet('led-rel-golden.xlsx');
    const actual = await buildMultiSheet('led-rel-actual.xlsx', {
      premiumDrift: { 'P-1001|2026-07': 2520.0000000004 },
    });

    const { compared } = await runWorkbook(golden, actual, {
      sheets: {
        Premiums: {
          keyColumns: ['PolicyId', 'Period'],
          tolerance: { Amount: 0 },
          relativeTolerance: { Amount: 1e-12 },
        },
      },
    });

    const within = parseLedger(formatLedger(compared, 'all'))
      .filter((r) => r['Status'] === 'within-tolerance');
    const amount = within.find((r) => r['Column'] === 'Amount')!;
    expect(amount).toBeDefined();

    // A relative rule resolves to a different allowance on every cell, so the
    // column's setting -- 1e-12 -- is not what this cell was measured by.
    // Printing that would show the reader a number no cell was judged against,
    // and they could not check the verdict against the gap beside it.
    expect(Number(amount['Tolerance'])).toBeCloseTo(2520 * 1e-12, 15);
    expect(Number(amount['Tolerance'])).not.toBe(1e-12);
    expect(Number(amount['Delta'])).toBeLessThan(Number(amount['Tolerance']));

    // The cell downstream of it has no relative rule of its own, so it is
    // judged by the absolute default and says so. The two columns of one table
    // can be measured by different allowances, and the ledger distinguishes.
    const tax = within.find((r) => r['Column'] === 'Tax')!;
    expect(Number(tax['Tolerance'])).toBe(0.001);
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
