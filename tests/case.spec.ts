import { test, expect } from '@playwright/test';
import { readFile, rm, access } from 'node:fs/promises';
import { join } from 'node:path';
import { runCase, formatLedger, runWorkbook } from '../src/index.js';
import type { CaseOptions } from '../src/index.js';
import { buildMultiSheet, DIR } from './fixtures.js';

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

/** Parses the ledger by header name, so adding a column cannot break a test. */
const parseLedger = (csv: string): Record<string, string>[] => {
  const [header, ...body] = rows(csv);
  const names = header!.split(',');
  return body.map((line) =>
    Object.fromEntries(line.split(',').map((v, i) => [names[i]!, v])),
  );
};

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

  test('the cell ledger records matches, not only differences', async () => {
    const dir = await caseDir('ledger-all');
    const golden = await buildMultiSheet('case-led-golden.xlsx');
    const actual = await buildMultiSheet('case-led-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);

    const csv = await readFile(r.files.cells, 'utf8');
    const lines = rows(csv);

    expect(lines[0]).toBe(
      'sheet,table,row_key,column,status,root_cause,baseline_address,actual_address,' +
      'baseline_value,actual_value,delta,tolerance,baseline_formula,actual_formula',
    );
    // Every cell in scope, so matches dominate the file.
    expect(lines.filter((l) => l.includes(',match,')).length).toBeGreaterThan(50);

    // The drifted Amount, and the Tax formula that reads it.
    const differing = lines.filter((l) => l.includes(',value-differs,'));
    expect(differing).toHaveLength(2);

    const cause = differing.filter((l) => l.includes(',value-differs,yes,'));
    expect(cause).toHaveLength(1);
    expect(cause[0]).toContain('P-1003');
    expect(cause[0]).toContain('Amount');
    expect(cause[0]).toContain('9999');

    // The consequence is recorded, but marked as not the cause.
    const consequence = differing.filter((l) => l.includes(',value-differs,no,'));
    expect(consequence).toHaveLength(1);
    expect(consequence[0]).toContain('Tax');
  });

  test('the ledger can be narrowed to differences only', async () => {
    const dir = await caseDir('ledger-diff');
    const golden = await buildMultiSheet('case-nar-golden.xlsx');
    const actual = await buildMultiSheet('case-nar-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, { ...SPEC, cellLedger: 'differences' });

    const lines = rows(await readFile(r.files.cells, 'utf8'));
    expect(lines.some((l) => l.includes(',match,'))).toBe(false);
    // Header, the drifted cell, and the Tax formula downstream of it.
    expect(lines).toHaveLength(3);
    expect(lines.filter((l) => l.includes(',value-differs,'))).toHaveLength(2);
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
      .filter((r) => r.status === 'within-tolerance');

    expect(within).toHaveLength(1);
    expect(within[0]!.tolerance).toBe('1');
    expect(within[0]!.delta).toBe('0.5');
    expect(within[0]!.column).toBe('Amount');
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
