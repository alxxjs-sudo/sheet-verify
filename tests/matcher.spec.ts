import { test } from '@playwright/test';
import { expect } from '../src/matcher.js';
import { access, rm, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { buildWorkbook, buildMultiSheet, DIR } from './fixtures.js';

const SPEC = { keyColumns: ['PolicyId'] };

const WB_SPEC = {
  sheets: {
    Policies: { keyColumns: ['PolicyId'] },
    Premiums: { keyColumns: ['PolicyId', 'Period'] },
    Regions: { keyColumns: ['Region'] },
  },
};

test.describe('toMatchSheetBaseline', () => {
  test('passes when the output matches its baseline', async () => {
    const baseline = await buildWorkbook('m1-baseline.xlsx');
    const actual = await buildWorkbook('m1-actual.xlsx');
    await expect(actual).toMatchSheetBaseline(baseline, SPEC);
  });

  test('passes through a pure column insertion when schema is not strict', async () => {
    const baseline = await buildWorkbook('m2-baseline.xlsx');
    const actual = await buildWorkbook('m2-actual.xlsx', { insertPremium: true });
    await expect(actual).toMatchSheetBaseline(baseline, SPEC);
  });

  test('fails that same insertion under strictSchema', async () => {
    const baseline = await buildWorkbook('m3-baseline.xlsx');
    const actual = await buildWorkbook('m3-actual.xlsx', { insertPremium: true });
    await expect(actual).not.toMatchSheetBaseline(baseline, { ...SPEC, strictSchema: true });
  });

  test('fails on a real defect and names it in the message', async () => {
    const baseline = await buildWorkbook('m4-baseline.xlsx');
    const actual = await buildWorkbook('m4-actual.xlsx', { rateDrift: { 'P-1005': 0.12 } });

    let message = '';
    try {
      await expect(actual).toMatchSheetBaseline(baseline, SPEC);
    } catch (e) {
      message = (e as Error).message;
    }
    expect(message).toContain('FORMULA CHANGES');
    expect(message).toContain('P-1005');
    expect(message).toContain('UPDATE_SHEET_BASELINE=1');
  });

  test('creates a missing baseline on first run', async () => {
    const baseline = join(DIR, 'm5-created.xlsx');
    await rm(baseline, { force: true });
    const actual = await buildWorkbook('m5-actual.xlsx');

    await expect(actual).toMatchSheetBaseline(baseline, SPEC);
    await expect(access(baseline)).resolves.toBeUndefined();
  });

  test('refuses to invent a baseline when createMissingBaseline is off', async () => {
    const baseline = join(DIR, 'm6-absent.xlsx');
    await rm(baseline, { force: true });
    const actual = await buildWorkbook('m6-actual.xlsx');

    await expect(actual).not.toMatchSheetBaseline(baseline, { ...SPEC, createMissingBaseline: false });
  });

  test('re-blesses the baseline when asked, and then matches', async () => {
    const baseline = await buildWorkbook('m7-baseline.xlsx');
    const actual = await buildWorkbook('m7-actual.xlsx', { rateDrift: { 'P-1005': 0.12 } });

    // Fails first.
    await expect(actual).not.toMatchSheetBaseline(baseline, SPEC);
    // Accept the change explicitly.
    await expect(actual).toMatchSheetBaseline(baseline, { ...SPEC, updateBaseline: true });
    // Now it is the baseline.
    await expect(actual).toMatchSheetBaseline(baseline, SPEC);
    expect((await readFile(baseline)).length).toBeGreaterThan(0);
  });

  test('reports a missing actual file clearly', async () => {
    const baseline = await buildWorkbook('m8-baseline.xlsx');
    let message = '';
    try {
      await expect(join(DIR, 'does-not-exist.xlsx')).toMatchSheetBaseline(baseline, SPEC);
    } catch (e) {
      message = (e as Error).message;
    }
    expect(message).toContain('actual file not found');
  });

  test('surfaces invariant failures through the matcher', async () => {
    const baseline = await buildWorkbook('m9-baseline.xlsx');
    const actual = await buildWorkbook('m9-actual.xlsx');
    const { inRange } = await import('../src/invariants.js');

    await expect(actual).not.toMatchSheetBaseline(baseline, {
      ...SPEC, invariants: [inRange('Rate', 0, 0.001)],
    });
  });
});

test.describe('toMatchWorkbookBaseline', () => {
  test('passes when every sheet matches', async () => {
    const baseline = await buildMultiSheet('w1-baseline.xlsx');
    const actual = await buildMultiSheet('w1-actual.xlsx');
    await expect(actual).toMatchWorkbookBaseline(baseline, WB_SPEC);
  });

  test('names the offending sheet in the failure message', async () => {
    const baseline = await buildMultiSheet('w2-baseline.xlsx');
    const actual = await buildMultiSheet('w2-actual.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    let message = '';
    try {
      await expect(actual).toMatchWorkbookBaseline(baseline, WB_SPEC);
    } catch (e) {
      message = (e as Error).message;
    }
    expect(message).toContain('SHEET "Premiums"');
    expect(message).toContain('P-1003');
    expect(message).toContain('UPDATE_SHEET_BASELINE=1');
    // The clean sheets must not be dragged into the failure detail.
    expect(message).not.toContain('SHEET "Regions"');
  });

  test('passes a column inserted on one sheet, and fails it under strictSchema', async () => {
    const baseline = await buildMultiSheet('w3-baseline.xlsx');
    const actual = await buildMultiSheet('w3-actual.xlsx', { insertPremium: true });

    await expect(actual).toMatchWorkbookBaseline(baseline, WB_SPEC);
    await expect(actual).not.toMatchWorkbookBaseline(baseline, {
      ...WB_SPEC, defaults: { strictSchema: true },
    });
  });

  test('creates a missing baseline on first run, then re-blesses on demand', async () => {
    const baseline = join(DIR, 'w4-created.xlsx');
    await rm(baseline, { force: true });
    const actual = await buildMultiSheet('w4-actual.xlsx');

    await expect(actual).toMatchWorkbookBaseline(baseline, WB_SPEC);
    await expect(access(baseline)).resolves.toBeUndefined();

    const drifted = await buildMultiSheet('w4-drifted.xlsx', {
      premiumDrift: { 'P-1001|2026-07': 42 },
    });
    await expect(drifted).not.toMatchWorkbookBaseline(baseline, WB_SPEC);
    await expect(drifted).toMatchWorkbookBaseline(baseline, { ...WB_SPEC, updateBaseline: true });
    await expect(drifted).toMatchWorkbookBaseline(baseline, WB_SPEC);
  });

  test('fails when a sheet has vanished from the output', async () => {
    const baseline = await buildMultiSheet('w5-baseline.xlsx');
    const actual = await buildMultiSheet('w5-actual.xlsx', { sheets: ['Policies', 'Premiums'] });

    let message = '';
    try {
      await expect(actual).toMatchWorkbookBaseline(baseline, WB_SPEC);
    } catch (e) {
      message = (e as Error).message;
    }
    expect(message).toContain('SHEETS REMOVED');
    expect(message).toContain('Regions');
  });
});
