import { test, expect } from '@playwright/test';
import { verifySheet } from '../src/verify.js';
import { formatReport, summarize } from '../src/report.js';
import type { SheetSpec } from '../src/types.js';
import { buildWorkbook, makeSharedFormulas, DIR } from './fixtures.js';
import { mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import ExcelJS from 'exceljs';

const SPEC: SheetSpec = { keyColumns: ['PolicyId'] };

/** A one-row sheet under whatever headings a test needs. */
async function headedWorkbook(name: string, headers: string[]): Promise<string> {
  await mkdir(DIR, { recursive: true });
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.addRow(headers);
  ws.addRow(['US - Florida', 6642514.759, 9618538.628]);
  const path = join(DIR, name);
  await wb.xlsx.writeFile(path);
  return path;
}


test.describe('the release-drift scenario', () => {
  test('an inserted column is a schema change, not data churn', async () => {
    const base = await buildWorkbook('c1-base.xlsx');
    const next = await buildWorkbook('c1-next.xlsx', {
      insertPremium: true,
      valueDrift: { 'P-1003': 249000 },
      rateDrift: { 'P-1005': 0.12 },
    });

    const d = await verifySheet(base, next, SPEC);

    // Schema drift is reported once, in its own layer.
    expect(d.schema.added).toEqual(['Premium']);
    expect(d.schema.removed).toEqual([]);
    expect(d.schema.moved.map((m) => m.column)).toEqual(['Rate', 'Annual Cost', 'Commission']);

    // Rows are matched by key, so the population is unchanged.
    expect(d.rows.added).toEqual([]);
    expect(d.rows.removed).toEqual([]);
    expect(d.rows.compared).toBe(5);

    // Exactly the planted formula defect, and nothing from the shift.
    expect(d.formulas).toHaveLength(1);
    expect(d.formulas[0]).toMatchObject({
      key: 'P-1005', column: 'Commission',
      baseA1: 'F6*0.1', nextA1: 'G6*0.12',
    });

    // The value drift and its two downstream cells.
    expect(d.values).toHaveLength(4);
    const roots = d.values.filter((v) => v.rootCause);
    expect(roots.map((v) => `${v.key}/${v.column}`).sort())
      .toEqual(['P-1003/Sum Insured', 'P-1005/Commission']);
    expect(d.ok).toBe(false);
  });

  test('root-cause grouping separates the cause from its cascade', async () => {
    const base = await buildWorkbook('c2-base.xlsx');
    const next = await buildWorkbook('c2-next.xlsx', { valueDrift: { 'P-1003': 249000 } });
    const d = await verifySheet(base, next, SPEC);

    // One edit, three changed cells: the source plus two dependent formulas.
    expect(d.values).toHaveLength(3);
    const roots = d.values.filter((v) => v.rootCause);
    expect(roots).toHaveLength(1);
    expect(roots[0]!.column).toBe('Sum Insured');

    const cascaded = d.values.filter((v) => !v.rootCause).map((v) => v.column).sort();
    expect(cascaded).toEqual(['Annual Cost', 'Commission']);
  });

  test('identical files compare clean', async () => {
    const a = await buildWorkbook('c3-a.xlsx');
    const b = await buildWorkbook('c3-b.xlsx');
    const d = await verifySheet(a, b, SPEC);
    expect(d.ok).toBe(true);
    expect(summarize(d)).toBe('identical');
  });
});

test.describe('formula normalisation modes', () => {
  test('header mode isolates the real defect where a1 and r1c1 do not', async () => {
    const base = await buildWorkbook('c4-base.xlsx');
    const next = await buildWorkbook('c4-next.xlsx', {
      insertPremium: true, rateDrift: { 'P-1005': 0.12 },
    });

    const counts: Record<string, number> = {};
    for (const mode of ['a1', 'r1c1', 'header'] as const) {
      const d = await verifySheet(base, next, { ...SPEC, formulaMode: mode });
      counts[mode] = d.formulas.length;
    }

    // A1 flags every shifted formula; R1C1 still flags Annual Cost because the
    // new column landed between its two operands; header resolution flags none.
    expect(counts).toEqual({ a1: 10, r1c1: 6, header: 1 });
  });
});

test.describe('shared formulas', () => {
  test('filled-down formulas are translated, not read as the master address', async () => {
    // Real Excel stores a filled-down column once and points the rest at it.
    const base = await buildWorkbook('c5-base.xlsx');
    await makeSharedFormulas(base, 'F', 2, 6, 0);
    const next = await buildWorkbook('c5-next.xlsx', { rateDrift: { 'P-1005': 0.12 } });
    await makeSharedFormulas(next, 'F', 2, 6, 0);

    const d = await verifySheet(base, next, SPEC);

    // If shared formulas were read as their master's address, every row would
    // normalise to the same string and the Commission defect would vanish.
    expect(d.formulas).toHaveLength(1);
    expect(d.formulas[0]!.key).toBe('P-1005');
    expect(d.formulas[0]!.column).toBe('Commission');
  });

  test('a shared column matches an unshared one with the same logic', async () => {
    const shared = await buildWorkbook('c6-shared.xlsx');
    await makeSharedFormulas(shared, 'F', 2, 6, 0);
    const plain = await buildWorkbook('c6-plain.xlsx');

    const d = await verifySheet(shared, plain, SPEC);
    expect(d.formulas).toHaveLength(0);
    expect(d.ok).toBe(true);
  });
});

test.describe('row population', () => {
  test('added and removed rows are reported by key', async () => {
    const base = await buildWorkbook('c7-base.xlsx');
    const next = await buildWorkbook('c7-next.xlsx', {
      dropRows: ['P-1002'],
      extraRows: [{ id: 'P-1099', holder: 'Todorov', region: 'Stara Zagora', sumInsured: 45000, rate: 0.02 }],
    });
    const d = await verifySheet(base, next, SPEC);

    expect(d.rows.removed).toEqual(['P-1002']);
    expect(d.rows.added).toEqual(['P-1099']);
    expect(d.values).toHaveLength(0); // surviving rows are untouched
  });

  test('reordering rows produces no differences at all', async () => {
    const base = await buildWorkbook('c8-base.xlsx');
    const next = await buildWorkbook('c8-next.xlsx', {
      dropRows: ['P-1001'],
      extraRows: [{ id: 'P-1001', holder: 'Ivanov', region: 'Sofia', sumInsured: 120000, rate: 0.021 }],
    });
    const d = await verifySheet(base, next, SPEC);
    expect(d.values).toHaveLength(0);
    expect(d.rows.added).toEqual([]);
    expect(d.rows.removed).toEqual([]);
    expect(d.ok).toBe(true);
  });
});

test.describe('type and tolerance handling', () => {
  test('a number written as text is caught even though it renders the same', async () => {
    const base = await buildWorkbook('c9-base.xlsx');
    const next = await buildWorkbook('c9-next.xlsx', { textDrift: ['P-1002'] });
    const d = await verifySheet(base, next, SPEC);

    expect(d.values).toHaveLength(0);       // "85000" equals 85000 once rendered
    expect(d.types).toHaveLength(1);        // but the type changed
    expect(d.types[0]).toMatchObject({ key: 'P-1002', column: 'Sum Insured', nextKind: 'string' });
    expect(d.ok).toBe(false);
  });

  test('per-column tolerance absorbs float noise without hiding real drift', async () => {
    const base = await buildWorkbook('c10-base.xlsx');
    const next = await buildWorkbook('c10-next.xlsx', { valueDrift: { 'P-1003': 240000.004 } });

    const loose = await verifySheet(base, next, { ...SPEC, tolerance: { 'Sum Insured': 0.01, '*': 0.01 } });
    expect(loose.values).toHaveLength(0);

    const strict = await verifySheet(base, next, SPEC);
    expect(strict.values.length).toBeGreaterThan(0);
  });

  test('a relative tolerance scales with the number, where one flat figure cannot', async () => {
    // One report, two magnitudes. The big cell drifts in the fifteenth
    // significant digit -- the last one Excel stores, so a total rebuilt in a
    // different order lands there and nothing has changed. The small cell moves
    // by a third of itself, which is a change by any reading.
    const base = await buildWorkbook('c30-base.xlsx', {
      valueDrift: { 'P-1001': 126339393111.699, 'P-1004': 0.00026 },
    });
    const next = await buildWorkbook('c30-next.xlsx', {
      valueDrift: { 'P-1001': 126339393111.6952, 'P-1004': 0.00035 },
    });

    // A flat tolerance gets both wrong, and in opposite directions. The noise
    // on the big number is 0.0038 and so exceeds 0.001; a 35% move on the small
    // number is 0.00009 and slips under it.
    const flat = await verifySheet(base, next, { ...SPEC, tolerance: 0.001 });
    const flatKeys = new Set(flat.values.map((v) => v.key));
    expect(flatKeys.has('P-1001')).toBe(true);
    expect(flatKeys.has('P-1004')).toBe(false);

    // Judged in proportion as well, each lands the right way round: 3.0e-14 of
    // its value is forgiven, 0.26 of its value is not.
    const scaled = await verifySheet(base, next, {
      ...SPEC, tolerance: 1e-9, relativeTolerance: 1e-12,
    });
    const scaledKeys = new Set(scaled.values.map((v) => v.key));
    expect(scaledKeys.has('P-1001')).toBe(false);
    expect(scaledKeys.has('P-1004')).toBe(true);
  });

  test('a relative tolerance does not forgive a difference someone made', async () => {
    const base = await buildWorkbook('c31-base.xlsx');
    const next = await buildWorkbook('c31-next.xlsx', { valueDrift: { 'P-1003': 249000 } });
    const d = await verifySheet(base, next, {
      ...SPEC, tolerance: 1e-9, relativeTolerance: 1e-12,
    });

    // 240000 -> 249000 is 3.75% of the value: ten orders of magnitude above
    // anything recalculation produces, and nowhere near forgivable.
    expect(d.values.some((v) => v.key === 'P-1003' && v.column === 'Sum Insured')).toBe(true);
    expect(d.ok).toBe(false);
  });

  test('left unset it changes nothing, so an existing config compares as before', async () => {
    const base = await buildWorkbook('c32-base.xlsx');
    const next = await buildWorkbook('c32-next.xlsx', {
      valueDrift: { 'P-1003': 240000.004 },
    });

    // The same drift, judged by the same absolute tolerance, with the relative
    // rule left out and then written down as its default. Neither reading may
    // differ from the other: a default that quietly forgave anything would be
    // the hidden slack this rule was built to not be.
    const unset = await verifySheet(base, next, { ...SPEC, tolerance: 0.01 });
    const zero = await verifySheet(base, next, {
      ...SPEC, tolerance: 0.01, relativeTolerance: 0,
    });
    expect(zero.values).toHaveLength(unset.values.length);
    expect(unset.values).toHaveLength(0);

    const strict = await verifySheet(base, next, { ...SPEC, tolerance: 0 });
    expect(strict.values.length).toBeGreaterThan(0);
  });

  test('ignored columns are excluded from comparison', async () => {
    const base = await buildWorkbook('c11-base.xlsx');
    const next = await buildWorkbook('c11-next.xlsx', { valueDrift: { 'P-1003': 249000 } });
    const d = await verifySheet(base, next, {
      ...SPEC, ignoreColumns: ['Sum Insured', 'Annual Cost', 'Commission'],
    });
    expect(d.values).toHaveLength(0);
    expect(d.ok).toBe(true);
  });
});

test.describe('comparison integrity', () => {
  test('formulas without cached values are refused, not silently passed', async () => {
    const base = await buildWorkbook('c12-base.xlsx');
    const next = await buildWorkbook('c12-next.xlsx', { omitCachedResults: true });
    const d = await verifySheet(base, next, SPEC);

    expect(d.errors.join(' ')).toContain('no cached value');
    expect(d.ok).toBe(false);
  });

  test('opting out of the cached-value check leaves formula comparison working', async () => {
    const base = await buildWorkbook('c13-base.xlsx');
    const next = await buildWorkbook('c13-next.xlsx', {
      omitCachedResults: true, rateDrift: { 'P-1005': 0.12 },
    });
    const d = await verifySheet(base, next, { ...SPEC, requireCachedValues: false });

    expect(d.errors).toEqual([]);
    expect(d.formulas).toHaveLength(1);   // logic drift still caught
  });

  test('a missing key column fails loudly instead of comparing positionally', async () => {
    const base = await buildWorkbook('c14-base.xlsx');
    const next = await buildWorkbook('c14-next.xlsx');
    const d = await verifySheet(base, next, { keyColumns: ['ContractRef'] });
    expect(d.errors.join(' ')).toContain('key column "ContractRef" not found');
    expect(d.ok).toBe(false);
  });

  test('a spec without keyColumns is rejected at the call site', async () => {
    const base = await buildWorkbook('c15-base.xlsx');
    await expect(verifySheet(base, base, { keyColumns: [] } as any))
      .rejects.toThrow(/keyColumns.*required/s);
  });
});

test.describe('report', () => {
  test('leads with the defect and keeps the schema change separate', async () => {
    const base = await buildWorkbook('c16-base.xlsx');
    const next = await buildWorkbook('c16-next.xlsx', {
      insertPremium: true, valueDrift: { 'P-1003': 249000 }, rateDrift: { 'P-1005': 0.12 },
    });
    const d = await verifySheet(base, next, SPEC);
    const text = formatReport(d);

    expect(text).toContain('FORMULA CHANGES (1)');
    expect(text).toContain('P-1005 · Commission');
    expect(text).toContain('SCHEMA CHANGES');
    expect(text).toContain('+ column "Premium"');
    // The cause is listed; the two cascaded cells are not, by default.
    expect(text).toContain('P-1003 · Sum Insured');
    expect(text).not.toContain('CASCADED');
    expect(text.indexOf('FORMULA CHANGES')).toBeLessThan(text.indexOf('SCHEMA CHANGES'));
  });

  test('a numeric column heading that drifted is the same column, not two', async () => {
    // From a real pro-forma report: columns headed by computed figures whose
    // last digits move when the sheet recalculates. Compared as text those read
    // as one column removed and another added, and every cell beneath them was
    // then reported twice -- burying the real findings under hundreds of rows.
    // A tolerance would have forgiven the same drift instantly had the number
    // been a value rather than a name.
    const golden = await headedWorkbook('numhead-golden.xlsx',
      ['Region', '788321.400221', '33792307.66114401']);
    const actual = await headedWorkbook('numhead-actual.xlsx',
      ['Region', '788321.4002209998', '33792307.661144']);

    const d = await verifySheet(golden, actual, { keyColumns: ['Region'] });
    expect(d.schema.removed).toEqual([]);
    expect(d.schema.added).toEqual([]);
    // Region plus the two numeric headings: three columns, none of them added
    // or removed.
    expect(d.schema.compared).toHaveLength(3);
    expect(d.ok).toBe(true);
  });

  test('headings that are genuinely different numbers stay different columns', async () => {
    const golden = await headedWorkbook('numhead-diff-g.xlsx', ['Region', '1000', '2000']);
    const actual = await headedWorkbook('numhead-diff-a.xlsx', ['Region', '1000', '3000']);

    const d = await verifySheet(golden, actual, { keyColumns: ['Region'] });
    expect(d.schema.removed).toEqual(['2000']);
    expect(d.schema.added).toEqual(['3000']);
  });
});
