import { test, expect } from '@playwright/test';
import { verifySheet } from '../src/verify.js';
import * as inv from '../src/invariants.js';
import { buildWorkbook } from './fixtures.js';

const SPEC = { keyColumns: ['PolicyId'] };

test.describe('baseline-free checks', () => {
  test('derived() catches a wrong calculation both files agree on', async () => {
    // Both sides carry the same bad rate, so a golden-file diff sees nothing.
    const base = await buildWorkbook('i1-base.xlsx', { rateDrift: { 'P-1005': 0.12 } });
    const next = await buildWorkbook('i1-next.xlsx', { rateDrift: { 'P-1005': 0.12 } });

    const plain = await verifySheet(base, next, SPEC);
    expect(plain.ok).toBe(true);   // the blind spot

    const checked = await verifySheet(base, next, {
      ...SPEC,
      invariants: [
        inv.derived('Commission', (_row, num) => num('Annual Cost') * 0.1, 1e-6),
      ],
    });
    expect(checked.ok).toBe(false);
    expect(checked.invariants).toHaveLength(1);
    expect(checked.invariants[0]).toMatchObject({ key: 'P-1005', column: 'Commission' });
    expect(checked.invariants[0]!.detail).toContain('recomputed');
  });

  test('inRange flags out-of-bounds values', async () => {
    const f = await buildWorkbook('i2.xlsx');
    const d = await verifySheet(f, f, { ...SPEC, invariants: [inv.inRange('Rate', 0, 0.02)] });
    // Rates above 0.02: P-1001 (0.021), P-1003 (0.025), P-1005 (0.029).
    expect(d.invariants.map((x) => x.key).sort()).toEqual(['P-1001', 'P-1003', 'P-1005']);
  });

  test('notBlank and unique hold on a well-formed sheet', async () => {
    const f = await buildWorkbook('i3.xlsx');
    const d = await verifySheet(f, f, {
      ...SPEC,
      invariants: [inv.notBlank('PolicyId', 'Holder'), inv.unique('PolicyId')],
    });
    expect(d.invariants).toEqual([]);
    expect(d.ok).toBe(true);
  });

  test('formulasHaveResults reports uncached formula cells', async () => {
    const f = await buildWorkbook('i4.xlsx', { omitCachedResults: true });
    const d = await verifySheet(f, f, {
      ...SPEC, requireCachedValues: false, invariants: [inv.formulasHaveResults()],
    });
    expect(d.invariants.length).toBe(10); // 5 rows x 2 formula columns
  });

  test('a throwing invariant is reported, not swallowed', async () => {
    const f = await buildWorkbook('i5.xlsx');
    const d = await verifySheet(f, f, {
      ...SPEC,
      invariants: [{ name: 'boom', check() { throw new Error('kaboom'); } }],
    });
    expect(d.invariants[0]!.detail).toContain('kaboom');
    expect(d.ok).toBe(false);
  });
});
