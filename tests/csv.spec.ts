import { test, expect } from '@playwright/test';
import { verifySheet, readSheet } from '../src/verify.js';
import { writeCsv } from './fixtures.js';

const SPEC = { keyColumns: ['PolicyId'] };

test.describe('CSV typing', () => {
  test('numeric-looking values become numbers so tolerance can apply', async () => {
    const f = await writeCsv('t1.csv');
    const m = await readSheet(f, SPEC);
    expect(m.rows.get('P-1001')!['Sum Insured']!.value).toBe(120000);
    expect(m.rows.get('P-1001')!['Rate']!.value).toBe(0.021);
  });

  test('key columns stay text, so leading zeros survive', async () => {
    const m = await readSheet(await writeCsv('t2.csv'), SPEC);
    expect(m.rows.get('P-1001')!['PolicyId']!.value).toBe('P-1001');
  });

  test("'none' keeps everything as text", async () => {
    const m = await readSheet(await writeCsv('t3.csv'), { ...SPEC, csv: { numeric: 'none' } });
    expect(m.rows.get('P-1001')!['Sum Insured']!.value).toBe('120000');
  });

  test('thousands separators are a real difference, not silently equal', async () => {
    const plain = await writeCsv('t4-plain.csv');
    const grouped = await writeCsv('t4-grouped.csv', { thousands: true });
    const d = await verifySheet(plain, grouped, SPEC);
    // "120,000" cannot be parsed as a number, so it surfaces as a type change
    // rather than passing as equal.
    expect(d.values.length + d.types.length).toBeGreaterThan(0);
  });
});

test.describe('CSV dialect', () => {
  test('delimiter is detected when not specified', async () => {
    const m = await readSheet(await writeCsv('t5.csv', { delimiter: ';' }), SPEC);
    expect(m.dialect?.delimiter).toBe(';');
    expect(m.rows.size).toBe(5);
  });

  test('BOM and line-ending drift is detected but not fatal by default', async () => {
    const a = await writeCsv('t6-a.csv', { bom: false, crlf: false });
    const b = await writeCsv('t6-b.csv', { bom: true, crlf: true });
    const d = await verifySheet(a, b, SPEC);
    expect(d.ok).toBe(true);          // the data is identical
    expect(d.errors).toEqual([]);
  });

  test('strictDialect turns encoding drift into a failure', async () => {
    const a = await writeCsv('t7-a.csv', { bom: false, crlf: false });
    const b = await writeCsv('t7-b.csv', { bom: true, crlf: true });
    const d = await verifySheet(a, b, { ...SPEC, csv: { strictDialect: true } });
    expect(d.ok).toBe(false);
    expect(d.errors.join(' ')).toContain('BOM added');
    expect(d.errors.join(' ')).toContain('line endings changed');
  });

  test('a BOM does not corrupt the first header name', async () => {
    const m = await readSheet(await writeCsv('t8.csv', { bom: true }), SPEC);
    expect(m.headers[0]).toBe('PolicyId');
  });
});

test.describe('CSV comparison', () => {
  test('value drift is caught with the same keyed alignment', async () => {
    const a = await writeCsv('t9-a.csv');
    const b = await writeCsv('t9-b.csv', { drift: true });
    const d = await verifySheet(a, b, SPEC);
    expect(d.values).toHaveLength(1);
    expect(d.values[0]).toMatchObject({ key: 'P-1003', column: 'Sum Insured', base: 240000, next: 249000 });
  });
});
