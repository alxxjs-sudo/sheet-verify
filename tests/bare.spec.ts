import { test, expect } from '@playwright/test';
import { readFile, writeFile, copyFile } from 'node:fs/promises';
import { join } from 'node:path';
import {
  stripCachedValues, cachedValueState, makeBare, openWorkbook, runWorkbook, sweep,
} from '../src/index.js';
import type { WorkbookSpec } from '../src/index.js';
import { buildSweepWorkbook, DIR } from './fixtures.js';

/**
 * Restoring a report Excel has recalculated. The fixtures stand in for the two
 * real states: `cacheResults` writes a file as Excel saves one, and the default
 * writes it as the generator does.
 */

const SPEC: WorkbookSpec = {
  defaults: { requireCachedValues: false },
  sheets: { Policies: { keyColumns: ['PolicyId'] } },
};

test.describe('cachedValueState', () => {
  test('tells a generated file from one that has been through Excel', async () => {
    const generated = await buildSweepWorkbook('bare-gen.xlsx');
    const saved = await buildSweepWorkbook('bare-saved.xlsx', { cacheResults: true });

    const g = await cachedValueState(await readFile(generated));
    const s = await cachedValueState(await readFile(saved));

    expect(g.cached).toBe(0);
    expect(g.bare).toBe(5);
    expect(s.cached).toBe(5);
    expect(s.bare).toBe(0);
  });
});

test.describe('stripCachedValues', () => {
  test('removes every stored result and says how many', async () => {
    const saved = await buildSweepWorkbook('bare-strip.xlsx', { cacheResults: true });

    const result = await stripCachedValues(await readFile(saved));

    expect(result.stripped).toBe(5);
    const after = await cachedValueState(result.buffer);
    expect(after.cached).toBe(0);
    expect(after.bare).toBe(5);
  });

  test('asks Excel to recalculate on open, so the blanks are not taken at face value', async () => {
    const saved = await buildSweepWorkbook('bare-calcpr.xlsx', { cacheResults: true });

    const result = await stripCachedValues(await readFile(saved));

    expect((await cachedValueState(result.buffer)).fullCalcOnLoad).toBe(true);
  });

  test('keeps the formulas themselves intact', async () => {
    const saved = await buildSweepWorkbook('bare-keep.xlsx', { cacheResults: true });

    const result = await stripCachedValues(await readFile(saved));
    const path = join(DIR, 'bare-keep-out.xlsx');
    await writeFile(path, result.buffer);

    // The whole point is that only the result goes. A stripped file that no
    // longer opens, or whose formulas moved, would be worse than the problem.
    const wb = await openWorkbook(path);
    const ws = wb.getWorksheet('Policies')!;
    expect(ws.getCell('E2').formula).toBe('C2*D2');
    expect(ws.getCell('A2').value).toBe('P-1001');
    expect(wb.worksheets.map((w) => w.name)).toEqual(['Policies', 'Notes']);
  });

  test('restores fullCalcOnLoad when a rewrite dropped it, and still opens', async () => {
    // Anything that rewrites a workbook can drop the flag -- ExcelJS does. The
    // file then shows blanks to whoever opens it, because nothing tells Excel
    // to work the numbers out.
    const generated = await buildSweepWorkbook('bare-calc-src.xlsx');
    const path = join(DIR, 'bare-calc.xlsx');
    const wb = await openWorkbook(generated);
    await wb.xlsx.writeFile(path);
    expect((await cachedValueState(await readFile(path))).fullCalcOnLoad).toBe(false);

    const result = await makeBare(path);

    expect(result?.restoredFullCalc).toBe(true);
    expect((await cachedValueState(await readFile(path))).fullCalcOnLoad).toBe(true);
    // The rewrite must leave valid XML: an earlier attempt produced
    // `<calcPr …/ fullCalcOnLoad="1">` and the file no longer opened.
    const reopened = await openWorkbook(path);
    expect(reopened.worksheets.map((w) => w.name)).toEqual(['Policies', 'Notes']);
  });

  test('writes a compressed package, not a stored one', async () => {
    // JSZip stores parts uncompressed unless told otherwise. Nothing breaks --
    // Excel reads both -- so this only shows up as files growing several times
    // over every time they are rewritten.
    const saved = await buildSweepWorkbook('bare-zip.xlsx', { cacheResults: true });
    const before = (await readFile(saved)).length;

    const result = await stripCachedValues(await readFile(saved));

    // Stripping removes content, so the result must not be larger than the
    // original; uncompressed it would be several times the size.
    expect(result.buffer.length).toBeLessThanOrEqual(before);
  });

  test('leaves an already-bare file alone', async () => {
    const generated = await buildSweepWorkbook('bare-noop.xlsx');

    const result = await stripCachedValues(await readFile(generated));

    expect(result.stripped).toBe(0);
    expect((await cachedValueState(result.buffer)).bare).toBe(5);
  });
});

test.describe('makeBare', () => {
  test('rewrites in place and reports nothing to do the second time', async () => {
    const saved = await buildSweepWorkbook('bare-inplace-src.xlsx', { cacheResults: true });
    const path = join(DIR, 'bare-inplace.xlsx');
    await copyFile(saved, path);

    const first = await makeBare(path);
    expect(first?.stripped).toBe(5);

    // Idempotent: a second run finds nothing and does not rewrite the file.
    expect(await makeBare(path)).toBeNull();
  });
});

test.describe('the problem it solves', () => {
  test('a skewed pair reports every formula as a value change', async () => {
    const golden = await buildSweepWorkbook('bare-skew-g.xlsx');
    const actual = await buildSweepWorkbook('bare-skew-a.xlsx', { cacheResults: true });

    const { diff } = await runWorkbook(golden, actual, SPEC);
    const policies = diff.sheets.find((x) => x.sheet === 'Policies')!;

    // Five formulas, five phantom differences, and not one thing changed.
    expect(policies.diff!.values.length).toBe(5);
    expect(policies.diff!.values.every((v) => v.base === null)).toBe(true);
  });

  test('stripping the report first leaves only the real edit', async () => {
    const golden = await buildSweepWorkbook('bare-fix-g.xlsx');
    // A report opened in Excel to plant one edit: it comes back cached.
    const actual = await buildSweepWorkbook('bare-fix-a.xlsx', {
      cacheResults: true,
      sumDrift: { 'P-1002': 90000 },
    });
    await makeBare(actual);

    const { diff, compared } = await runWorkbook(golden, actual, SPEC);
    const policies = diff.sheets.find((x) => x.sheet === 'Policies')!;

    // Exactly the edit, and nothing else.
    expect(policies.diff!.values.length).toBe(1);
    expect(policies.diff!.values[0]!.column).toBe('Sum Insured');
    expect(policies.diff!.values[0]!.base).toBe(85000);
    expect(policies.diff!.values[0]!.next).toBe(90000);

    // And layer 2 agrees: one cell in the whole file.
    const s = await sweep(golden, actual, compared);
    expect(s.totalDifferences).toBe(1);
  });
});
