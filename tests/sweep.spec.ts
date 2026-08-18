import { test, expect } from '@playwright/test';
import { readFile, rm, access } from 'node:fs/promises';
import { join } from 'node:path';
import {
  runWorkbook, runCase, sweep, formatSweepReport, summarizeSweep, ledgerRows,
} from '../src/index.js';
import type { CaseOptions, SweepResult, WorkbookSpec } from '../src/index.js';
import { buildSweepWorkbook, DIR } from './fixtures.js';

/**
 * Layer 2 is a coverage audit, so almost every test here is the same shape:
 * plant a change, then assert whether layer 1 could have seen it.
 */

/**
 * Keys Policies only, so Notes is left uncompared -- which is what gives layer
 * 2 something to find. The positional fallback would otherwise compare Notes
 * too, and these tests are about the case where nothing checked a cell at all.
 *
 * The fixtures write formulas without cached results, as the real reports do,
 * so the check for them is off -- otherwise every case here fails on that
 * instead of on what it is testing.
 */
const SPEC: WorkbookSpec = {
  matchUnkeyedRowsByPosition: false,
  defaults: { requireCachedValues: false },
  sheets: { Policies: { keyColumns: ['PolicyId'] } },
};

const exists = (p: string) => access(p).then(() => true, () => false);

/** Runs both layers the way runCase does, and hands back layer 2. */
async function sweepOf(
  golden: string,
  actual: string,
  spec: WorkbookSpec = SPEC,
): Promise<SweepResult> {
  const { compared } = await runWorkbook(golden, actual, spec);
  return sweep(golden, actual, compared);
}

const at = (s: SweepResult, sheet: string, address: string) =>
  s.differences.find((d) => d.sheet === sheet && d.address === address);

test.describe('sweep', () => {
  test('finds nothing when the two files are identical', async () => {
    const golden = await buildSweepWorkbook('sw-same-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-same-actual.xlsx');

    const s = await sweepOf(golden, actual);

    expect(s.totalDifferences).toBe(0);
    expect(s.totalGaps).toBe(0);
    expect(summarizeSweep(s)).toBe('every cell identical');
    // It still had to look at the whole file to say so.
    expect(s.cellsSwept).toBeGreaterThan(0);
    expect(s.cellsCompared).toBeGreaterThan(0);
  });

  test('a change layer 1 checked is marked compared, not a gap', async () => {
    const golden = await buildSweepWorkbook('sw-cov-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-cov-actual.xlsx', {
      sumDrift: { 'P-1003': 249000 },
    });

    const s = await sweepOf(golden, actual);

    expect(s.totalDifferences).toBe(1);
    expect(s.totalGaps).toBe(0);
    expect(at(s, 'Policies', 'C4')?.status).toBe('compared');
    expect(summarizeSweep(s)).toContain('accounted for');
  });

  test('a change on an unkeyed sheet is a gap — the whole point', async () => {
    const golden = await buildSweepWorkbook('sw-gap-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-gap-actual.xlsx', { note: 'Extended' });

    // Layer 1 cannot key Notes, so it passes without ever looking.
    const { diff } = await runWorkbook(golden, actual, SPEC);
    expect(diff.sheets.find((x) => x.sheet === 'Notes')?.status).toBe('skipped');
    expect(diff.ok).toBe(true);

    const s = await sweepOf(golden, actual);

    expect(s.totalGaps).toBe(1);
    const cell = at(s, 'Notes', 'B3')!;
    expect(cell.status).toBe('gap');
    expect(cell.reason).toBe('sheet not compared');
    expect(cell.base).toBe('Standard');
    expect(cell.next).toBe('Extended');
  });

  test('positional matching closes the gap layer 2 would otherwise report', async () => {
    const golden = await buildSweepWorkbook('sw-posfix-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-posfix-actual.xlsx', { note: 'Extended' });

    // Same files, same change, with the fallback left on: Notes is compared by
    // position, so the change is a finding rather than a blind spot.
    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      sheets: {
        Policies: { keyColumns: ['PolicyId'] },
        // Notes has a title row above its headers, which detection works out
        // for itself; here it is stated. What it still has no key for is its
        // rows, which is the point.
        Notes: { headerRow: 2 },
      },
    };
    const { diff, compared } = await runWorkbook(golden, actual, spec);
    const notes = diff.sheets.find((x) => x.sheet === 'Notes')!;

    expect(notes.status).toBe('compared');
    expect(notes.diff!.values).toHaveLength(1);
    expect(diff.ok).toBe(false);

    const s = await sweep(golden, actual, compared);
    expect(s.totalGaps).toBe(0);
  });

  test('a positionally matched row is labelled with its own text, not just an ordinal', async () => {
    // "#3" tells a reviewer nothing. The row they are looking at is the one
    // that says "Coverage", and that text is sitting in the row already.
    const golden = await buildSweepWorkbook('sw-label-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-label-actual.xlsx', { note: 'Extended' });

    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      sheets: { Policies: { keyColumns: ['PolicyId'] }, Notes: { headerRow: 2 } },
    };
    const { compared } = await runWorkbook(golden, actual, spec);
    const rows = [...ledgerRows(compared, 'differences')];
    const note = rows.find((r) => r.sheet === 'Notes')!;

    expect(note.rowKey).toContain('#1');
    expect(note.rowKey).toContain('Coverage');
  });

  test('a change on a deliberately excluded sheet is not a gap either', async () => {
    // Being told there is a blind spot on a sheet you excluded on purpose is a
    // false alarm, and false alarms are what stop the headline being read.
    const golden = await buildSweepWorkbook('sw-ignsheet-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-ignsheet-actual.xlsx', { note: 'Extended' });

    const { compared } = await runWorkbook(golden, actual, {
      ...SPEC,
      ignoreSheets: ['Notes'],
    });
    const s = await sweep(golden, actual, compared, { ignoreSheets: ['Notes'] });

    expect(s.totalGaps).toBe(0);
    expect(s.totalExcluded).toBe(1);
    expect(at(s, 'Notes', 'B3')?.status).toBe('excluded');
  });

  test('a deliberately ignored column is excluded, not counted as a gap', async () => {
    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      sheets: { Policies: { keyColumns: ['PolicyId'], ignoreColumns: ['Run Stamp'] } },
    };
    const golden = await buildSweepWorkbook('sw-ign-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-ign-actual.xlsx', {
      stamp: '2026-08-17T11:30:00Z',
    });

    const { diff } = await runWorkbook(golden, actual, spec);
    expect(diff.ok).toBe(true);

    const s = await sweepOf(golden, actual, spec);

    // One per data row. Excluding a column is a decision, not a blind spot, so
    // it must not inflate the headline the way a genuine gap does.
    expect(s.totalExcluded).toBe(5);
    expect(s.totalGaps).toBe(0);
    expect(at(s, 'Policies', 'F2')?.status).toBe('excluded');
    expect(summarizeSweep(s)).toContain('accounted for');
  });

  test('a column layer 1 reported as added is not counted as a gap either', async () => {
    const golden = await buildSweepWorkbook('sw-addcol-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-addcol-actual.xlsx', { insertColumn: true });

    const { diff } = await runWorkbook(golden, actual, SPEC);
    const policies = diff.sheets.find((x) => x.sheet === 'Policies')!;
    expect(policies.diff!.schema.added).toContain('Premium');

    const s = await sweepOf(golden, actual);

    // The new column's own cells were never compared -- but layer 1 named the
    // column, so reporting them again as unchecked would be double-counting.
    expect(s.totalReported).toBeGreaterThan(0);
    expect(at(s, 'Policies', 'C2')?.status).toBe('reported');
  });

  test('a formula change is seen without evaluating anything', async () => {
    const golden = await buildSweepWorkbook('sw-f-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-f-actual.xlsx', { formulaDrift: true });

    const s = await sweepOf(golden, actual);

    // No cell on either side carries a cached value, so a formula compares as
    // its text. Five rows of Annual Cost changed.
    expect(s.totalDifferences).toBe(5);
    const cell = at(s, 'Policies', 'E2')!;
    expect(cell.base).toBe('=C2*D2');
    expect(cell.next).toBe('=C2*D2*1.05');
    expect(cell.status).toBe('compared');
  });

  test('a baseline nobody opened in Excel is not 5 value changes', async () => {
    // The real shape of this: the golden file comes straight from the
    // generator and carries no cached results, while the report under test was
    // opened once and saved, so Excel wrote them. The formulas are identical.
    const golden = await buildSweepWorkbook('sw-recalc-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-recalc-actual.xlsx', { cacheResults: true });

    const s = await sweepOf(golden, actual);

    // A result only one side has says nothing about the other. Reporting these
    // would mean every formula in the file shows up on every run.
    expect(s.totalDifferences).toBe(0);
  });

  test('but a cached result both sides have and disagree on is a difference', async () => {
    const golden = await buildSweepWorkbook('sw-stale-golden.xlsx', { cacheResults: true });
    const actual = await buildSweepWorkbook('sw-stale-actual.xlsx', {
      cacheResults: true,
      sumDrift: { 'P-1002': 90000 },
    });

    const s = await sweepOf(golden, actual);

    // The input changed, so both the input cell and the result it feeds moved.
    expect(s.totalDifferences).toBe(2);
    expect(at(s, 'Policies', 'C3')).toBeTruthy();
    expect(at(s, 'Policies', 'E3')?.base).toContain('→');
  });

  test('flags a reshaped sheet so positional noise arrives explained', async () => {
    const golden = await buildSweepWorkbook('sw-shape-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-shape-actual.xlsx', { insertColumn: true });

    const s = await sweepOf(golden, actual);

    const policies = s.sheets.find((x) => x.sheet === 'Policies')!;
    expect(policies.reshaped).toBe(true);
    expect(policies.differing).toBeGreaterThan(1);
    // Notes was untouched, so it must not be dragged into the noise.
    expect(s.sheets.find((x) => x.sheet === 'Notes')!.differing).toBe(0);
  });

  test('a sheet the baseline lacks is listed, not counted as differences', async () => {
    const golden = await buildSweepWorkbook('sw-add-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-add-actual.xlsx', { extraSheet: true });

    const s = await sweepOf(golden, actual);

    const added = s.sheets.find((x) => x.sheet === 'Addendum')!;
    expect(added.status).toBe('added');
    expect(added.cells).toBeGreaterThan(0);
    // Every cell of a new sheet is "different from nothing". Counting them
    // would bury the cells that genuinely went unchecked, and layer 1 already
    // reports the sheet as added.
    expect(added.differing).toBe(0);
    expect(s.totalDifferences).toBe(0);
    expect(formatSweepReport(s)).toContain('NOT SWEPT (1)');
  });

  test('the report leads with the number that matters', async () => {
    const golden = await buildSweepWorkbook('sw-rep-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-rep-actual.xlsx', { note: 'Extended' });

    const text = formatSweepReport(await sweepOf(golden, actual));

    expect(text).toContain('UNCHECKED DIFFERENCES (1)');
    expect(text).toContain('Notes!B3');
    expect(text).toContain('sheet not compared');
  });

  test('says so plainly when the gaps hid nothing', async () => {
    const golden = await buildSweepWorkbook('sw-clean-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-clean-actual.xlsx', {
      sumDrift: { 'P-1001': 121000 },
    });

    const text = formatSweepReport(await sweepOf(golden, actual));

    expect(text).toContain('nothing changed where nobody was looking');
    expect(text).not.toContain('UNCHECKED DIFFERENCES');
  });

  test('names the sheets nobody is checking, even when they match today', async () => {
    const golden = await buildSweepWorkbook('sw-blind-golden.xlsx');
    const actual = await buildSweepWorkbook('sw-blind-actual.xlsx');

    const text = formatSweepReport(await sweepOf(golden, actual));

    // Notes is identical in both, so it produces no difference -- but nothing
    // is watching it, and that is worth saying before it changes.
    expect(text).toContain('NO COVERAGE');
    expect(text).toContain('Notes');
  });
});

test.describe('runCase and layer 2', () => {
  const caseDir = async (name: string) => {
    const dir = join(DIR, 'sweep-cases', name);
    await rm(dir, { recursive: true, force: true });
    return dir;
  };

  test('writes the report even when it finds nothing', async () => {
    const dir = await caseDir('clean');
    const golden = await buildSweepWorkbook('sw-case-g.xlsx');
    await runCase(golden, dir, SPEC as CaseOptions);

    const actual = await buildSweepWorkbook('sw-case-a.xlsx');
    const r = await runCase(actual, dir, SPEC as CaseOptions);

    // Unlike differences.xlsx, absence would be ambiguous here: "nothing
    // changed, and here is what was checked to say so" is the answer, not the
    // lack of one.
    expect(await exists(r.files.report)).toBe(true);
    const text = await readFile(r.files.report, 'utf8');
    expect(text).toContain('# clean');
    expect(text).not.toContain('Differences found');
    expect(r.sweep?.totalGaps).toBe(0);
  });

  test('reports a gap the case otherwise passes over', async () => {
    const dir = await caseDir('gap');
    const golden = await buildSweepWorkbook('sw-case-g2.xlsx');
    await runCase(golden, dir, SPEC as CaseOptions);

    const actual = await buildSweepWorkbook('sw-case-a2.xlsx', { note: 'Extended' });
    const r = await runCase(actual, dir, SPEC as CaseOptions);

    // Layer 1 passes; layer 2 does not change that, and says what it saw.
    expect(r.ok).toBe(true);
    expect(r.sweep?.totalGaps).toBe(1);
    expect(await readFile(r.files.report, 'utf8')).toContain('B3');
  });

  test('sweepCells: false skips it and clears a stale file', async () => {
    const dir = await caseDir('off');
    const golden = await buildSweepWorkbook('sw-case-g3.xlsx');
    await runCase(golden, dir, SPEC as CaseOptions);

    const actual = await buildSweepWorkbook('sw-case-a3.xlsx', { note: 'Extended' });
    const on = await runCase(actual, dir, SPEC as CaseOptions);
    expect(await exists(on.files.report)).toBe(true);

    const off = await runCase(actual, dir, { ...SPEC, sweepCells: false } as CaseOptions);
    expect(off.sweep).toBeNull();
    // The report is still written; it simply carries no layer 2 sections.
    expect(await readFile(off.files.report, 'utf8')).not.toContain('nothing checked them');
  });

  /**
   * A tolerance is written against a column name and this layer works in
   * addresses, so the two could easily disagree about which gaps matter -- and
   * a layer 2 that ignored tolerances would go on reporting float noise in the
   * headline count after the reader had said, in config, that they did not
   * care about it.
   */
  test.describe('tolerance', () => {
    /** A gap of 1e-7 on P-1003's 240000: the last bits of a double, as a
     * recalculation in a different order leaves behind. */
    const DRIFT = { 'P-1003': 240000.0000001 };

    test('turned off, float noise is a difference like any other', async () => {
      const golden = await buildSweepWorkbook('sw-tol-off-g.xlsx');
      const actual = await buildSweepWorkbook('sw-tol-off-a.xlsx', { sumDrift: DRIFT });

      const s = await sweepOf(golden, actual, {
        ...SPEC,
        defaults: { requireCachedValues: false, tolerance: 0 },
      });

      expect(s.totalDifferences).toBe(1);
      expect(s.totalTolerated).toBe(0);
    });

    test('the default absorbs it without being asked', async () => {
      const golden = await buildSweepWorkbook('sw-tol-default-g.xlsx');
      const actual = await buildSweepWorkbook('sw-tol-default-a.xlsx', { sumDrift: DRIFT });

      const s = await sweepOf(golden, actual);

      expect(s.totalDifferences).toBe(0);
      expect(s.totalTolerated).toBe(1);
    });

    test('a column tolerance from layer 1 quiets the same cell here', async () => {
      const golden = await buildSweepWorkbook('sw-tol-col-g.xlsx');
      const actual = await buildSweepWorkbook('sw-tol-col-a.xlsx', { sumDrift: DRIFT });

      const s = await sweepOf(golden, actual, {
        ...SPEC,
        sheets: {
          Policies: { keyColumns: ['PolicyId'], tolerance: { 'Sum Insured': 0.01 } },
        },
      });

      // Not counted, but not hidden either: it is listed with its own gap.
      expect(s.totalDifferences).toBe(0);
      expect(s.totalTolerated).toBe(1);
      expect(s.tolerated[0]?.address).toBe('C4');
    });

    test('the blanket tolerance covers cells layer 1 never reached', async () => {
      const golden = await buildSweepWorkbook('sw-tol-gap-g.xlsx');
      const actual = await buildSweepWorkbook('sw-tol-gap-a.xlsx', { sumDrift: DRIFT });

      // Nothing is keyed, so layer 1 compares nothing and the drift is a gap.
      const bare = { matchUnkeyedRowsByPosition: false, defaults: { requireCachedValues: false } };
      const { compared } = await runWorkbook(golden, actual, bare);
      expect(compared).toHaveLength(0);

      const loose = await sweep(golden, actual, compared, { tolerance: 0.01 });
      expect(loose.totalGaps).toBe(0);
      expect(loose.totalTolerated).toBe(1);

      // And the same run without it still reports the gap, so the tolerance is
      // what changed the answer rather than the sweep losing sight of the cell.
      const strict = await sweep(golden, actual, compared);
      expect(strict.totalGaps).toBe(1);
    });

    test('a formula whose text changed is never tolerated', async () => {
      const golden = await buildSweepWorkbook('sw-tol-formula-g.xlsx');
      const actual = await buildSweepWorkbook('sw-tol-formula-a.xlsx', { formulaDrift: true });

      const s = await sweepOf(golden, actual, {
        ...SPEC,
        defaults: { requireCachedValues: false, tolerance: 1e9 },
        sheets: { Policies: { keyColumns: ['PolicyId'] } },
      });

      // However wide the tolerance, a changed calculation is a change.
      expect(s.totalTolerated).toBe(0);
      expect(s.totalDifferences).toBeGreaterThan(0);
    });
  });
});
