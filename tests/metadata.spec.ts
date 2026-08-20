import { test, expect } from '@playwright/test';
import { mkdir, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { runCase, runWorkbook, sweep, parseMetadata, ledgerRows } from '../src/index.js';
import type { SweepResult, WorkbookSpec } from '../src/index.js';
import { DIR } from './fixtures.js';

/**
 * Report metadata: the run's identity rather than its content.
 *
 * The behaviour being pinned here is not "these cells are ignored" -- an
 * ignored cell is invisible, and invisible exclusions are how a comparison
 * quietly stops meaning anything. It is "these cells are read, listed, and
 * left out of the verdict", which is a different and much narrower claim.
 */

interface Build {
  /** Value of the report name in the fused title cell. */
  name?: string;
  /** Value beside the "Report ID" label. */
  id?: number;
  /** Value beside "Report Creator". */
  creator?: string;
  /** The bare generated-on date, which carries no label of its own. */
  generated?: string;
  /** A real figure, so a test can prove metadata handling did not eat it. */
  premium?: number;
  /** Header of the data table's second column, to test a look-alike label. */
  secondColumn?: string;
  /** Merge A12:A13, so A13 is a slave carrying A12's value. */
  merge?: boolean;
  /** Value of the merged banner at A12. */
  banner?: string;
}

async function build(path: string, o: Build = {}): Promise<string> {
  await mkdir(DIR, { recursive: true });
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Summary');

  // A1 fuses label and value the way a generated title cell does; A2 is a bare
  // date; A4:B6 is a label/value block.
  ws.getCell('A1').value = { formula: `"Report name: " & "${o.name ?? 'RUN-A'}"`, result: `Report name: ${o.name ?? 'RUN-A'}` };
  ws.getCell('A2').value = o.generated ?? '2026-08-17';
  ws.getCell('A4').value = 'Report ID';
  ws.getCell('B4').value = o.id ?? 4542;
  ws.getCell('A5').value = 'Report Creator';
  ws.getCell('B5').value = o.creator ?? 'A Person';
  ws.getCell('A6').value = 'Currency';
  ws.getCell('B6').value = 'USD';

  ws.getCell('A8').value = 'PolicyId';
  ws.getCell('B8').value = o.secondColumn ?? 'Premium';
  ws.getCell('A9').value = 'P-1';
  ws.getCell('B9').value = o.premium ?? 1000;
  ws.getCell('A10').value = 'P-2';
  ws.getCell('B10').value = 2000;

  if (o.merge) {
    ws.getCell('A12').value = o.banner ?? 'run 1786955263151';
    ws.mergeCells('A12:A13');
  }

  const full = join(DIR, path);
  await wb.xlsx.writeFile(full);
  return full;
}

const SPEC: WorkbookSpec = {
  defaults: { requireCachedValues: false },
  metadata: ['Report name', 'Report ID', 'Report Creator', 'Summary!A2'],
  sheets: { Summary: { headerRow: 8, keyColumns: ['PolicyId'] } },
};

async function run(golden: string, actual: string, spec = SPEC): Promise<SweepResult> {
  const { compared } = await runWorkbook(golden, actual, spec);
  return sweep(golden, actual, compared, { metadata: spec.metadata });
}

const found = (s: SweepResult, address: string) =>
  s.metadata.find((m) => m.address === address);

test.describe('report metadata', () => {
  test('a fused title, a labelled value and a bare date are all recognised', async () => {
    const golden = await build('meta-a.xlsx');
    const actual = await build('meta-b.xlsx');
    const s = await run(golden, actual);

    // A1 fuses label and value, so only it. A4/B4 and A5/B5 are pairs. A2 is
    // named by address because it has no label to be found by.
    expect(s.metadata.map((m) => m.address).sort()).toEqual(
      ['A1', 'A2', 'A4', 'A5', 'B4', 'B5'],
    );
    expect(found(s, 'B4')?.rule).toBe('report id');
    expect(found(s, 'A2')?.rule).toBe('A2');
  });

  test('metadata that differs is listed, and does not fail the run', async () => {
    const golden = await build('meta-same.xlsx');
    const actual = await build('meta-drift.xlsx', {
      name: 'RUN-B', id: 9999, creator: 'Someone Else', generated: '2026-08-18',
    });

    const { diff, compared } = await runWorkbook(golden, actual, SPEC);
    const s = await sweep(golden, actual, compared, { metadata: SPEC.metadata });

    expect(diff.ok).toBe(true);
    expect(s.totalDifferences).toBe(0);
    expect(s.metadataDiffering).toBe(4);

    // Listed with both sides, so an unexpected one is still there to be seen.
    const id = found(s, 'B4')!;
    expect(id.base).toBe('4542');
    expect(id.next).toBe('9999');
  });

  test('a figure beside a metadata label is still compared', async () => {
    const golden = await build('meta-fig-a.xlsx');
    const actual = await build('meta-fig-b.xlsx', { id: 9999, premium: 1500 });
    const { diff } = await runWorkbook(golden, actual, SPEC);

    expect(diff.ok).toBe(false);
    const sheet = diff.sheets.find((x) => x.sheet === 'Summary')!;
    expect(sheet.diff!.values).toHaveLength(1);
    expect(sheet.diff!.values[0]!.address).toBe('B9');
  });

  test('a label in a key-value table is dropped by layer 1 too, not just layer 2', async () => {
    // Without the fold into ignoreRows the same cell gets two verdicts: layer 2
    // calls it metadata while layer 1 reports it as a value change.
    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      metadata: ['Report ID', 'Report Creator'],
      sheets: { Summary: { headerRow: 4, keyColumns: ['Report ID'] } },
    };
    const golden = await build('meta-kv-a.xlsx');
    const actual = await build('meta-kv-b.xlsx', { id: 9999, creator: 'Someone Else' });

    const { diff } = await runWorkbook(golden, actual, spec);
    const sheet = diff.sheets.find((x) => x.sheet === 'Summary')!;
    expect(sheet.diff!.values).toHaveLength(0);
  });

  test('a sheet qualifier confines a label that means something else elsewhere', async () => {
    // "Premium" is a column heading here. Qualified to another sheet, the
    // pattern must not touch it -- and the column must still be compared.
    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      metadata: ['Cover!Premium'],
      sheets: { Summary: { headerRow: 8, keyColumns: ['PolicyId'] } },
    };
    const golden = await build('meta-q-a.xlsx');
    const actual = await build('meta-q-b.xlsx', { premium: 1500 });

    const { diff, compared } = await runWorkbook(golden, actual, spec);
    const s = await sweep(golden, actual, compared, { metadata: spec.metadata });

    expect(s.metadata).toHaveLength(0);
    expect(diff.sheets.find((x) => x.sheet === 'Summary')!.diff!.values).toHaveLength(1);
  });

  test('a label only counts when nothing but a separator follows it', async () => {
    const rules = parseMetadata(['Report ID']);
    expect(rules.labels).toEqual([{ sheet: '', label: 'report id' }]);

    // "Report IDs by region" begins with the label and is not it.
    const golden = await build('meta-near-a.xlsx', { secondColumn: 'Report IDs by region' });
    const actual = await build('meta-near-b.xlsx', {
      secondColumn: 'Report IDs by region', premium: 1500,
    });
    const s = await run(golden, actual);

    expect(s.metadata.some((m) => m.address === 'B8')).toBe(false);
  });

  test('a * in a label covers however a report type spells its own name', async () => {
    // "Summary Report Name", "Regional Report Name", "Quarterly Report Name":
    // one pattern, or a config that misses whichever type is added next.
    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      metadata: ['*Report Name'],
      sheets: { Summary: { headerRow: 4, keyColumns: ['Report ID'] } },
    };
    const golden = await build('meta-glob-a.xlsx', { name: 'RUN-A' });
    const actual = await build('meta-glob-b.xlsx', { name: 'RUN-B' });

    const { compared } = await runWorkbook(golden, actual, spec);
    const s = await sweep(golden, actual, compared, { metadata: spec.metadata });

    // A1 is `="Report name: " & <name>`, fused, so only that cell.
    expect(s.metadata.map((m) => m.address)).toEqual(['A1']);
    expect(s.metadataDiffering).toBe(1);
    expect(s.totalDifferences).toBe(0);
  });

  test('an address rule reaches layer 1 too, so no cell gets two verdicts', async () => {
    // The bug this pins: the fold into ignoreRows covered label rules only, so
    // a cell named by address was set aside by layer 2 and reported as a value
    // change by layer 1 -- the same cell, in one report, twice.
    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      metadata: ['Summary!B9'],
      sheets: { Summary: { headerRow: 8, keyColumns: ['PolicyId'] } },
    };
    const golden = await build('meta-addr-a.xlsx');
    const actual = await build('meta-addr-b.xlsx', { premium: 1500 });

    const { diff, compared } = await runWorkbook(golden, actual, spec);
    const s = await sweep(golden, actual, compared, { metadata: spec.metadata });

    expect(diff.sheets.find((x) => x.sheet === 'Summary')!.diff!.values).toHaveLength(0);
    expect(s.metadata.map((m) => m.address)).toEqual(['B9']);
    expect(s.metadataDiffering).toBe(1);
    expect(diff.ok).toBe(true);
  });

  test('a merged banner counts once, so naming the master is enough', async () => {
    // ExcelJS hands every slave the master's value. Read literally, one changed
    // banner arrives as a change per cell in the range, and an exclusion naming
    // the master leaves the slaves still reporting it.
    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      metadata: ['Summary!A12'],
      sheets: { Summary: { headerRow: 8, endRow: 13, keyColumns: ['PolicyId'] } },
    };
    const golden = await build('meta-merge-a.xlsx', { merge: true, banner: 'run 111' });
    const actual = await build('meta-merge-b.xlsx', { merge: true, banner: 'run 222' });

    const { diff, compared } = await runWorkbook(golden, actual, spec);
    const s = await sweep(golden, actual, compared, { metadata: spec.metadata });

    // A13 is a slave: blank, not a second copy of the banner.
    expect(diff.sheets.find((x) => x.sheet === 'Summary')!.diff!.values).toHaveLength(0);
    expect(s.metadata.map((m) => m.address)).toEqual(['A12']);
  });

  test('the cell ledger does not call a metadata cell a difference', async () => {
    // differences.xlsx is read as a list of defects. A run-stamped report name
    // sitting in it under "value-differs" is the false alarm the metadata list
    // exists to remove -- it belongs in the full ledger, not that one.
    const spec: WorkbookSpec = {
      defaults: { requireCachedValues: false },
      metadata: ['Summary!B9'],
      sheets: { Summary: { headerRow: 8, keyColumns: ['PolicyId'] } },
    };
    const golden = await build('meta-led-a.xlsx');
    const actual = await build('meta-led-b.xlsx', { premium: 1500 });
    const { compared } = await runWorkbook(golden, actual, spec);

    const differences = [...ledgerRows(compared, 'differences')];
    expect(differences.filter((r) => r.actualAddress === 'B9')).toHaveLength(0);

    // Still in the complete ledger, marked for what it is.
    const all = [...ledgerRows(compared, 'all')].filter((r) => r.actualAddress === 'B9');
    expect(all).toHaveLength(1);
    expect(all[0]!.status).toBe('metadata');
  });

  test('the report says what it skipped, and why', async () => {
    const dir = join(DIR, 'cases', 'metadata-report');
    const golden = await build('meta-rep-a.xlsx');
    const actual = await build('meta-rep-b.xlsx', { id: 9999 });

    await runCase(golden, dir, SPEC);
    const r = await runCase(actual, dir, SPEC);
    const text = await readFile(r.files.report, 'utf8');

    expect(text).toContain('Two-layer verification');
    expect(text).toContain('## Not verified, on purpose');
    expect(text).toContain('### Report metadata');
    // The skipped cell is named with both its values, not merely counted.
    expect(text).toContain('4542');
    expect(text).toContain('9999');
    // And the old fraction, which read as "223 cells nobody looked at", is gone.
    expect(text).not.toContain('cells checked');
    expect(text).not.toContain('differing, unchecked');
  });
});
