import { test, expect } from '@playwright/test';
import {
  verifyWorkbook, mergeSheetSpec, resolveTables, formatWorkbookReport, summarizeWorkbook,
} from '../src/index.js';
import { ExcelReader } from '../src/reader-excel.js';
import { resolveSpec } from '../src/model.js';
import type { SheetOutcome, WorkbookSpec } from '../src/types.js';
import { buildMultiSheet, buildTwoTableSheet } from './fixtures.js';

/** The spec a real caller writes: sheets share nothing but the defaults. */
const SPEC: WorkbookSpec = {
  defaults: { tolerance: { '*': 0 } },
  sheets: {
    Policies: { keyColumns: ['PolicyId'] },
    Premiums: { keyColumns: ['PolicyId', 'Period'] },
    Regions: { keyColumns: ['Region'] },
  },
};

const statusOf = (sheets: SheetOutcome[], name: string) =>
  sheets.find((s) => s.sheet === name)?.status;

test.describe('verifyWorkbook', () => {
  test('identical workbooks match, every sheet compared', async () => {
    const base = await buildMultiSheet('wb-base.xlsx');
    const next = await buildMultiSheet('wb-same.xlsx');

    const d = await verifyWorkbook(base, next, SPEC);

    expect(d.ok).toBe(true);
    expect(d.reviewOnly).toBe(false);
    expect(d.sheets.map((s) => s.status)).toEqual(['compared', 'compared', 'compared']);
    expect(summarizeWorkbook(d)).toBe('identical');
  });

  test('a column inserted on one sheet stays a schema change and does not touch the others', async () => {
    const base = await buildMultiSheet('wb-ins-base.xlsx');
    const next = await buildMultiSheet('wb-ins-next.xlsx', { insertPremium: true });

    const d = await verifyWorkbook(base, next, SPEC);

    // No defects anywhere: the shifted formulas resolve to the same header refs.
    expect(d.ok).toBe(true);
    expect(d.reviewOnly).toBe(true);

    const policies = d.sheets.find((s) => s.sheet === 'Policies')!.diff!;
    expect(policies.schema.added).toEqual(['Premium']);
    expect(policies.values).toHaveLength(0);
    expect(policies.formulas).toHaveLength(0);

    // A schema-only release must not read as "no differences".
    expect(summarizeWorkbook(d)).toBe('1 table to review');

    for (const name of ['Premiums', 'Regions']) {
      const other = d.sheets.find((s) => s.sheet === name)!.diff!;
      expect(other.ok).toBe(true);
      expect(other.schema.added).toHaveLength(0);
    }
  });

  test('a defect is attributed to the sheet it is on, leaving the rest passing', async () => {
    const base = await buildMultiSheet('wb-def-base.xlsx');
    const next = await buildMultiSheet('wb-def-next.xlsx', {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });

    const d = await verifyWorkbook(base, next, SPEC);

    expect(d.ok).toBe(false);
    const premiums = d.sheets.find((s) => s.sheet === 'Premiums')!.diff!;
    expect(premiums.ok).toBe(false);
    expect(premiums.values.filter((v) => v.rootCause)).toHaveLength(1);
    expect(premiums.values[0]!.key).toContain('P-1003');

    expect(d.sheets.find((s) => s.sheet === 'Policies')!.diff!.ok).toBe(true);
    expect(d.sheets.find((s) => s.sheet === 'Regions')!.diff!.ok).toBe(true);
    expect(summarizeWorkbook(d)).toContain('1 sheet failing');
  });

  test('a new sheet is noted rather than compared, and does not fail the run', async () => {
    const base = await buildMultiSheet('wb-add-base.xlsx');
    const next = await buildMultiSheet('wb-add-next.xlsx', {
      sheets: ['Policies', 'Premiums', 'Regions', 'Premium Detail'],
    });

    const d = await verifyWorkbook(base, next, SPEC);

    expect(d.ok).toBe(true);
    expect(d.reviewOnly).toBe(true);
    expect(summarizeWorkbook(d)).toContain('1 sheet added');
    expect(d.sheetSchema.added).toEqual(['Premium Detail']);
    expect(statusOf(d.sheets, 'Premium Detail')).toBe('added');
    expect(d.sheets.filter((s) => s.status === 'compared')).toHaveLength(3);
    expect(formatWorkbookReport(d)).toContain('new, not compared');
  });

  test('a removed sheet is a defect, not a review item', async () => {
    const base = await buildMultiSheet('wb-rm-base.xlsx');
    const next = await buildMultiSheet('wb-rm-next.xlsx', { sheets: ['Policies', 'Premiums'] });

    const d = await verifyWorkbook(base, next, SPEC);

    expect(d.ok).toBe(false);
    expect(d.sheetSchema.removed).toEqual(['Regions']);
    expect(statusOf(d.sheets, 'Regions')).toBe('removed');
    expect(formatWorkbookReport(d)).toContain('no longer produced');
  });

  test('a sheet with no key configured is reported as a coverage gap, not silently skipped', async () => {
    const base = await buildMultiSheet('wb-skip-base.xlsx');
    const next = await buildMultiSheet('wb-skip-next.xlsx');

    const d = await verifyWorkbook(base, next, {
      sheets: { Policies: { keyColumns: ['PolicyId'] }, Premiums: { keyColumns: ['PolicyId', 'Period'] } },
    });

    expect(d.ok).toBe(true);
    expect(d.reviewOnly).toBe(true);
    expect(statusOf(d.sheets, 'Regions')).toBe('skipped');
    expect(d.sheets.find((s) => s.sheet === 'Regions')!.reason).toContain('keyColumns');
    expect(formatWorkbookReport(d)).toContain('Regions');
  });

  test('strictSheets turns added and unconfigured sheets into failures', async () => {
    const base = await buildMultiSheet('wb-strict-base.xlsx');
    const next = await buildMultiSheet('wb-strict-next.xlsx', {
      sheets: ['Policies', 'Premiums', 'Regions', 'Scratch'],
    });

    const lenient = await verifyWorkbook(base, next, SPEC);
    expect(lenient.ok).toBe(true);

    const strict = await verifyWorkbook(base, next, { ...SPEC, strictSheets: true });
    expect(strict.ok).toBe(false);
  });

  test('ignoreSheets excludes a tab from both comparison and the added list', async () => {
    const base = await buildMultiSheet('wb-ign-base.xlsx');
    const next = await buildMultiSheet('wb-ign-next.xlsx', {
      sheets: ['Policies', 'Premiums', 'Regions', 'Scratch'],
    });

    const d = await verifyWorkbook(base, next, { ...SPEC, ignoreSheets: ['Scratch'] });

    expect(d.ok).toBe(true);
    expect(statusOf(d.sheets, 'Scratch')).toBe('ignored');
    expect(d.sheetSchema.added).toHaveLength(0);
    expect(d.reviewOnly).toBe(false);
  });

  test('an ignored sheet that disappears does not fail the run', async () => {
    const base = await buildMultiSheet('wb-ignrm-base.xlsx', {
      sheets: ['Policies', 'Premiums', 'Regions', 'Scratch'],
    });
    const next = await buildMultiSheet('wb-ignrm-next.xlsx');

    const d = await verifyWorkbook(base, next, { ...SPEC, ignoreSheets: ['Scratch'] });

    expect(d.sheetSchema.removed).toHaveLength(0);
    expect(d.ok).toBe(true);
    expect(d.reviewOnly).toBe(false);
  });

  test('reordered sheets are reported as moved, not as added and removed', async () => {
    const base = await buildMultiSheet('wb-ord-base.xlsx');
    const next = await buildMultiSheet('wb-ord-next.xlsx', {
      sheets: ['Premiums', 'Policies', 'Regions'],
    });

    const d = await verifyWorkbook(base, next, SPEC);

    expect(d.sheetSchema.added).toHaveLength(0);
    expect(d.sheetSchema.removed).toHaveLength(0);
    expect(d.sheetSchema.moved).toEqual([
      { sheet: 'Policies', from: 1, to: 2 },
      { sheet: 'Premiums', from: 2, to: 1 },
    ]);
    expect(d.ok).toBe(true);
  });

  test('sheet names match case-insensitively, the way Excel treats them', async () => {
    const base = await buildMultiSheet('wb-case-base.xlsx');
    const next = await buildMultiSheet('wb-case-next.xlsx');

    const d = await verifyWorkbook(base, next, {
      sheets: {
        policies: { keyColumns: ['PolicyId'] },
        PREMIUMS: { keyColumns: ['PolicyId', 'Period'] },
        Regions: { keyColumns: ['Region'] },
      },
    });

    expect(d.sheets.every((s) => s.status === 'compared')).toBe(true);
    expect(d.errors).toHaveLength(0);
  });

  test('a spec entry matching no sheet is an error, since it is almost always a typo', async () => {
    const base = await buildMultiSheet('wb-typo-base.xlsx');
    const next = await buildMultiSheet('wb-typo-next.xlsx');

    const d = await verifyWorkbook(base, next, {
      ...SPEC,
      sheets: { ...SPEC.sheets, Policyz: { keyColumns: ['PolicyId'] } },
    });

    expect(d.ok).toBe(false);
    expect(d.errors.join()).toContain('Policyz');
  });

  test('CSV is rejected with a pointer to the single-sheet API', async () => {
    const base = await buildMultiSheet('wb-csv-base.xlsx');
    await expect(verifyWorkbook('nope.csv', base, SPEC)).rejects.toThrow(/verifySheet/);
  });
});

test.describe('several tables on one sheet', () => {
  /** Info block on rows 1-5, data table header on row 7. */
  const TWO_TABLE: WorkbookSpec = {
    sheets: {
      Policies: {
        tables: {
          Info: { headerRow: 1, keyColumns: ['Field'] },
          Detail: { headerRow: 7, keyColumns: ['PolicyId'] },
        },
      },
    },
  };

  test('each table is bounded by the next, so the info block excludes the data below it', async () => {
    const base = await buildTwoTableSheet('tt-base.xlsx');
    const next = await buildTwoTableSheet('tt-same.xlsx');

    const d = await verifyWorkbook(base, next, TWO_TABLE);

    expect(d.ok).toBe(true);
    expect(d.sheets.map((s) => s.label)).toEqual(['Policies · Info', 'Policies · Detail']);

    // The info block has 3 property rows -- not the 3 + blank + 5 policy rows
    // it would swallow if it ran to the bottom of the sheet.
    const info = d.sheets.find((s) => s.table === 'Info')!.diff!;
    expect(info.base.rows).toBe(3);
    expect(info.schema.compared).toEqual(['Field', 'Value']);

    const detail = d.sheets.find((s) => s.table === 'Detail')!.diff!;
    expect(detail.base.rows).toBe(5);
    expect(detail.schema.compared).toContain('Annual Cost');
  });

  test('a changed info field is reported against the info table alone', async () => {
    const base = await buildTwoTableSheet('tt-rel-base.xlsx');
    const next = await buildTwoTableSheet('tt-rel-next.xlsx', { release: '4.3.0' });

    const d = await verifyWorkbook(base, next, TWO_TABLE);

    expect(d.ok).toBe(false);
    const info = d.sheets.find((s) => s.table === 'Info')!.diff!;
    expect(info.values).toHaveLength(1);
    expect(info.values[0]!.key).toBe('Release');
    expect(d.sheets.find((s) => s.table === 'Detail')!.diff!.ok).toBe(true);
  });

  test('a per-run timestamp can be ignored on the info table without touching the data table', async () => {
    const base = await buildTwoTableSheet('tt-ts-base.xlsx');
    const next = await buildTwoTableSheet('tt-ts-next.xlsx', {
      generatedAt: '2026-08-16T09:31:44Z',
    });

    const noisy = await verifyWorkbook(base, next, TWO_TABLE);
    expect(noisy.ok).toBe(false);

    // The info block is keyed by field name, so the row is dropped by key.
    const quiet = await verifyWorkbook(base, next, {
      sheets: {
        Policies: {
          tables: {
            Info: { headerRow: 1, keyColumns: ['Field'], ignoreRows: ['Generated At'] },
            Detail: { headerRow: 7, keyColumns: ['PolicyId'] },
          },
        },
      },
    });
    expect(quiet.ok).toBe(true);
  });

  test('a defect in the data table is still found with the info block present', async () => {
    const base = await buildTwoTableSheet('tt-def-base.xlsx');
    const next = await buildTwoTableSheet('tt-def-next.xlsx', { drift: 999000 });

    const d = await verifyWorkbook(base, next, TWO_TABLE);

    expect(d.ok).toBe(false);
    const detail = d.sheets.find((s) => s.table === 'Detail')!.diff!;
    expect(detail.values.filter((v) => v.rootCause)).toHaveLength(1);
    expect(detail.values[0]!.column).toBe('Sum Insured');
    expect(d.sheets.find((s) => s.table === 'Info')!.diff!.ok).toBe(true);
  });

  test('a row added to the info block shifts the data table, and is absorbed', async () => {
    const base = await buildTwoTableSheet('tt-shift-base.xlsx');
    const next = await buildTwoTableSheet('tt-shift-next.xlsx', { extraInfo: true });

    const d = await verifyWorkbook(base, next, TWO_TABLE);

    // The info block gained a property; the data table is unchanged and
    // still sits on row 7, so it compares clean.
    const info = d.sheets.find((s) => s.table === 'Info')!.diff!;
    expect(info.rows.added).toEqual(['Source System']);
    expect(d.sheets.find((s) => s.table === 'Detail')!.diff!.ok).toBe(true);
  });

  test('tables are ordered by header row regardless of declaration order', () => {
    const [first, second] = resolveTables(
      {
        sheets: {
          Policies: {
            tables: {
              Detail: { headerRow: 7, keyColumns: ['PolicyId'] },
              Info: { headerRow: 1, keyColumns: ['Field'] },
            },
          },
        },
      },
      'Policies',
    );

    expect(first!.table).toBe('Info');
    expect(first!.spec!.endRow).toBe(6);
    expect(second!.table).toBe('Detail');
    expect(second!.spec!.endRow).toBe(0);
  });

  test('a table with no key is a coverage gap naming the table, not the sheet', async () => {
    const base = await buildTwoTableSheet('tt-nokey-base.xlsx');
    const next = await buildTwoTableSheet('tt-nokey-next.xlsx');

    const d = await verifyWorkbook(base, next, {
      sheets: {
        Policies: {
          tables: {
            Info: { headerRow: 1 },
            Detail: { headerRow: 7, keyColumns: ['PolicyId'] },
          },
        },
      },
    });

    const info = d.sheets.find((s) => s.table === 'Info')!;
    expect(info.status).toBe('skipped');
    expect(info.reason).toContain('this table');
    expect(formatWorkbookReport(d)).toContain('Policies · Info');
  });
});

test.describe('spec merging', () => {
  test('tolerance records merge instead of the per-sheet one replacing the default', () => {
    const merged = mergeSheetSpec(
      { tolerance: { '*': 0.01, Rate: 0 } },
      { tolerance: { Amount: 0.5 } },
    );
    expect(merged.tolerance).toEqual({ '*': 0.01, Rate: 0, Amount: 0.5 });
  });

  test('ignoreColumns and invariants accumulate rather than override', () => {
    const merged = mergeSheetSpec(
      { ignoreColumns: ['Generated At'], invariants: [{ name: 'a', check: () => [] }] },
      { ignoreColumns: ['Run Id'], invariants: [{ name: 'b', check: () => [] }] },
    );
    expect(merged.ignoreColumns).toEqual(['Generated At', 'Run Id']);
    expect(merged.invariants!.map((i) => i.name)).toEqual(['a', 'b']);
  });

  test('a plain numeric tolerance becomes the * fallback when merged', () => {
    expect(mergeSheetSpec({ tolerance: 0.25 }, {}).tolerance).toEqual({ '*': 0.25 });
  });

  test('scalar options are replaced by the per-sheet value', () => {
    const merged = mergeSheetSpec({ headerRow: 1, strictSchema: true }, { headerRow: 3 });
    expect(merged.headerRow).toBe(3);
    expect(merged.strictSchema).toBe(true);
  });
});

test.describe('ExcelReader.readWorkbook', () => {
  test('lists every sheet but builds models only for configured ones', async () => {
    const path = await buildMultiSheet('wb-reader.xlsx');
    const spec = resolveSpec({ keyColumns: ['PolicyId'] });

    const seen: string[] = [];
    const { sheets, models } = await new ExcelReader().readWorkbook(path, (name) => {
      seen.push(name);
      return name === 'Policies' ? [{ table: 'Policies', key: 'Policies', spec }] : [];
    });

    // Every sheet is offered exactly once -- the file is parsed a single time,
    // rather than reopened per sheet as a read()-per-sheet loop would do.
    expect(seen).toEqual(['Policies', 'Premiums', 'Regions']);
    expect(sheets).toEqual(['Policies', 'Premiums', 'Regions']);
    expect([...models.keys()]).toEqual(['Policies']);
    expect(models.get('Policies')!.rows.size).toBe(5);
  });
});
