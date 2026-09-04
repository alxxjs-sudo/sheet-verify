import { test, expect } from '@playwright/test';
import { join } from 'node:path';
import { writeFile, mkdir } from 'node:fs/promises';
import {
  registerReader, readerFor, readSheet, verifyWorkbook, buildModel, CSV_SHEET,
} from '../src/index.js';
import type { ResolvedSpec, SheetModel, SheetReader, TableRequest } from '../src/index.js';
import { DIR } from './fixtures.js';

/**
 * The reader seam is what keeps ExcelJS replaceable -- it has had no release
 * since October 2023, so being able to swap it without touching comparison
 * logic is a deliberate design constraint rather than a nicety.
 */

/** A trivial reader for a made-up format: one header line, then rows. */
class PipeReader implements SheetReader {
  readonly extensions = ['.pipe'];

  async read(path: string, spec: ResolvedSpec): Promise<SheetModel> {
    const { readFile } = await import('node:fs/promises');
    const lines = (await readFile(path, 'utf8')).trim().split('\n');
    const headerLine = lines[spec.headerRow - 1] ?? '';
    const headers = headerLine.split('|').map((name, i) => ({ name, colNum: i + 1 }));

    const rows = lines.slice(spec.headerRow).map((line, r) => ({
      rowNum: spec.headerRow + r + 1,
      cells: new Map(
        line.split('|').map((value, i) => [
          i + 1,
          { address: `C${i + 1}R${r}`, kind: 'string' as const, value, formula: null },
        ]),
      ),
    }));

    return buildModel({ source: path, sheet: 'pipe', headers, rows, spec });
  }
}

const writePipe = async (name: string, body: string) => {
  await mkdir(DIR, { recursive: true });
  const full = join(DIR, name);
  await writeFile(full, body, 'utf8');
  return full;
};

test.describe('registerReader', () => {
  test('a custom reader handles its own extension, and unregisters cleanly', async () => {
    const path = await writePipe('r1.pipe', 'Id|Name\nA-1|Ivanov\nA-2|Petrov\n');

    expect(() => readerFor(path)).toThrow(/no reader for "\.pipe"/);

    const unregister = registerReader(new PipeReader());
    try {
      const model = await readSheet(path, { keyColumns: ['Id'] });
      expect(model.headers).toEqual(['Id', 'Name']);
      expect(model.rows.size).toBe(2);
      expect(model.rows.get('A-2')!['Name']!.value).toBe('Petrov');
    } finally {
      unregister();
    }

    // The registry is left as it was found.
    expect(() => readerFor(path)).toThrow(/no reader for "\.pipe"/);
  });

  test('an unknown extension names the ones that are known', async () => {
    expect(() => readerFor('report.parquet')).toThrow(/\.xlsx/);
    expect(() => readerFor('report.parquet')).toThrow(/\.csv/);
  });

  test('a reader without readWorkbook is refused for workbook comparison, with a reason', async () => {
    const path = await writePipe('r2.pipe', 'Id|Name\nA-1|Ivanov\n');
    const unregister = registerReader(new PipeReader());
    try {
      await expect(
        verifyWorkbook(path, path, { sheets: { pipe: { keyColumns: ['Id'] } } }),
      ).rejects.toThrow(/must implement readWorkbook/);
    } finally {
      unregister();
    }
  });

  test('a registered reader overrides a built-in for the same extension', async () => {
    class FakeXlsx extends PipeReader {
      override readonly extensions = ['.xlsx'];
    }

    const unregister = registerReader(new FakeXlsx());
    try {
      expect(readerFor('anything.xlsx')).toBeInstanceOf(FakeXlsx);
    } finally {
      unregister();
    }
    // Removing it restores the built-in rather than leaving a hole.
    expect(readerFor('anything.xlsx').extensions).toContain('.xlsm');
  });
});

test.describe('workbook readers', () => {
  test('the built-in Excel reader satisfies the workbook interface', () => {
    const reader = readerFor('x.xlsx') as { readWorkbook?: unknown };
    expect(typeof reader.readWorkbook).toBe('function');
  });

  test('the CSV reader also satisfies the workbook interface, as one pseudo-sheet', () => {
    const reader = readerFor('x.csv') as { readWorkbook?: unknown };
    expect(typeof reader.readWorkbook).toBe('function');
    expect(CSV_SHEET).toBe('CSV');
  });

  test('a table request list is honoured verbatim by a conforming reader', async () => {
    const seen: TableRequest[] = [];
    class Recording extends PipeReader {
      override readonly extensions = ['.pipe2'];
      async readWorkbook(path: string, tablesFor: (s: string) => TableRequest[]) {
        const requests = tablesFor('pipe');
        seen.push(...requests);
        const models = new Map<string, SheetModel>();
        for (const r of requests) models.set(r.key, await this.read(path, r.spec));
        return { sheets: ['pipe'], models };
      }
    }

    const path = await writePipe('r3.pipe2', 'Id|Name\nA-1|Ivanov\n');
    const unregister = registerReader(new Recording());
    try {
      const d = await verifyWorkbook(path, path, { sheets: { pipe: { keyColumns: ['Id'] } } });
      expect(d.ok).toBe(true);
      expect(seen.map((r) => r.table)).toEqual(['pipe', 'pipe']); // once per file
    } finally {
      unregister();
    }
  });
});

/**
 * A zip read as a workbook whose sheets are its members.
 *
 * The same move the CSV reader makes one level up: a workbook here is a list
 * of named tables, and an archive of CSVs is exactly that. Presenting it as
 * one is what lets keys, tolerances, the sweep and the added/removed-sheet
 * rules reach inside an archive without any of them knowing it is one.
 *
 * Real shape, from a premium allocation download: `details.zip` holding
 * `policy.csv` and `net_of_fac.csv`, 8,476 rows each, whose rows arrive in a
 * different order on every run.
 */
test.describe('an archive as a workbook', () => {
  const zipOf = async (name: string, files: Record<string, string>): Promise<string> => {
    const { default: JSZip } = await import('jszip');
    const zip = new JSZip();
    for (const [inner, body] of Object.entries(files)) zip.file(inner, body);
    await mkdir(DIR, { recursive: true });
    const path = join(DIR, name);
    await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }));
    return path;
  };

  const ROWS = ['Ref,Region,Amount', 'R-1,North,10', 'R-2,South,20', 'R-3,East,30'];
  // Configuring a sheet neither file holds is an error in its own right, so
  // an archive of one member gets a spec naming one member.
  const ONE = { defaults: { requireCachedValues: false }, sheets: { 'policy.csv': { keyColumns: ['Ref'] } } };
  const SPEC = {
    defaults: { requireCachedValues: false },
    sheets: {
      'policy.csv': { keyColumns: ['Ref'] },
      'net_of_fac.csv': { keyColumns: ['Ref'] },
    },
  };

  test('its members are the sheets, and they pair by name', async () => {
    const golden = await zipOf('zip-a-g.zip', {
      'policy.csv': ROWS.join('\n'),
      'net_of_fac.csv': ROWS.join('\n'),
    });
    const actual = await zipOf('zip-a-a.zip', {
      'policy.csv': ROWS.join('\n'),
      'net_of_fac.csv': ROWS.join('\n'),
    });

    const diff = await verifyWorkbook(golden, actual, SPEC);
    expect(diff.base.sheets).toEqual(['policy.csv', 'net_of_fac.csv']);
    expect(diff.ok).toBe(true);
    expect(diff.sheets.filter((s) => s.status === 'compared')).toHaveLength(2);
  });

  test('rows in a different order inside a member are not a difference', async () => {
    // The whole reason to reach inside: these files are byte-different and
    // hold the same rows, and only a keyed comparison can tell you that.
    const golden = await zipOf('zip-order-g.zip', { 'policy.csv': ROWS.join('\n') });
    const actual = await zipOf('zip-order-a.zip', {
      'policy.csv': [ROWS[0]!, ROWS[3]!, ROWS[1]!, ROWS[2]!].join('\n'),
    });

    const diff = await verifyWorkbook(golden, actual, ONE);
    expect(diff.ok).toBe(true);
    expect(diff.sheets[0]!.diff!.rows.compared).toBe(3);
    expect(diff.sheets[0]!.diff!.values).toHaveLength(0);
  });

  test('a changed figure inside a member is still found', async () => {
    const golden = await zipOf('zip-diff-g.zip', { 'policy.csv': ROWS.join('\n') });
    const actual = await zipOf('zip-diff-a.zip', {
      'policy.csv': ['Ref,Region,Amount', 'R-1,North,10', 'R-2,South,999', 'R-3,East,30'].join('\n'),
    });

    const diff = await verifyWorkbook(golden, actual, SPEC);
    expect(diff.ok).toBe(false);
    expect(diff.sheets[0]!.diff!.values).toHaveLength(1);
  });

  test('a member the golden had and the report does not is a removed sheet', async () => {
    // Inherited, not invented: a zip member is a sheet, and a removed sheet is
    // already a defect while an added one is review. So "unused.zip exists in
    // one case and not another" needed no policy of its own.
    const golden = await zipOf('zip-gone-g.zip', {
      'policy.csv': ROWS.join('\n'),
      'net_of_fac.csv': ROWS.join('\n'),
    });
    const actual = await zipOf('zip-gone-a.zip', { 'policy.csv': ROWS.join('\n') });

    const diff = await verifyWorkbook(golden, actual, SPEC);
    expect(diff.ok).toBe(false);
    expect(diff.sheetSchema.removed).toEqual(['net_of_fac.csv']);
  });

  test('an archive of one table can be read as that table', async () => {
    // `unused.zip` in a real tree holds exactly one unused_policy.csv.
    const one = await zipOf('zip-single.zip', { 'unused_policy.csv': ROWS.join('\n') });
    const model = await readSheet(one, { sheet: 1, keyColumns: ['Ref'] });
    expect(model.rows.size).toBe(3);
  });

  test('an archive of several says so rather than picking one', async () => {
    const many = await zipOf('zip-many.zip', {
      'policy.csv': ROWS.join('\n'),
      'net_of_fac.csv': ROWS.join('\n'),
    });
    await expect(readSheet(many, { sheet: 1, keyColumns: ['Ref'] }))
      .rejects.toThrow(/holds 2 tables/);
  });

  test('a member no reader handles is passed over, not fatal', async () => {
    // An archive is free to carry a manifest or a licence. Neither is a table.
    const golden = await zipOf('zip-extra-g.zip', {
      'policy.csv': ROWS.join('\n'),
      'README.txt.bak': 'not a table',
    });
    const actual = await zipOf('zip-extra-a.zip', {
      'policy.csv': ROWS.join('\n'),
      'README.txt.bak': 'still not a table',
    });

    const diff = await verifyWorkbook(golden, actual, ONE);
    expect(diff.base.sheets).toEqual(['policy.csv']);
    expect(diff.ok).toBe(true);
  });

  test('the reader is found by extension like any other', async () => {
    expect(readerFor('anything.zip').extensions).toContain('.zip');
  });
});
