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
