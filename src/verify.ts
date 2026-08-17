import { extname } from 'node:path';
import type { DiffResult, SheetModel, SheetReader, SheetSpec } from './types.js';
import { resolveSpec } from './model.js';
import { compare } from './compare.js';
import { ExcelReader } from './reader-excel.js';
import { CsvReader } from './reader-csv.js';

const READERS: SheetReader[] = [new ExcelReader(), new CsvReader()];

/**
 * Registers an additional reader, or overrides a built-in for an extension.
 * This is the seam that keeps ExcelJS replaceable: a reader implementing
 * `SheetReader` swaps it out without touching any comparison logic. To take
 * part in workbook and case comparisons it must implement `readWorkbook` too.
 *
 * Returns a function that removes it again, so a test or a scoped override
 * does not leak into the rest of the process.
 */
export function registerReader(reader: SheetReader): () => void {
  READERS.unshift(reader);
  return () => {
    const i = READERS.indexOf(reader);
    if (i >= 0) READERS.splice(i, 1);
  };
}

export function readerFor(path: string): SheetReader {
  const ext = extname(path).toLowerCase();
  const r = READERS.find((x) => x.extensions.includes(ext));
  if (!r) {
    const known = [...new Set(READERS.flatMap((x) => x.extensions))].join(', ');
    throw new Error(`sheet-verify: no reader for "${ext}". Known extensions: ${known}`);
  }
  return r;
}

export async function readSheet(path: string, spec: SheetSpec): Promise<SheetModel> {
  const resolved = resolveSpec(spec);
  return readerFor(path).read(path, resolved);
}

/** Reads both files and compares them. The main entry point. */
export async function verifySheet(
  baselinePath: string,
  actualPath: string,
  spec: SheetSpec,
): Promise<DiffResult> {
  const resolved = resolveSpec(spec);
  const [base, next] = await Promise.all([
    readerFor(baselinePath).read(baselinePath, resolved),
    readerFor(actualPath).read(actualPath, resolved),
  ]);
  return compare(base, next, resolved);
}
