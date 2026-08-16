import { readFile } from 'node:fs/promises';
import { parse } from 'csv-parse/sync';
import type { CellValue, CsvDialect, ResolvedSpec, SheetModel, SheetReader } from './types.js';
import { buildModel, type RawCell } from './model.js';
import { numToCol } from './a1.js';

/** Guards against turning identifiers like "0012" or "1-2" into numbers. */
const STRICT_NUMBER = /^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][-+]?[0-9]+)?$/;

function detectDialect(text: string, forced: string): CsvDialect {
  const bom = text.charCodeAt(0) === 0xfeff;
  const body = bom ? text.slice(1) : text;

  const crlf = (body.match(/\r\n/g) ?? []).length;
  const lf = (body.match(/(?<!\r)\n/g) ?? []).length;
  const cr = (body.match(/\r(?!\n)/g) ?? []).length;
  const kinds = [crlf && 'CRLF', lf && 'LF', cr && 'CR'].filter(Boolean) as string[];
  const lineEnding =
    kinds.length === 0 ? 'none' : kinds.length > 1 ? 'mixed' : (kinds[0] as CsvDialect['lineEnding']);

  let delimiter = forced;
  if (!delimiter) {
    const firstLine = body.split(/\r\n|\n|\r/, 1)[0] ?? '';
    // Count candidates outside quoted regions.
    const unquoted = firstLine.replace(/"(?:[^"]|"")*"/g, '');
    const counts = [',', ';', '\t', '|'].map((d) => [d, unquoted.split(d).length - 1] as const);
    counts.sort((a, b) => b[1] - a[1]);
    delimiter = counts[0]![1] > 0 ? counts[0]![0] : ',';
  }

  return { delimiter, bom, lineEnding };
}

export class CsvReader implements SheetReader {
  readonly extensions = ['.csv', '.tsv', '.txt'];

  async read(path: string, spec: ResolvedSpec): Promise<SheetModel> {
    const buf = await readFile(path);
    const text = buf.toString('utf8');
    const dialect = detectDialect(text, spec.csv.delimiter);

    const records: string[][] = parse(dialect.bom ? text.slice(1) : text, {
      delimiter: dialect.delimiter,
      relax_column_count: true,
      skip_empty_lines: true,
      relax_quotes: true,
    });

    const headerLine = records[spec.headerRow - 1];
    if (!headerLine) {
      throw new Error(`sheet-verify: no header row at line ${spec.headerRow} of ${path}`);
    }

    const headers = headerLine
      .map((name, i) => ({ name: String(name ?? '').trim(), colNum: i + 1 }))
      .filter((h) => h.name !== '');

    const rows = records.slice(spec.headerRow).map((rec, i) => {
      const cells = new Map<number, RawCell>();
      for (const { name, colNum } of headers) {
        const text = rec[colNum - 1];
        cells.set(colNum, {
          address: `${numToCol(colNum)}${spec.headerRow + i + 1}`,
          kind: 'string',
          value: this.coerce(text, name, spec),
          formula: null, // CSV carries results only; formulas never survive export
        });
      }
      return { rowNum: spec.headerRow + i + 1, cells };
    });

    const model = buildModel({ source: path, sheet: '', headers, rows, spec });
    model.dialect = dialect;
    return model;
  }

  private coerce(text: string | undefined, column: string, spec: ResolvedSpec): CellValue {
    if (text === undefined || text === null) return null;
    const t = text.trim();
    if (t === '') return null;

    const mode = spec.csv.numeric;
    if (mode === 'none') return t;
    if (mode === 'tolerance-only' && !(column in spec.tolerance)) return t;
    if (spec.keyColumns.includes(column)) return t; // identifiers stay text

    if (STRICT_NUMBER.test(t)) {
      const n = Number(t);
      if (Number.isFinite(n)) return n;
    }
    return t;
  }
}

/** Reports encoding-level drift that looks like data change but is not. */
export function dialectDrift(a: SheetModel, b: SheetModel): string[] {
  if (!a.dialect || !b.dialect) return [];
  const out: string[] = [];
  const show = (d: string) => (d === '\t' ? '\\t' : d);
  if (a.dialect.delimiter !== b.dialect.delimiter)
    out.push(`delimiter changed: "${show(a.dialect.delimiter)}" -> "${show(b.dialect.delimiter)}"`);
  if (a.dialect.bom !== b.dialect.bom)
    out.push(`BOM ${a.dialect.bom ? 'removed' : 'added'}`);
  if (a.dialect.lineEnding !== b.dialect.lineEnding)
    out.push(`line endings changed: ${a.dialect.lineEnding} -> ${b.dialect.lineEnding}`);
  return out;
}
