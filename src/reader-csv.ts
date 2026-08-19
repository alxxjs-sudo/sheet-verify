import { readFile } from 'node:fs/promises';
import { parse } from 'csv-parse/sync';
import type {
  CellValue, CsvDialect, ResolvedSpec, SheetModel, TableRequest, WorkbookReader,
} from './types.js';
import { buildModel, type RawCell } from './model.js';
import { columnRange, numToCol } from './a1.js';

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

/**
 * The single pseudo-sheet a CSV presents. A constant, not the file name:
 * `golden.csv` and `actual.csv` must land on the same sheet name to be paired.
 */
export const CSV_SHEET = 'CSV';

export class CsvReader implements WorkbookReader {
  readonly extensions = ['.csv', '.tsv', '.txt'];

  /**
   * A CSV is one table, so it presents itself as a one-sheet workbook. That
   * lets cases, ledgers and the CLI treat CSV exactly like a workbook instead
   * of every caller special-casing it.
   */
  async readWorkbook(
    path: string,
    tablesFor: (sheet: string) => TableRequest[],
  ): Promise<{ sheets: string[]; models: Map<string, SheetModel> }> {
    const models = new Map<string, SheetModel>();
    for (const req of tablesFor(CSV_SHEET)) {
      models.set(req.key, { ...(await this.read(path, req.spec)), table: req.table });
    }
    return { sheets: [CSV_SHEET], models };
  }

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

    // `headerRow` names the line above the data, so 0 means the file opens
    // straight into data with no header line at all. Columns are named after
    // themselves then, which is what the Excel reader does with a blank header
    // row, so a key written for one file works for the other.
    const width = records.reduce((w: number, r: string[]) => Math.max(w, r.length), 0);
    const headerLine =
      spec.headerRow > 0
        ? records[spec.headerRow - 1]
        : Array.from({ length: width }, (_, i) =>
            records.some((r: string[]) => String(r[i] ?? '').trim() !== '')
              ? `Column ${numToCol(i + 1)}`
              : '',
          );
    if (!headerLine) {
      throw new Error(`sheet-verify: no header row at line ${spec.headerRow} of ${path}`);
    }

    // A CSV holds one table, so a column bound is rarely useful here -- but the
    // option is on the shared spec, and silently ignoring one that was written
    // would be worse than honouring it.
    const bound = columnRange(spec.columns, headerLine.length);

    const headers = headerLine
      .map((name, i) => ({ name: String(name ?? '').trim(), colNum: i + 1 }))
      .filter((h) => h.name !== '' && h.colNum >= bound.from && h.colNum <= bound.to);

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
