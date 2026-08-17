import ExcelJS from 'exceljs';
import type { Cell, CellValue, ResolvedSpec } from './types.js';
import type { ComparedTable } from './workbook.js';
import { canonHeader, displayKey, toleranceFor } from './model.js';

/**
 * What happened to one cell. `match` is only ever emitted under the 'all'
 * scope -- by default the ledger records what changed, since a run where
 * nothing changed has nothing to read.
 */
export type CellStatus =
  | 'match'
  | 'value-differs'
  | 'formula-differs'
  | 'type-differs'
  | 'within-tolerance'
  | 'ignored-column'
  | 'ignored-row'
  | 'row-added'
  | 'row-removed'
  | 'column-added'
  | 'column-removed';

export type LedgerScope =
  /** Only cells that were not a plain match. The default. */
  | 'differences'
  /** Every cell in scope, matches included. Grows with rows x columns. */
  | 'all'
  | 'none';

export interface LedgerRow {
  sheet: string;
  table: string;
  rowKey: string;
  column: string;
  status: CellStatus;
  /** 'yes' for a cause, 'no' for something downstream of one, '' otherwise. */
  rootCause: '' | 'yes' | 'no';
  baselineAddress: string;
  actualAddress: string;
  baselineValue: CellValue;
  actualValue: CellValue;
  delta: number | null;
  tolerance: number | null;
  baselineFormula: string | null;
  actualFormula: string | null;
}

interface ColumnDef {
  key: keyof LedgerRow;
  header: string;
  width: number;
}

/** Widths are set so nothing is truncated at the default zoom. */
const COLUMNS: ColumnDef[] = [
  { key: 'sheet', header: 'Sheet', width: 18 },
  { key: 'table', header: 'Table', width: 14 },
  { key: 'rowKey', header: 'Row key', width: 26 },
  { key: 'column', header: 'Column', width: 20 },
  { key: 'status', header: 'Status', width: 18 },
  { key: 'rootCause', header: 'Root cause', width: 12 },
  { key: 'baselineAddress', header: 'Golden cell', width: 12 },
  { key: 'actualAddress', header: 'Actual cell', width: 12 },
  { key: 'baselineValue', header: 'Golden value', width: 20 },
  { key: 'actualValue', header: 'Actual value', width: 20 },
  { key: 'delta', header: 'Delta', width: 14 },
  { key: 'tolerance', header: 'Tolerance', width: 11 },
  { key: 'baselineFormula', header: 'Golden formula', width: 30 },
  { key: 'actualFormula', header: 'Actual formula', width: 30 },
];

const num = (n: number): number => Number(n.toPrecision(12));

function normalised(c: Cell, spec: ResolvedSpec): string | null {
  if (!c.formula) return null;
  if (spec.formulaMode === 'a1') return c.formula;
  if (spec.formulaMode === 'r1c1') return c.r1c1;
  return c.headerRef;
}

function statusOf(b: Cell, n: Cell, spec: ResolvedSpec, tol: number): CellStatus {
  if (spec.compareFormulas && normalised(b, spec) !== normalised(n, spec)) {
    return 'formula-differs';
  }
  if (b.kind !== n.kind && b.kind !== 'empty' && n.kind !== 'empty') return 'type-differs';

  if (typeof b.value === 'number' && typeof n.value === 'number') {
    const delta = Math.abs(b.value - n.value);
    if (delta === 0) return 'match';
    return delta <= tol ? 'within-tolerance' : 'value-differs';
  }
  return String(b.value ?? '') === String(n.value ?? '') ? 'match' : 'value-differs';
}

const deltaOf = (b?: Cell, n?: Cell): number | null =>
  typeof b?.value === 'number' && typeof n?.value === 'number'
    ? num(Math.abs(b.value - n.value))
    : null;

/** Walks the compared models and yields one row per cell in scope. */
export function* ledgerRows(
  tables: ComparedTable[],
  scope: LedgerScope = 'differences',
): Generator<LedgerRow> {
  if (scope === 'none') return;

  for (const t of tables) {
    const { base, next, spec } = t;
    const emit = (status: CellStatus) => scope === 'all' || status !== 'match';

    // A value difference is either the cause or a consequence of one. The
    // distinction is the most useful thing in the report, so the ledger
    // carries it rather than leaving it only in diff.json.
    const roots = new Set(
      t.diff.values.filter((v) => v.rootCause).map((v) => `${v.key} ${v.column}`),
    );
    const cascades = new Set(
      t.diff.values.filter((v) => !v.rootCause).map((v) => `${v.key} ${v.column}`),
    );
    const cause = (key: string, column: string): LedgerRow['rootCause'] => {
      const k = `${key} ${column}`;
      if (roots.has(k)) return 'yes';
      return cascades.has(k) ? 'no' : '';
    };

    /**
     * An excluded cell still earns a row when it actually differs -- that is
     * the evidence the exclusion is doing work. One that was excluded *and*
     * identical is as uninteresting as any other match.
     */
    const unchanged = (b?: Cell, n?: Cell): boolean =>
      !!b && !!n &&
      String(b.value ?? '') === String(n.value ?? '') &&
      (b.formula ?? '') === (n.formula ?? '');

    const canon = (h: string) => canonHeader(h, spec.looseHeaders);
    const ignoredCols = new Set(spec.ignoreColumns.map(canon));
    const nextByCanon = new Map(next.headers.map((h) => [canon(h), h]));
    const baseByCanon = new Map(base.headers.map((h) => [canon(h), h]));

    const row = (
      key: string,
      column: string,
      status: CellStatus,
      b?: Cell,
      n?: Cell,
      tolerance: number | null = null,
    ): LedgerRow => ({
      sheet: t.sheet,
      table: t.table,
      rowKey: displayKey(key, spec),
      column,
      status,
      rootCause: cause(key, column),
      baselineAddress: b?.address ?? '',
      actualAddress: n?.address ?? '',
      baselineValue: b?.value ?? null,
      actualValue: n?.value ?? null,
      delta: deltaOf(b, n),
      tolerance,
      baselineFormula: b?.formula ?? null,
      actualFormula: n?.formula ?? null,
    });

    for (const key of base.order) {
      const b = base.rows.get(key)!;
      const n = next.rows.get(key);

      if (!n) {
        for (const column of base.headers) yield row(key, column, 'row-removed', b[column]);
        continue;
      }
      if (spec.ignoreRows.includes(key)) {
        for (const column of base.headers) {
          const nextName = nextByCanon.get(canon(column));
          const nc = nextName ? n[nextName] : undefined;
          if (scope !== 'all' && unchanged(b[column], nc)) continue;
          yield row(key, column, 'ignored-row', b[column], nc);
        }
        continue;
      }

      for (const column of base.headers) {
        const nextName = nextByCanon.get(canon(column));

        if (ignoredCols.has(canon(column))) {
          const nc = nextName ? n[nextName] : undefined;
          if (scope === 'all' || !unchanged(b[column], nc)) {
            yield row(key, column, 'ignored-column', b[column], nc);
          }
          continue;
        }
        if (!nextName) {
          yield row(key, column, 'column-removed', b[column]);
          continue;
        }

        const bc = b[column];
        const nc = n[nextName];
        if (!bc || !nc) continue;
        const tol = toleranceFor(spec, nextName);
        const status = statusOf(bc, nc, spec, tol);
        if (emit(status)) yield row(key, column, status, bc, nc, tol);
      }
    }

    // Rows and columns present only in the new output.
    for (const key of next.order) {
      if (base.rows.has(key)) continue;
      const n = next.rows.get(key)!;
      for (const column of next.headers) {
        yield row(key, column, 'row-added', undefined, n[column]);
      }
    }
    for (const column of next.headers) {
      if (baseByCanon.has(canon(column))) continue;
      for (const key of next.order) {
        if (!base.rows.has(key)) continue;
        yield row(key, column, 'column-added', undefined, next.rows.get(key)![column]);
      }
    }
  }
}

/* --- CSV ----------------------------------------------------------------- */

const text = (v: CellValue): string => {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'number') return String(num(v));
  return String(v);
};

/** RFC 4180 quoting. */
function csv(fields: (CellValue | number)[]): string {
  return fields
    .map((f) => {
      const s = text(f);
      return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    })
    .join(',');
}

/**
 * CSV form, yielded line by line so a large ledger can be streamed to disk
 * rather than built in memory. Preferred over xlsx for the 'all' scope on a
 * big case, where the row count runs into the millions.
 */
export function* ledgerCsvLines(
  tables: ComparedTable[],
  scope: LedgerScope = 'differences',
): Generator<string> {
  if (scope === 'none') return;
  yield csv(COLUMNS.map((c) => c.header));
  for (const r of ledgerRows(tables, scope)) {
    yield csv(COLUMNS.map((c) => r[c.key] as CellValue));
  }
}

/** The whole CSV as one string. Prefer `ledgerCsvLines` for large cases. */
export function formatLedger(
  tables: ComparedTable[],
  scope: LedgerScope = 'differences',
): string {
  return [...ledgerCsvLines(tables, scope)].join('\n');
}

/* --- xlsx ---------------------------------------------------------------- */

const INK = 'FF12151C';
const MUTED = 'FF6B7280';

/**
 * The header is painted explicitly rather than left to the table theme.
 * Themes vary in whether they give the header band a dark fill, and a theme
 * that does plus dark text of our own leaves the labels invisible.
 */
const HEADER_FILL = 'FF1F2937';
const HEADER_TEXT = 'FFFFFFFF';

function styleHeader(ws: ExcelJS.Worksheet, columns: number): void {
  const row = ws.getRow(1);
  row.font = { bold: true, color: { argb: HEADER_TEXT } };
  row.alignment = { vertical: 'middle' };
  row.height = 20;
  for (let c = 1; c <= columns; c++) {
    row.getCell(c).fill = {
      type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_FILL },
    };
  }
}

/** Excel's hard limit, less the header and a row to report the truncation. */
const MAX_DATA_ROWS = 1_048_574;

const INVALID_SHEET_CHARS = /[\\/?*[\]:]/g;

/** Excel rejects some characters and anything past 31 chars, and duplicates. */
function worksheetName(label: string, used: Set<string>): string {
  const base = label.replace(INVALID_SHEET_CHARS, '-').trim().slice(0, 31) || 'Sheet';
  let name = base;
  for (let i = 2; used.has(name.toLowerCase()); i++) {
    const suffix = ` (${i})`;
    name = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(name.toLowerCase());
  return name;
}

/** Status colours: ink on a tint, so the column reads as a legend. */
const STATUS_STYLE: Record<CellStatus, { font: string; fill: string }> = {
  'value-differs': { font: 'FF9F2F26', fill: 'FFFBE9E7' },
  'formula-differs': { font: 'FF9F2F26', fill: 'FFFBE9E7' },
  'type-differs': { font: 'FF9F2F26', fill: 'FFFBE9E7' },
  'within-tolerance': { font: 'FF8A5A12', fill: 'FFFDF3DE' },
  'row-added': { font: 'FF2F5E44', fill: 'FFE8F3EC' },
  'column-added': { font: 'FF2F5E44', fill: 'FFE8F3EC' },
  'row-removed': { font: 'FF8A5A12', fill: 'FFFDF3DE' },
  'column-removed': { font: 'FF8A5A12', fill: 'FFFDF3DE' },
  'ignored-column': { font: MUTED, fill: 'FFF2F3F5' },
  'ignored-row': { font: MUTED, fill: 'FFF2F3F5' },
  match: { font: MUTED, fill: 'FFF2F3F5' },
};

/**
 * Writes the ledger as a formatted worksheet: a real Excel table, so it
 * arrives with filter buttons and banded rows, with the header frozen and
 * numbers written as numbers so sorting and filtering behave.
 */
export async function writeLedgerWorkbook(
  path: string,
  tables: ComparedTable[],
  scope: LedgerScope = 'differences',
): Promise<number> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'sheet-verify';
  const ws = wb.addWorksheet('Cells', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  COLUMNS.forEach((c, i) => { ws.getColumn(i + 1).width = c.width; });

  const rows = [...ledgerRows(tables, scope)];
  const values = rows.map((r) =>
    COLUMNS.map((c) => {
      const v = r[c.key];
      return v === null || v === undefined ? '' : (v as string | number);
    }),
  );

  if (values.length) {
    ws.addTable({
      name: 'Cells',
      ref: 'A1',
      headerRow: true,
      style: { theme: 'TableStyleLight8', showRowStripes: true },
      columns: COLUMNS.map((c) => ({ name: c.header, filterButton: true })),
      rows: values,
    });
  } else {
    // addTable rejects an empty body, so a clean run gets the header alone.
    ws.addRow(COLUMNS.map((c) => c.header));
    ws.autoFilter = { from: 'A1', to: { row: 1, column: COLUMNS.length } };
  }

  styleHeader(ws, COLUMNS.length);

  const statusCol = COLUMNS.findIndex((c) => c.key === 'status') + 1;
  const causeCol = COLUMNS.findIndex((c) => c.key === 'rootCause') + 1;
  const formulaCols = COLUMNS
    .map((c, i) => (c.key.endsWith('Formula') ? i + 1 : 0))
    .filter(Boolean);

  rows.forEach((r, i) => {
    const rowNum = i + 2;
    const style = STATUS_STYLE[r.status];
    const cell = ws.getCell(rowNum, statusCol);
    cell.font = { color: { argb: style.font }, bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: style.fill } };

    if (r.rootCause) {
      ws.getCell(rowNum, causeCol).font = {
        color: { argb: r.rootCause === 'yes' ? 'FF9F2F26' : MUTED },
        bold: r.rootCause === 'yes',
      };
    }
    // Formulas are code; a monospace face makes a changed operand visible.
    for (const c of formulaCols) {
      ws.getCell(rowNum, c).font = { name: 'Consolas', size: 10, color: { argb: INK } };
    }
  });

  await wb.xlsx.writeFile(path);
  return rows.length;
}

/**
 * The minimal record of what was compared: which cell on each side, what each
 * held, and the verdict. No deltas, tolerances or formulas -- those belong in
 * the differences ledger, which is short enough to carry them.
 */
const COMPARED_COLUMNS: ColumnDef[] = [
  { key: 'rowKey', header: 'Row key', width: 26 },
  { key: 'column', header: 'Column', width: 22 },
  { key: 'baselineAddress', header: 'Golden cell', width: 12 },
  { key: 'actualAddress', header: 'Actual cell', width: 12 },
  { key: 'baselineValue', header: 'Golden value', width: 22 },
  { key: 'actualValue', header: 'Actual value', width: 22 },
  { key: 'status', header: 'Status', width: 17 },
];

/**
 * Every cell compared, one worksheet per compared table.
 *
 * Split because this is the file that gets large: a workbook of five sheets is
 * five tabs to scan rather than one sheet of everything, and each tab stays
 * clear of Excel's row ceiling on its own. Returns the row count per tab.
 */
export async function writeComparedWorkbook(
  path: string,
  tables: ComparedTable[],
): Promise<Record<string, number>> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'sheet-verify';

  const used = new Set<string>();
  const counts: Record<string, number> = {};

  for (const t of tables) {
    const name = worksheetName(t.label, used);
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    COMPARED_COLUMNS.forEach((c, i) => { ws.getColumn(i + 1).width = c.width; });
    ws.addRow(COMPARED_COLUMNS.map((c) => c.header));

    let n = 0;
    for (const r of ledgerRows([t], 'all')) {
      if (n >= MAX_DATA_ROWS) {
        ws.addRow([`… truncated at ${MAX_DATA_ROWS.toLocaleString('en-US')} rows`]);
        break;
      }
      ws.addRow(
        COMPARED_COLUMNS.map((c) => {
          const v = r[c.key];
          return v === null || v === undefined ? '' : (v as string | number);
        }),
      );
      n++;
    }

    ws.autoFilter = { from: 'A1', to: { row: 1, column: COMPARED_COLUMNS.length } };
    styleHeader(ws, COMPARED_COLUMNS.length);
    counts[name] = n;
  }

  // A workbook with no worksheets cannot be written.
  if (!tables.length) wb.addWorksheet('Compared');

  await wb.xlsx.writeFile(path);
  return counts;
}
