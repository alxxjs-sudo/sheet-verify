import ExcelJS from 'exceljs';
import type {
  CellKind, CellValue, ResolvedSpec, SheetModel, TableRequest, WorkbookReader,
} from './types.js';
import { buildModel, type RawCell } from './model.js';
import { columnRange, numToCol } from './a1.js';
import { openWorkbook } from './open-xlsx.js';

/** ExcelJS ValueType numeric enum. */
const VT = {
  Null: 0, Merge: 1, Number: 2, String: 3, Date: 4, Hyperlink: 5,
  Formula: 6, SharedString: 7, RichText: 8, Boolean: 9, Error: 10,
} as const;

function kindOf(type: number): CellKind {
  switch (type) {
    case VT.Number: return 'number';
    case VT.Date: return 'date';
    case VT.Boolean: return 'boolean';
    case VT.Error: return 'error';
    case VT.Hyperlink: return 'hyperlink';
    case VT.Formula: return 'formula';
    case VT.Null:
    case VT.Merge: return 'empty';
    default: return 'string';
  }
}

/** Flattens ExcelJS's several object-shaped cell values to a scalar. */
function scalar(v: unknown): CellValue {
  if (v === null || v === undefined) return null;
  // A date-formatted cell holding a formula that returned "" decodes to an
  // Invalid Date. It is an empty cell, and letting one into the model makes
  // every later toISOString() throw.
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v;
  if (typeof v === 'object') {
    const o = v as Record<string, any>;
    if ('richText' in o) return (o.richText as any[]).map((t) => t.text).join('');
    if ('hyperlink' in o) return o.text ?? o.hyperlink ?? null;
    if ('error' in o) return String(o.error);
    if ('result' in o) return scalar(o.result);
    return null;
  }
  return v as CellValue;
}

/** The first non-empty text literal in a formula, e.g. `="# Risks "&IF(...)`. */
const FIRST_LITERAL = /"([^"]+)"/;

/**
 * What to call a column.
 *
 * Reports build header text with formulas -- `="# Risks "&IF($A15="Buildings",
 * "(Buildings)", "")` -- and a file straight from the generator holds no
 * computed result for them, so the name reads as empty and the whole column
 * goes uncompared. On one sheet that was six columns of six thousand cells
 * each.
 *
 * The formula is the same on both sides, so its leading literal identifies the
 * column just as well as the rendered text would -- and it is what the header
 * actually says. A formula with no literal gives no name at all; see the note
 * on the fallback below for why. Repeats are disambiguated downstream.
 */
export function headerName(cell: ExcelJS.Cell): string {
  const rendered = String(scalar(cell.value) ?? '').trim();
  if (rendered) return rendered;

  const formula = (cell.formula as string | undefined) || '';
  if (!formula) return '';
  // The leading literal, or nothing. Falling back to the formula itself used to
  // seem harmless -- "stable if ugly" -- and is the opposite of stable: it
  // names the column `SUM(B15:B18)`, which moves the moment a row is inserted
  // above it, so both files report the column as one removed and one added.
  // Worse, it makes a row of totals look like a row of twenty-five names, so
  // the header search prefers it to the real header row above.
  //
  // A formula that builds header text contains that text. One that computes a
  // number does not, and has no name to give.
  return (FIRST_LITERAL.exec(formula)?.[1] ?? '').trim();
}

/**
 * What a cell contributes to detection's picture of which cells are occupied.
 *
 * Distinct from `headerName` because the two answer different questions. A
 * formula with no literal has no *name*, but it is not an empty cell: a row of
 * nothing but formulas reading as blank splits the block it sits in and strands
 * the real header row above the split.
 */
export function cellToken(cell: ExcelJS.Cell): string {
  const rendered = String(scalar(cell.value) ?? '').trim();
  if (rendered) return rendered;
  return ((cell.formula as string | undefined) || '').trim();
}

/** Whether a column holds anything at all between two rows. */
function columnHasData(ws: ExcelJS.Worksheet, colNum: number, from: number, to: number): boolean {
  for (let r = from; r <= to; r++) {
    const cell = ws.getRow(r).getCell(colNum);
    if (cell.formula) return true;
    const v = scalar(cell.value);
    if (v !== null && String(v).trim() !== '') return true;
  }
  return false;
}

/** Extracts one table from an already-parsed worksheet. */
function modelFrom(ws: ExcelJS.Worksheet, path: string, spec: ResolvedSpec): SheetModel {
  // endRow bounds a table that stops short of the sheet's last row, so an
  // info block does not swallow the data table underneath it.
  const last = spec.endRow > 0 ? Math.min(spec.endRow, ws.rowCount) : ws.rowCount;

  const headers: { name: string; colNum: number }[] = [];
  // `headerRow` names the row above the data, so a table whose data starts on
  // row 1 has no header row to read -- there is no row 0 to ask for. Every
  // column is then named after itself, exactly as a blank header row already
  // names them, and row 1 is data like any other.
  const headerRow = spec.headerRow > 0 ? ws.getRow(spec.headerRow) : null;

  // A table can be bounded left and right as well as top and bottom, for a
  // sheet that puts two of them side by side. Unbounded is the default and
  // means the whole width.
  const { from, to } = columnRange(spec.columns, ws.columnCount);

  for (let colNum = from; colNum <= Math.min(to, ws.columnCount); colNum++) {
    const name = headerRow ? headerName(headerRow.getCell(colNum)) : '';
    if (name) {
      headers.push({ name, colNum });
      continue;
    }
    // A column can hold data and still have no header -- the county beside
    // "Geography Level" is six thousand names under a blank heading. Skipping
    // those left them uncompared and unusable as part of a key. Naming them
    // after the column is positional, so it only holds while the columns stay
    // put; a real heading always wins over it.
    if (columnHasData(ws, colNum, spec.headerRow + 1, last)) {
      headers.push({ name: `Column ${numToCol(colNum)}`, colNum });
    }
  }

  if (!headers.length) {
    const where = spec.columns ? ` in columns ${spec.columns}` : '';
    const at =
      spec.headerRow > 0
        ? `on row ${spec.headerRow}`
        : 'from row 1, which headerRow 0 reads as data with no header row';
    throw new Error(
      `sheet-verify: no headers found ${at}${where} of "${ws.name}" in ${path}`,
    );
  }

  const rows: { rowNum: number; cells: Map<number, RawCell> }[] = [];
  for (let r = spec.headerRow + 1; r <= last; r++) {
    const row = ws.getRow(r);
    const cells = new Map<number, RawCell>();
    for (const { colNum } of headers) {
      const cell = row.getCell(colNum);

      // A merged range holds its value once, on the master. ExcelJS hands the
      // same value to every slave, so reading them repeats it down and across
      // the range -- one changed banner arrives as a dozen changed cells, and
      // an exclusion naming the master leaves the rest still reporting. Layer
      // 2 and detection both already read a merged range this way.
      if (cell.isMerged && cell.master !== cell) {
        cells.set(colNum, {
          address: cell.address, kind: 'empty', value: null, formula: null,
        });
        continue;
      }

      // IMPORTANT: the `formula` getter resolves shared (filled-down) formulas
      // back to translated text. Reading `cell.value.formula` instead returns
      // only the master's address for every slave cell.
      const formula = (cell.formula as string | undefined) || null;
      cells.set(colNum, {
        address: cell.address,
        kind: kindOf(cell.type as number),
        value: scalar(cell.value),
        formula,
      });
    }
    rows.push({ rowNum: r, cells });
  }

  return buildModel({ source: path, sheet: ws.name, headers, rows, spec });
}

export class ExcelReader implements WorkbookReader {
  readonly extensions = ['.xlsx', '.xlsm'];

  async read(path: string, spec: ResolvedSpec): Promise<SheetModel> {
    const wb = await openWorkbook(path);

    const ws =
      typeof spec.sheet === 'number' ? wb.worksheets[spec.sheet] : wb.getWorksheet(spec.sheet);
    if (!ws) {
      const names = wb.worksheets.map((w) => w.name).join(', ');
      throw new Error(`sheet-verify: sheet ${JSON.stringify(spec.sheet)} not found in ${path}. Available: ${names}`);
    }

    return modelFrom(ws, path, spec);
  }

  async readWorkbook(
    path: string,
    tablesFor: (sheet: string) => TableRequest[],
  ): Promise<{ sheets: string[]; models: Map<string, SheetModel> }> {
    const wb = await openWorkbook(path);

    const sheets: string[] = [];
    const models = new Map<string, SheetModel>();
    for (const ws of wb.worksheets) {
      sheets.push(ws.name);
      for (const req of tablesFor(ws.name)) {
        models.set(req.key, { ...modelFrom(ws, path, req.spec), table: req.table });
      }
    }
    return { sheets, models };
  }
}
