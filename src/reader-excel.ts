import ExcelJS from 'exceljs';
import type {
  CellKind, CellValue, ResolvedSpec, SheetModel, TableRequest, WorkbookReader,
} from './types.js';
import { buildModel, type RawCell } from './model.js';
import { numToCol } from './a1.js';
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
 * The formula is the same on both sides, so it identifies the column just as
 * well as its rendered text would. The leading literal is used where there is
 * one because it is what the header actually says; failing that the formula
 * itself, which is stable if ugly. Repeats are disambiguated downstream.
 */
export function headerName(cell: ExcelJS.Cell): string {
  const rendered = String(scalar(cell.value) ?? '').trim();
  if (rendered) return rendered;

  const formula = (cell.formula as string | undefined) || '';
  if (!formula) return '';
  return (FIRST_LITERAL.exec(formula)?.[1] ?? formula).trim();
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
  const headerRow = ws.getRow(spec.headerRow);

  for (let colNum = 1; colNum <= ws.columnCount; colNum++) {
    const name = headerName(headerRow.getCell(colNum));
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
    throw new Error(`sheet-verify: no headers found on row ${spec.headerRow} of "${ws.name}" in ${path}`);
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
