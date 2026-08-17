import ExcelJS from 'exceljs';
import type {
  CellKind, CellValue, ResolvedSpec, SheetModel, TableRequest, WorkbookReader,
} from './types.js';
import { buildModel, type RawCell } from './model.js';
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
  if (v instanceof Date) return v;
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

/** Extracts one table from an already-parsed worksheet. */
function modelFrom(ws: ExcelJS.Worksheet, path: string, spec: ResolvedSpec): SheetModel {
  const headers: { name: string; colNum: number }[] = [];
  ws.getRow(spec.headerRow).eachCell({ includeEmpty: false }, (cell, colNum) => {
    const name = String(scalar(cell.value) ?? '').trim();
    if (name) headers.push({ name, colNum });
  });
  if (!headers.length) {
    throw new Error(`sheet-verify: no headers found on row ${spec.headerRow} of "${ws.name}" in ${path}`);
  }

  // endRow bounds a table that stops short of the sheet's last row, so an
  // info block does not swallow the data table underneath it.
  const last = spec.endRow > 0 ? Math.min(spec.endRow, ws.rowCount) : ws.rowCount;

  const rows: { rowNum: number; cells: Map<number, RawCell> }[] = [];
  for (let r = spec.headerRow + 1; r <= last; r++) {
    const row = ws.getRow(r);
    const cells = new Map<number, RawCell>();
    for (const { colNum } of headers) {
      const cell = row.getCell(colNum);
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
