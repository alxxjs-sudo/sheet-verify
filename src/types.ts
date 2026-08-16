/** Public contract for sheet-verify. */

export type CellValue = string | number | boolean | Date | null;

export type CellKind =
  | 'empty'
  | 'string'
  | 'number'
  | 'boolean'
  | 'date'
  | 'formula'
  | 'hyperlink'
  | 'error';

export interface Cell {
  /** A1 address in the source sheet, e.g. "H6". Kept for reporting only. */
  address: string;
  kind: CellKind;
  /** Literal value, or the cached result for a formula cell. */
  value: CellValue;
  /** Original A1 formula text, without leading "=". Null for non-formula cells. */
  formula: string | null;
  /** Formula with refs rewritten to R1C1, relative to this cell. */
  r1c1: string | null;
  /** Formula with refs resolved to header names. Invariant under column moves. */
  headerRef: string | null;
}

export type Row = Record<string, Cell>;

export interface SheetModel {
  /** Source file path, for reporting. */
  source: string;
  /** Worksheet name. "" for CSV. */
  sheet: string;
  /** Header names in positional order. */
  headers: string[];
  /** header name -> 1-based column number in the source. */
  headerIndex: Map<string, number>;
  /** composite key -> row. */
  rows: Map<string, Row>;
  /** Key order as encountered, so reports can follow file order. */
  order: string[];
  /** Keys seen more than once. Ambiguous rows are excluded from `rows`. */
  duplicateKeys: string[];
  /** Header names appearing more than once. */
  duplicateHeaders: string[];
  /** Formula cells that carry no cached value. See `requireCachedValues`. */
  uncachedFormulaCells: string[];
  /** CSV only. Encoding-level details that drift silently between releases. */
  dialect?: CsvDialect;
}

export interface CsvDialect {
  delimiter: string;
  bom: boolean;
  lineEnding: 'CRLF' | 'LF' | 'CR' | 'mixed' | 'none';
}

export type FormulaMode = 'a1' | 'r1c1' | 'header';

export interface SheetSpec {
  /** Worksheet name or 0-based index. Ignored for CSV. Default 0. */
  sheet?: string | number;
  /** 1-based row holding the headers. Default 1. */
  headerRow?: number;
  /**
   * Column(s) forming the row identity. Multiple columns form a composite key.
   * Rows with a blank key are skipped.
   */
  keyColumns: string[];
  /** Joins composite key parts. Default "␟" (unit separator). */
  keySeparator?: string;
  /**
   * Numeric tolerance. A single number applies to every column; a record
   * applies per column name, with `*` as the fallback. Default 0 (exact).
   */
  tolerance?: number | Record<string, number>;
  /** Columns excluded from all comparison (timestamps, run ids, ...). */
  ignoreColumns?: string[];
  /** Compare formulas at all. Default true. */
  compareFormulas?: boolean;
  /** How formulas are normalised before comparison. Default 'header'. */
  formulaMode?: FormulaMode;
  /**
   * Fail when a formula cell has no cached value. Without this, a
   * value comparison silently passes on empty. Default true.
   */
  requireCachedValues?: boolean;
  /** Trim leading/trailing whitespace on string values. Default true. */
  trimStrings?: boolean;
  /** Match header names case- and whitespace-insensitively. Default true. */
  looseHeaders?: boolean;
  /**
   * Treat schema changes (added/removed/moved columns, added/removed rows)
   * as failures rather than review items. Default false.
   */
  strictSchema?: boolean;
  /** Assertions that must hold for any correct output, baseline or not. */
  invariants?: Invariant[];
  /** CSV-specific parsing options. Ignored for Excel sources. */
  csv?: CsvOptions;
}

export interface CsvOptions {
  /** Field delimiter. Detected from the header line when omitted. */
  delimiter?: string;
  /**
   * How text becomes numbers. CSV carries no types, so this decides whether
   * "1000" compares as a number or a string.
   *  - 'auto'           numeric-looking values, except those with leading zeros
   *  - 'tolerance-only' only columns with an explicit tolerance
   *  - 'none'           everything stays text
   * Default 'auto'.
   */
  numeric?: 'auto' | 'tolerance-only' | 'none';
  /** Report BOM / delimiter / line-ending differences as failures. Default false. */
  strictDialect?: boolean;
}

/** A check that does not need a baseline. Runs against the new output. */
export interface Invariant {
  name: string;
  check(model: SheetModel): InvariantFailure[] | void;
}

export interface InvariantFailure {
  invariant: string;
  key?: string;
  column?: string;
  address?: string;
  detail: string;
}

export interface MovedColumn {
  column: string;
  from: number;
  to: number;
}

export interface ValueDiff {
  key: string;
  column: string;
  address: string;
  base: CellValue;
  next: CellValue;
  /** Absolute delta for numeric pairs. */
  delta?: number;
  /** True when this cell's formula also changed, i.e. likely the cause. */
  formulaChanged: boolean;
  /** True when no upstream cell in this row changed, i.e. likely a root cause. */
  rootCause: boolean;
}

export interface TypeDiff {
  key: string;
  column: string;
  address: string;
  baseKind: CellKind;
  nextKind: CellKind;
  value: CellValue;
}

export interface FormulaDiff {
  key: string;
  column: string;
  address: string;
  /** Normalised forms actually compared. */
  base: string | null;
  next: string | null;
  /** Original A1 text, for the report. */
  baseA1: string | null;
  nextA1: string | null;
}

export interface DiffResult {
  ok: boolean;
  /** True when only review-level changes were found and none are defects. */
  reviewOnly: boolean;
  base: { source: string; sheet: string; rows: number; columns: number };
  next: { source: string; sheet: string; rows: number; columns: number };
  schema: {
    added: string[];
    removed: string[];
    moved: MovedColumn[];
    /** Columns compared: present in both, minus ignored. */
    compared: string[];
  };
  rows: {
    added: string[];
    removed: string[];
    compared: number;
    duplicateKeysBase: string[];
    duplicateKeysNext: string[];
  };
  values: ValueDiff[];
  types: TypeDiff[];
  formulas: FormulaDiff[];
  invariants: InvariantFailure[];
  /** Structural problems that make the comparison itself untrustworthy. */
  errors: string[];
}

/** Anything that can produce a SheetModel. Keeps ExcelJS swappable. */
export interface SheetReader {
  /** File extensions handled, lowercase, with dot. */
  extensions: string[];
  read(path: string, spec: ResolvedSpec): Promise<SheetModel>;
}

/**
 * A reader for formats holding several sheets. Separate from SheetReader
 * because the whole point is parsing the file *once* for every sheet --
 * calling read() per sheet re-parses the workbook each time.
 */
export interface WorkbookReader extends SheetReader {
  /**
   * Returns every sheet name in workbook order, plus models for the sheets
   * `specFor` accepts. Sheets it returns null for are listed but not built,
   * so unconfigured tabs cost nothing.
   */
  readWorkbook(
    path: string,
    specFor: (sheet: string) => ResolvedSpec | null,
  ): Promise<{ sheets: string[]; models: Map<string, SheetModel> }>;
}

export type ResolvedSpec = Required<
  Omit<SheetSpec, 'tolerance' | 'invariants' | 'sheet' | 'csv'>
> & {
  sheet: string | number;
  tolerance: Record<string, number>;
  invariants: Invariant[];
  csv: Required<CsvOptions>;
};

/* --- workbook level ------------------------------------------------------ */

/**
 * Per-sheet spec inside a workbook. `sheet` is dropped because the worksheet
 * name is the key, and `keyColumns` is optional because it can come from
 * `defaults` -- a sheet that ends up without one is reported as not compared
 * rather than throwing.
 */
export type WorkbookSheetSpec = Omit<SheetSpec, 'sheet' | 'keyColumns'> & {
  keyColumns?: string[];
};

export interface WorkbookSpec {
  /**
   * Applied to every sheet, then overridden per sheet. Merging is per-field:
   * `tolerance` records merge, `ignoreColumns` and `invariants` concatenate,
   * everything else is replaced by the per-sheet value.
   */
  defaults?: WorkbookSheetSpec;
  /** Per-sheet spec, keyed by worksheet name. Names match case-insensitively. */
  sheets?: Record<string, WorkbookSheetSpec>;
  /** Sheets excluded entirely -- scratch tabs, lookup data, notes. */
  ignoreSheets?: string[];
  /**
   * Treat added and unconfigured sheets as failures rather than review items.
   * Removed sheets are always failures. Default false.
   */
  strictSheets?: boolean;
}

export type SheetStatus =
  /** Present in both and configured: the four layers ran. */
  | 'compared'
  /** In the new output only. Nothing to compare against, so it is only noted. */
  | 'added'
  /** In the baseline only. Output that used to be produced has vanished. */
  | 'removed'
  /** Present in both, but no keyColumns resolved. A visible coverage gap. */
  | 'skipped'
  /** Excluded by `ignoreSheets`. */
  | 'ignored';

export interface SheetOutcome {
  sheet: string;
  status: SheetStatus;
  /** Set only when status is 'compared'. */
  diff?: DiffResult;
  /** Why the sheet was not compared. */
  reason?: string;
}

export interface MovedSheet {
  sheet: string;
  from: number;
  to: number;
}

export interface WorkbookDiffResult {
  ok: boolean;
  /** True when only review-level changes were found and none are defects. */
  reviewOnly: boolean;
  base: { source: string; sheets: string[] };
  next: { source: string; sheets: string[] };
  /** The sheet-level layer: what changed about the tabs themselves. */
  sheetSchema: { added: string[]; removed: string[]; moved: MovedSheet[] };
  /** One entry per sheet, in baseline order, with additions last. */
  sheets: SheetOutcome[];
  /** Structural problems that make the comparison itself untrustworthy. */
  errors: string[];
}
