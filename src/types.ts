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
  /** Table within the sheet, when the sheet holds more than one. */
  table?: string;
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
  /**
   * 1-based row holding the headers. Default 1.
   *
   * Read it as "the row above the data", which is what it has always meant --
   * the table runs from `headerRow + 1`. That makes `0` a legal value and not
   * a special case: the data starts at row 1 and there is no header row.
   * Columns are then named after themselves -- `Column A`, `Column B` -- the
   * same naming a blank header row already produces.
   *
   * Set it to 0 for a block with no headers of its own, the usual case being a
   * key-value block of label/value pairs sitting at the top of a sheet. Picking
   * a header row there costs twice over: that row stops being data, and the
   * value column takes its name from a *value*. Where the value is the report's
   * own name or id it differs between any two runs, so the column pairs with
   * nothing in the other file and every row of the block reports as one column
   * removed and another added.
   */
  headerRow?: number;
  /**
   * 1-based last row of the table, inclusive. Bounds a table that does not
   * run to the bottom of the sheet -- an info block above the data, say.
   * Default 0, meaning read to the last row. Set for you when a sheet
   * declares several tables.
   */
  endRow?: number;
  /**
   * Columns this table occupies, as a range of letters: `"H:J"`. Default `""`,
   * meaning every column on the sheet.
   *
   * Reports put tables side by side -- a definitions table in H:J beside a
   * key-value block in A:B, both starting on row 1. A blank row cannot
   * separate those, so without a column bound they are read as one table:
   * the two header rows fuse, the shorter table's rows read as blank, and a
   * key named in one of them can be found in the other. Set for you when
   * detection sees a gap of blank columns wide enough to be a separation
   * rather than a spacer.
   *
   * Only the named columns are read. Everything outside is another table's,
   * or nobody's.
   */
  columns?: string;
  /**
   * Column(s) forming the row identity. Multiple columns form a composite key.
   * Rows with a blank key are skipped.
   */
  keyColumns: string[];
  /**
   * Match rows by their position in the table rather than by a key. Set when
   * `keyColumns` is empty, which is the only time it is allowed.
   *
   * A key is always better: it survives rows being added, removed or reordered,
   * which is the whole reason this library exists. But plenty of real tables
   * have nothing that identifies a row -- a breakdown by state and county
   * carries subtotal rows with both blank -- and refusing to compare those at
   * all leaves most of the workbook unchecked. When both sides hold the same
   * number of rows, position identifies them exactly.
   *
   * The cost is that an inserted row shifts every row under it, so one change
   * reads as many. The report says which tables were matched this way.
   */
  matchRowsByPosition?: boolean;
  /**
   * A blank part of a key inherits the value above it. Default false.
   *
   * Reports write a grouping column once, at the top of its group, and leave
   * it blank on every row beneath -- a portfolio name over three hundred loss
   * rows. Read literally that column is empty almost everywhere, so it adds
   * nothing to a key, and the rows of one group collide with the rows of the
   * next: on one real sheet, "Portfolio + Event ID" gave 485 distinct keys
   * across 997 rows because the same events recur under every portfolio.
   *
   * Filling downward reads the sheet the way a person does -- this row is
   * under that heading, so it belongs to it. It also gives a group's total
   * row a key, since those carry the heading and nothing else.
   *
   * Only key building is affected. No cell value is invented, and nothing
   * filled here is ever compared.
   */
  fillKeyDown?: boolean;
  /** Joins composite key parts. Default "␟" (unit separator). */
  keySeparator?: string;
  /**
   * Numeric tolerance. A single number applies to every column; a record
   * applies per column name, with `*` as the fallback. Default 0 (exact).
   */
  tolerance?: number | Record<string, number>;
  /**
   * Numeric tolerance as a proportion of the value, applied alongside
   * `tolerance` -- a cell passes if it is close enough by either. Same shape:
   * a single number for every column, or a record keyed by column name with
   * `*` as the fallback. Default 0 (off, and the comparison is exactly what it
   * was before this existed).
   *
   * For recalculation drift in a workbook, 1e-12 is a reasonable starting
   * point: Excel keeps 15 significant digits, so drift lands around 1e-15 to
   * 1e-14 of the value whatever its magnitude, and 1e-12 clears that by a
   * comfortable margin while staying far below any difference a person made.
   */
  relativeTolerance?: number | Record<string, number>;
  /** Columns excluded from all comparison (timestamps, run ids, ...). */
  ignoreColumns?: string[];
  /**
   * Rows excluded from all comparison, by key. The row-wise counterpart of
   * `ignoreColumns`, for key-value blocks where a per-run value such as
   * "Generated At" is a row rather than a column. Composite keys are matched
   * in their joined form.
   */
  ignoreRows?: string[];
  /**
   * Cell addresses on this sheet holding report metadata. Derived from the
   * workbook-level `metadata` list, not written by hand -- see WorkbookSpec.
   */
  metadataCells?: string[];
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
export interface TableRequest {
  /** Table name within the sheet. */
  table: string;
  /** Key the resulting model is filed under. */
  key: string;
  spec: ResolvedSpec;
}

export interface WorkbookReader extends SheetReader {
  /**
   * Returns every sheet name in workbook order, plus a model per table that
   * `tablesFor` asks for, filed under the request's `key`. Sheets it returns
   * no requests for are listed but not built, so unconfigured tabs cost
   * nothing beyond being named.
   */
  readWorkbook(
    path: string,
    tablesFor: (sheet: string) => TableRequest[],
  ): Promise<{ sheets: string[]; models: Map<string, SheetModel> }>;
}

export type ResolvedSpec = Required<
  Omit<SheetSpec, 'tolerance' | 'relativeTolerance' | 'invariants' | 'sheet' | 'csv'>
> & {
  sheet: string | number;
  tolerance: Record<string, number>;
  relativeTolerance: Record<string, number>;
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
  /**
   * This sheet is not in every report of this type, so its absence is normal.
   *
   * Configuring a sheet that neither file contains is otherwise an error, and
   * deliberately so: it is almost always a misspelled name, and a misspelled
   * name means the sheet you thought you had configured is quietly being
   * compared without your settings. That check cannot tell a typo from a sheet
   * that simply is not in this month's report, so say which you meant. Then the
   * settings can live once at the report-type level and apply to the cases that
   * do have the sheet.
   */
  optional?: boolean;
  /**
   * Sheets holding more than one table -- an "output info" block above the
   * data is the common case. Each entry gets its own headers, key and
   * tolerances. Tables are bounded by the next one's `headerRow`, so no row
   * counting is needed as the data grows.
   */
  tables?: Record<string, TableSpec>;
};

/** One table within a sheet. Same options as a sheet, minus further nesting. */
export type TableSpec = Omit<WorkbookSheetSpec, 'tables'>;

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
   * Report metadata: cells that identify the run rather than describe it --
   * report name, report id, creation timestamp. The test is narrow: the value
   * has to differ between any two runs *by construction*, the way a freshly
   * minted id does. Comparing those reports a difference every time, and a
   * finding that is always present is one nobody reads.
   *
   * Something that merely sounds like metadata does not qualify. A creator
   * name is stable when the same account generates every report, so a change
   * there means the wrong account ran it -- a finding, not noise.
   *
   * An entry is either a label, matched against a cell's text and taking the
   * value beside it (`"Report ID"` covers `A1 "Report ID"` and `B1 4542`, and
   * also a fused `="Report ID: " & id`), or a cell reference, `"Cover!A2"` for
   * a bare date with no label of its own, or `"A2"` for the same cell on every
   * sheet. Either form takes a sheet qualifier -- `"Report Info!Report ID"` --
   * for a word that means run identity in a header block and a column heading
   * somewhere else.
   *
   * Matched cells are read, listed in the report with both values, and left
   * out of the verdict. Nothing downstream of them is chased either: a caption
   * that reads the report name is metadata too.
   *
   * Do not put anything the figures depend on here either -- view of risk,
   * currency, model version, the as-at date of the data. If one of those
   * moves, the numbers under it should have moved too, and that is worth
   * being told.
   */
  metadata?: string[];
  /**
   * Treat added and unconfigured sheets as failures rather than review items.
   * Removed sheets are always failures. Default false.
   */
  strictSheets?: boolean;
  /**
   * Compare a table that has no row key by matching rows on position instead
   * of leaving it unchecked. Default true.
   *
   * Detection never invents a key, and real reports are full of tables that
   * have none -- a geography breakdown whose subtotal rows are blank in every
   * identifying column, say. Left alone those tables are simply not compared,
   * which on these reports meant most of the workbook going unchecked.
   *
   * Position is exact whenever both sides hold the same rows in the same
   * order, and reads one inserted row as many changed ones when they do not.
   * Every table matched this way is named in the report, so the weaker
   * guarantee is never invisible. Set false to go back to not comparing them.
   */
  matchUnkeyedRowsByPosition?: boolean;

  /**
   * Which cases beneath this file to run, as paths relative to the folder the
   * file sits in. Without it, everything runs.
   *
   *   { "cases": ["comparison_report/*", "!comparison_report/case_002"] }
   *
   * `*` stands for a run of characters within one path segment and `**` for
   * any number of segments; a leading `!` excludes. Naming a folder selects
   * everything under it, so `"comparison_report"` and `"comparison_report/**"`
   * mean the same thing.
   *
   * Written in a meta.json, it says which cases that folder is for, so a run
   * of the whole tree does what the tree says rather than what the command
   * line remembers to ask for. Every file that carries one narrows further:
   * a case has to be selected by all of them to run.
   *
   * A case left out is not a case that passed. The count of what was set aside
   * is printed with the results, because a selection nobody notices is how a
   * report quietly stops being checked.
   */
  cases?: string[];

  /**
   * What each sheet should hold, checked after the comparison and failed when
   * it does not hold. Keyed by worksheet name; either the number of tables
   * layer 1 should have compared there, or the A1 ranges it should have
   * covered.
   *
   *   { "expect": { "Report Info": ["A2:B17", "H2:J22", "A19:B24"], "Cover": 1 } }
   *
   * This is the counterpart to detection being a guess. Detection is re-made
   * from the files on every run, which is what lets a report change shape
   * without breaking the config -- and it means a table can stop being
   * compared without anything failing. The only signal is a smaller number in
   * a summary line nobody was watching. On one report here an entire block
   * went unread by layer 1 for the life of the tool, and it took someone
   * opening the spreadsheet and counting to find out.
   *
   * Written once from the "What was verified" section of a report, it turns
   * that into a failure with a name.
   *
   * It is an assertion, never an instruction: it changes nothing about what is
   * compared, so an entry that is wrong stops the run and says so rather than
   * quietly comparing the wrong thing -- which is exactly what pinning a
   * `headerRow` does when a report shifts. Ranges belong in a case.json, since
   * they describe two particular files; a count often holds for a whole report
   * type and can go in its meta.json.
   */
  expect?: Record<string, number | string[]>;

  /* --- labelling: says what a case is, changes nothing about the run ------ */

  /**
   * What this case is for, in the words of whoever wrote it: "a peril column
   * added between two others". Titles the report and heads the case's block in
   * the run log, so a failure names the scenario rather than a folder.
   *
   * Read from the case's own `case.json` and nowhere else. A label inherited
   * from a folder above would describe every case beneath it identically,
   * which is worse than having none -- the log would repeat one sentence
   * thirty-nine times and say nothing.
   */
  label?: string;
  /**
   * The kind of report these cases compare: "Global Standard Cat Report".
   * Inherited like any other setting, so it is written once beside the report
   * type it names, and every case below carries it into the log and report.
   */
  reportType?: string;
  /**
   * The same thing, under the word a generator's own domain uses. A Conditional
   * EP is an analysis and a Data Transmittal is an entity, and a tree naming
   * them that way should not have every one of those folders filed under
   * "Unspecified report type". `reportType` wins where more than one is set.
   */
  analysisType?: string;
  entityType?: string;
  /**
   * Where the two files came from -- ids, download timestamps, hashes. Carried
   * for provenance and never read by the comparison, so its shape is whatever
   * the tool that fetched the pair chose to record.
   */
  source?: unknown;
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
  /** Table within the sheet. Equal to `sheet` when the sheet holds one table. */
  table: string;
  /** How the outcome is named in reports: "Policies", or "Policies · Info". */
  label: string;
  status: SheetStatus;
  /** Set only when status is 'compared'. */
  diff?: DiffResult;
  /**
   * The rectangle layer 1 read, per file, as A1 ranges: "H2:J22". Set only
   * when status is 'compared'.
   *
   * What was covered is a different question from what was configured, and the
   * one people actually ask: a spec with no `endRow` runs to the bottom of the
   * sheet and says nothing about where the table stopped. The two sides are
   * kept apart because they can differ, and a table that grew or moved between
   * runs is worth seeing as such.
   */
  range?: { base: string; next: string };
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
