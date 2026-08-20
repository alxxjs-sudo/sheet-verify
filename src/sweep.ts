import { extname } from 'node:path';
import { readFile } from 'node:fs/promises';
import { parse } from 'csv-parse/sync';
import ExcelJS from 'exceljs';
import { columnRange, numToCol } from './a1.js';
import { addressOf, impactOf, type FormulaCell } from './impact.js';
import { canonHeader, toleranceFor } from './model.js';
import { metadataOn, parseMetadata, type MetadataRules } from './metadata.js';
import { openWorkbook } from './open-xlsx.js';
import { CSV_SHEET } from './reader-csv.js';
import type { ComparedTable } from './workbook.js';

/**
 * The second layer: every cell of both files, compared by address.
 *
 * Layer 1 is selective. It aligns columns by header name and rows by business
 * key, which is what lets it survive schema drift -- and it means it only
 * compares tables it could find a key for. Tables it could not key are reported
 * as "not compared", and nothing today says whether anything was hiding in
 * them.
 *
 * This layer answers that. Being positional it needs no keys, no headers and no
 * configuration, so it reaches every cell in the file. Its headline number is
 * not how many cells differ -- it is how many differ *among the cells layer 1
 * never compared*. Zero means layer 1's coverage gaps hid nothing and its
 * verdict stands for the whole file. Anything else points straight at what
 * needs configuring.
 *
 * The trade is that a positional comparison lights up when the layout moves:
 * one inserted column shifts every cell to its right. That is exactly the noise
 * layer 1 exists to avoid, which is why this layer only ever reports and never
 * decides the outcome -- layer 1 has already said "a column moved", so the
 * noise arrives explained. Sheets whose extent changed are flagged `reshaped`
 * for the same reason.
 *
 * Note what a cell is compared *as*: a literal contributes its value, a formula
 * cell contributes its formula text *and* its cached result if it has one. No
 * formula is evaluated, and none needs to be -- a formula plus its inputs
 * determines its result, so a changed input shows up where the input lives,
 * which is the root cause rather than the cascade.
 *
 * The cached result is compared only when both sides have one. A file that has
 * never been opened in Excel carries formulas with no results at all, so
 * comparing a never-opened baseline against a report someone opened once would
 * otherwise report every formula in the file as a value change. Absent is
 * unknown, not different.
 */

/**
 * What layer 1 did about a cell. Only `gap` counts against it: a cell in a
 * column layer 1 flagged as added is not a blind spot, it is a finding written
 * up one level higher, and an excluded column was excluded on purpose.
 * Collapsing those three into "uncompared" inflates the headline with things
 * the user was already told.
 */
export type SweepStatus =
  /** Layer 1 held the two cells up against each other. */
  | 'compared'
  /** Not compared, but its column was reported as added or removed. */
  | 'reported'
  /** Excluded on purpose, by ignoreColumns or ignoreRows. */
  | 'excluded'
  /** Run identity -- report name, id, creator, generated-on. See metadata.ts. */
  | 'metadata'
  /** Nobody looked. */
  | 'gap';

/** Where a gap sits, so it can be closed. Set only on `gap` cells. */
export type SweepReason =
  /** In a compared table's rows, but in a column layer 1 never paired. */
  | 'column not compared'
  /** In a compared table's span, but the row was not matched. */
  | 'row not compared'
  /** Between or above tables: titles, notes, blank gutters. */
  | 'outside any compared table'
  /** The sheet itself was never compared -- ignored, or nothing keyable on it. */
  | 'sheet not compared';

export interface SweepCell {
  sheet: string;
  address: string;
  row: number;
  column: number;
  /** Token from the baseline: a literal value, or "=" plus formula text. */
  base: string;
  next: string;
  status: SweepStatus;
  /** Set only when `status` is 'gap'. */
  reason?: SweepReason;
  /** Set only when `status` is 'metadata': the pattern that excluded it. */
  rule?: string;
  /**
   * Set only on a tolerated cell: the tolerance it was judged against. Kept per
   * cell because it is resolved per column, so one run can forgive a hundredth
   * in one place and nothing at all in another.
   */
  tolerance?: number;
  /**
   * The compared table the cell falls in, named as reports name it. Absent for
   * a cell outside every table layer 1 compared.
   */
  table?: string;
}

export interface SweptSheet {
  sheet: string;
  status: 'swept' | 'added' | 'removed';
  /** Cells holding something on either side. */
  cells: number;
  compared: number;
  differing: number;
  /** Differing cells nobody looked at. */
  gaps: number;
  /** Numeric cells whose gap is inside the tolerance set for their column. */
  tolerated: number;
  /** The two sides differ in extent, so positional drift is expected here. */
  reshaped: boolean;
}

/** A formula cell that reads something which differs. See `impact.ts`. */
export interface AffectedCell {
  sheet: string;
  address: string;
  /**
   * Header of the column it sits in, when it falls inside a table layer 1
   * compared. Absent for a formula in a title block or an unkeyed area, where
   * there is no header to name it by.
   */
  column?: string;
  /** True when it reads a differing cell only through another formula. */
  indirect: boolean;
  /**
   * The cell whose change reaches this one. Always named; its two values are
   * carried too when that cell is itself a difference.
   *
   * This is what makes the finding answerable without opening both files. A
   * reader who has already spotted a difference Excel shows and the ledger does
   * not needs to be told which change drives it, not merely that something
   * does.
   */
  via: {
    sheet: string;
    address: string;
    /**
     * What the driving cell holds on each side. Present only when that cell is
     * itself a difference; a cell reached through another formula has no
     * stored value to quote, and inventing one would be worse than a blank.
     */
    base?: string;
    next?: string;
  };
}

export interface SweepResult {
  base: string;
  next: string;
  sheets: SweptSheet[];
  cellsSwept: number;
  cellsCompared: number;
  /**
   * Formula cells that will recalculate to a different number, because they
   * read something that differs. Nothing is evaluated -- these are reached by
   * following references outward from the cells that changed.
   */
  affected: AffectedCell[];
  totalAffected: number;
  /** Differing cells that no formula reads, so nothing downstream moves. */
  inertChanges: number;
  /** Differing cells, capped by `limit`. The counts below are exact. */
  differences: SweepCell[];
  totalDifferences: number;
  /** The headline number: differing cells nobody looked at. */
  totalGaps: number;
  /** Differing cells whose column layer 1 reported as added or removed. */
  totalReported: number;
  /** Differing cells in a deliberately excluded column or row. */
  totalExcluded: number;
  /**
   * Every cell recognised as run identity, differing or not. Listed in full:
   * skipping a comparison is only defensible if the reader can see what was
   * skipped.
   */
  metadata: SweepCell[];
  /** Metadata cells whose two sides differ -- expected, and not a finding. */
  metadataDiffering: number;
  /**
   * Cells whose two numbers differ by less than the tolerance set for their
   * column. Kept out of `totalDifferences`, and listed in full: a tolerance is
   * a statement that a gap that size does not matter, not a reason to stop
   * showing it.
   */
  tolerated: SweepCell[];
  totalTolerated: number;
}

export interface SweepOptions {
  /**
   * Differing cells recorded. Unlimited by default: the report lists every
   * one, and a caller that wants a preview can ask for fewer.
   */
  limit?: number;
  /**
   * Sheets excluded from the comparison on purpose. Their cells are still
   * swept -- it is worth knowing what changed on them -- but reported as
   * excluded rather than as something nobody looked at. Being told a gap
   * exists on a sheet you deliberately excluded is a false alarm, and false
   * alarms are what make the headline number stop being read.
   */
  ignoreSheets?: string[];
  /**
   * Patterns naming report metadata -- see metadata.ts. Cells they match are
   * read, listed, and left out of the verdict.
   */
  metadata?: string[];
  /**
   * Tolerance for cells this sweep reaches but layer 1 did not: title blocks,
   * unkeyed tables, anywhere outside a compared table's rows. Per-column
   * tolerances come from the compared tables themselves, so this is the `*`
   * fallback and nothing else.
   *
   * Without it a tolerance would only ever quiet layer 1, and the headline
   * count -- which comes from here -- would go on reporting float noise the
   * reader has already said they do not care about.
   */
  tolerance?: number;
}

const DEFAULT_LIMIT = Number.MAX_SAFE_INTEGER;

/** Excel forbids two sheets whose names differ only in case. */
const canon = (s: string): string => s.trim().toLowerCase();

const addr = (row: number, col: number): string => `${numToCol(col)}${row}`;

/** One cell reduced to a comparable string, at full stored precision. */
function literal(v: unknown): string {
  if (v === null || v === undefined) return '';
  // An Invalid Date is an empty cell wearing a date format; see reader-excel.
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? '' : v.toISOString();
  // Full precision: rounding here would make this layer blind to exactly the
  // small differences layer 1 now reports, and the two layers must agree.
  if (typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'object') {
    const o = v as Record<string, any>;
    if ('richText' in o) return (o.richText as any[]).map((t) => t.text).join('').trim();
    if ('hyperlink' in o) return String(o.text ?? o.hyperlink ?? '').trim();
    if ('error' in o) return String(o.error);
    if ('result' in o) return literal(o.result);
    return '';
  }
  return String(v).trim();
}

/**
 * A formula cell's token: marker + formula + separator + cached result.
 * Encoded into one string rather than an object because a large sheet holds six
 * figures of these. The markers are written as escapes rather than literal
 * control characters, which would make this file read as binary to grep and
 * diff; no cell text can hold one, so they cannot collide with a literal.
 */
const FORMULA_MARK = '\u0002';
const RESULT_MARK = '\u0001';

const formulaToken = (formula: string, result: string) =>
  `${FORMULA_MARK}${formula}${RESULT_MARK}${result}`;

const isFormula = (token: string) => token.startsWith(FORMULA_MARK);

/** Splits a formula token back into its text and its cached result. */
function parts(token: string): { formula: string; result: string } {
  const at = token.indexOf(RESULT_MARK);
  return { formula: token.slice(1, at), result: token.slice(at + 1) };
}

/**
 * Whether two cells say the same thing. Two formula cells agree when their text
 * agrees and neither carries a result that contradicts the other -- a result
 * only one side has says nothing about the other.
 */
function same(a: string, b: string): boolean {
  if (a === b) return true;
  if (!isFormula(a) || !isFormula(b)) return false;

  const x = parts(a);
  const y = parts(b);
  if (x.formula !== y.formula) return false;
  return x.result === '' || y.result === '';
}

/**
 * What a cell says, as a person reads it. A formula contributes its cached
 * result when it has one, since that is the text on screen; with no result
 * stored, its own text is the best available reading of it.
 */
function textOf(token: string): string {
  if (!isFormula(token)) return token;
  const { formula, result } = parts(token);
  return result === '' ? formula : result;
}

/** How a token reads in a report. */
function display(token: string): string {
  if (!isFormula(token)) return token;
  const { formula, result } = parts(token);
  return result === '' ? `=${formula}` : `=${formula} → ${result}`;
}

/** A sheet flattened to address -> token. Blank cells are simply absent. */
interface Grid {
  cells: Map<string, string>;
  rows: number;
  columns: number;
}

function gridOf(ws: ExcelJS.Worksheet): Grid {
  const cells = new Map<string, string>();

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      // A merged range carries its value once, on the master. Reading the
      // slaves too would repeat that token across the range and turn one
      // changed heading into a dozen differences.
      if (cell.isMerged && cell.master !== cell) return;

      // IMPORTANT: the `formula` getter resolves shared (filled-down) formulas
      // back to translated text; cell.value.formula gives only the master's.
      const formula = (cell.formula as string | undefined) || null;
      if (formula) {
        // `result` is absent on a file Excel has never recalculated. That is
        // recorded as an empty result, not as an empty value, so `same()` can
        // tell "no result stored" from "the result is blank".
        const v = cell.value as { result?: unknown } | null;
        const result = v && typeof v === 'object' && 'result' in v ? literal(v.result) : '';
        cells.set(addr(rowNumber, colNumber), formulaToken(formula.trim(), result));
        return;
      }
      const token = literal(cell.value);
      if (token !== '') cells.set(addr(rowNumber, colNumber), token);
    });
  });

  return { cells, rows: ws.rowCount, columns: ws.columnCount };
}

/** Every sheet of a file, flattened. CSV presents as its single pseudo-sheet. */
export async function sweepFile(path: string): Promise<Map<string, Grid>> {
  const out = new Map<string, Grid>();

  if (['.csv', '.tsv', '.txt'].includes(extname(path).toLowerCase())) {
    const text = (await readFile(path)).toString('utf8').replace(/^﻿/, '');
    const table: string[][] = parse(text, {
      relax_column_count: true,
      skip_empty_lines: false,
      bom: true,
      trim: true,
    });
    const cells = new Map<string, string>();
    let columns = 0;
    table.forEach((row, r) => {
      columns = Math.max(columns, row.length);
      row.forEach((v, c) => {
        if (v !== '') cells.set(addr(r + 1, c + 1), v);
      });
    });
    out.set(CSV_SHEET, { cells, rows: table.length, columns });
    return out;
  }

  const wb = await openWorkbook(path);
  for (const ws of wb.worksheets) out.set(ws.name, gridOf(ws));
  return out;
}

/**
 * What layer 1 actually looked at, per sheet. Addresses are collected from both
 * sides because a moved column means the same logical cell sits at different
 * addresses in each file; a swept cell counts as covered only when layer 1
 * examined that address in both.
 */
interface Coverage {
  base: Set<string>;
  next: Set<string>;
  /** Row numbers layer 1 compared, for explaining a miss. */
  rows: Set<number>;
  columns: Set<number>;
  /** Columns layer 1 reported as added or removed. */
  reported: Set<number>;
  /** Columns and rows excluded on purpose. */
  excludedColumns: Set<number>;
  excludedRows: Set<number>;
  /** Row spans of the tables layer 1 compared. */
  spans: { from: number; to: number }[];
}

/**
 * Column names by position, per sheet, so a cell found positionally can be
 * reported by the header a reader recognises.
 *
 * Layer 2 knows only addresses -- that is what lets it reach everywhere -- so
 * "Geocoding!E15" is all it can say on its own. A sheet holds several tables
 * with different headers, hence the row span on each entry: the same column
 * number means different things above and below a table boundary.
 */
interface HeaderBand { from: number; to: number; byCol: Map<number, string> }

function headerBandsOf(compared: ComparedTable[]): Map<string, HeaderBand[]> {
  const out = new Map<string, HeaderBand[]>();
  for (const t of compared) {
    const byCol = new Map<number, string>();
    for (const model of [t.base, t.next]) {
      for (const [name, col] of model.headerIndex) if (!byCol.has(col)) byCol.set(col, name);
    }
    if (!byCol.size) continue;
    const key = canon(t.sheet);
    const bands = out.get(key) ?? [];
    bands.push({
      from: Math.max(t.spec.headerRow, 1),
      to: t.spec.endRow > 0 ? t.spec.endRow : Number.MAX_SAFE_INTEGER,
      byCol,
    });
    out.set(key, bands);
  }
  return out;
}

const columnAt = (
  bands: Map<string, HeaderBand[]>,
  sheet: string,
  row: number,
  col: number,
): string | undefined =>
  bands.get(canon(sheet))?.find((b) => row >= b.from && row <= b.to)?.byCol.get(col);

/** A compared table's tolerances, laid out by the cells they cover. */
interface ToleranceBand {
  /** How the table is named in reports: "Portfolio Totals · Table 3". */
  label: string;
  from: number;
  to: number;
  /** Column bounds, for a sheet whose tables sit side by side. */
  fromCol: number;
  toCol: number;
  /** Per column number, resolved from the column's own name. */
  byCol: Map<number, number>;
  /** The table's `*` entry, for a column inside its rows but not in its headers. */
  star: number;
}

/**
 * Where each tolerance applies, read off the tables layer 1 compared. The
 * sweep works in addresses and a tolerance is written against a column name,
 * so the translation has to happen somewhere; doing it here keeps the two
 * layers agreeing about which gaps matter.
 */
function toleranceBandsOf(compared: ComparedTable[]): Map<string, ToleranceBand[]> {
  const out = new Map<string, ToleranceBand[]>();
  for (const t of compared) {
    const star = t.spec.tolerance['*'] ?? 0;
    const byCol = new Map<number, number>();
    for (const model of [t.base, t.next]) {
      for (const [name, col] of model.headerIndex) {
        if (byCol.has(col)) continue;
        byCol.set(col, toleranceFor(t.spec, name));
      }
    }
    if (!byCol.size && star === 0) continue;
    const key = canon(t.sheet);
    const bands = out.get(key) ?? [];
    // A table bounded left and right shares its rows with the table beside it,
    // so rows alone no longer identify which one a cell belongs to.
    const cols = columnRange(t.spec.columns, 0);
    bands.push({
      label: t.label,
      from: Math.max(t.spec.headerRow, 1),
      to: t.spec.endRow > 0 ? t.spec.endRow : Number.MAX_SAFE_INTEGER,
      fromCol: t.spec.columns ? cols.from : 1,
      toCol: t.spec.columns ? cols.to : Number.MAX_SAFE_INTEGER,
      byCol,
      star,
    });
    out.set(key, bands);
  }
  return out;
}

const bandAt = (
  bands: Map<string, ToleranceBand[]>,
  sheet: string,
  row: number,
  column: number,
): ToleranceBand | undefined =>
  bands.get(canon(sheet))?.find(
    (b) => row >= b.from && row <= b.to && column >= b.fromCol && column <= b.toCol,
  );

const toleranceIn = (band: ToleranceBand | undefined, fallback: number, col: number): number =>
  band ? band.byCol.get(col) ?? band.star : fallback;

/**
 * Whether two cells hold numbers that differ by no more than `tol`.
 *
 * This layer works in the text a cell displays -- it never sees the cell's
 * type -- so "a number" here means text that reads as one on both sides. A
 * version written "4.20" would qualify, which is the argument for setting a
 * tolerance to the size of the rounding it is meant to absorb rather than to
 * whatever round figure feels comfortable. Layer 1 has the types and does not
 * share this weakness; the two agree on every cell layer 1 reached.
 *
 * A formula whose text changed is a change whatever its result says, so those
 * are never tolerated.
 */
function withinTolerance(a: string, b: string, tol: number): boolean {
  if (tol <= 0) return false;
  if (isFormula(a) !== isFormula(b)) return false;
  if (isFormula(a) && isFormula(b) && parts(a).formula !== parts(b).formula) return false;

  const x = Number(textOf(a));
  const y = Number(textOf(b));
  if (!Number.isFinite(x) || !Number.isFinite(y)) return false;
  if (textOf(a).trim() === '' || textOf(b).trim() === '') return false;
  return Math.abs(x - y) <= tol;
}

function coverageOf(compared: ComparedTable[]): Map<string, Coverage> {
  const out = new Map<string, Coverage>();

  for (const table of compared) {
    const key = canon(table.sheet);
    let cov = out.get(key);
    if (!cov) {
      cov = {
        base: new Set(), next: new Set(), rows: new Set(), columns: new Set(),
        reported: new Set(), excludedColumns: new Set(), excludedRows: new Set(), spans: [],
      };
      out.set(key, cov);
    }

    // A column layer 1 called added or removed had no counterpart to compare
    // against, so its cells were never checked -- but the user was told about
    // the column. Same for one ignoreColumns names.
    const loose = table.spec.looseHeaders;
    const ignored = new Set(table.spec.ignoreColumns.map((c) => canonHeader(c, loose)));
    for (const [model, names] of [
      [table.next, table.diff.schema.added],
      [table.base, table.diff.schema.removed],
    ] as const) {
      for (const name of names) {
        const colNum = model.headerIndex.get(name);
        if (colNum !== undefined) cov.reported.add(colNum);
      }
    }
    for (const model of [table.base, table.next]) {
      for (const header of model.headers) {
        if (!ignored.has(canonHeader(header, loose))) continue;
        const colNum = model.headerIndex.get(header);
        if (colNum !== undefined) cov.excludedColumns.add(colNum);
      }
      for (const rowKey of table.spec.ignoreRows) {
        const row = model.rows.get(rowKey);
        const cell = row && Object.values(row)[0];
        if (cell) cov.excludedRows.add(rowNumberOf(cell.address));
      }
    }

    const columns = table.diff.schema.compared;
    // The header row is compared in substance -- a renamed column surfaces as
    // an added/removed pair -- so counting it as covered keeps a rename from
    // also arriving here as an unexplained gap.
    // A table with no header row -- `headerRow` 0, its data starting at the top
    // of the sheet -- has no such row to mark, and no row 0 exists to hold the
    // marks. Recording them anyway would put addresses in the covered set that
    // no sweep can ever visit.
    const header = table.spec.headerRow > 0 ? table.spec.headerRow : null;
    cov.spans.push({
      from: header ?? 1,
      to: table.spec.endRow > 0 ? table.spec.endRow : Number.MAX_SAFE_INTEGER,
    });
    if (header) cov.rows.add(header);

    for (const side of ['base', 'next'] as const) {
      const model = table[side];
      for (const column of columns) {
        const colNum = model.headerIndex.get(column);
        if (colNum !== undefined) {
          cov.columns.add(colNum);
          if (header) cov[side].add(addr(header, colNum));
        }
      }
    }

    // A row counts as compared when it was matched on both sides -- an added
    // or removed row is layer 1's finding, not a coverage gap, but its cells
    // were never held up against anything.
    for (const rowKey of table.base.order) {
      const baseRow = table.base.rows.get(rowKey);
      const nextRow = table.next.rows.get(rowKey);
      if (!baseRow || !nextRow) continue;

      for (const column of columns) {
        const b = baseRow[column];
        const n = nextRow[column];
        if (b) {
          cov.base.add(b.address);
          cov.rows.add(rowNumberOf(b.address));
        }
        if (n) {
          cov.next.add(n.address);
          cov.rows.add(rowNumberOf(n.address));
        }
      }
    }
  }

  return out;
}

const rowNumberOf = (address: string): number => Number(address.replace(/^[A-Z$]+/i, ''));

/** What layer 1 did about one cell, and where the gap is when it did nothing. */
function classify(
  cov: Coverage | undefined,
  address: string,
  row: number,
  column: number,
  sheetExcluded = false,
  metadataRule?: string,
): { status: SweepStatus; reason?: SweepReason; rule?: string } {
  // Sheet exclusion outranks metadata. Naming a tab in `ignoreSheets` is a
  // decision about the whole tab, and lifting two of its cells back out into
  // the metadata list would contradict it -- a column heading that happens to
  // read "Program name" on a sheet nobody looks at is not the report's name.
  if (sheetExcluded) return { status: 'excluded' };
  if (metadataRule !== undefined) return { status: 'metadata', rule: metadataRule };
  if (!cov) return { status: 'gap', reason: 'sheet not compared' };
  if (cov.base.has(address) && cov.next.has(address)) return { status: 'compared' };
  if (cov.excludedColumns.has(column) || cov.excludedRows.has(row)) return { status: 'excluded' };
  if (cov.reported.has(column)) return { status: 'reported' };

  if (!cov.spans.some((s) => row >= s.from && row <= s.to)) {
    return { status: 'gap', reason: 'outside any compared table' };
  }
  return { status: 'gap', reason: cov.rows.has(row) ? 'column not compared' : 'row not compared' };
}

/**
 * Sweeps both files and marks each differing cell as one layer 1 compared or
 * one it never reached. `compared` comes from the layer 1 run, so no file is
 * re-read to work out coverage.
 */
export async function sweep(
  basePath: string,
  nextPath: string,
  compared: ComparedTable[],
  options: SweepOptions = {},
): Promise<SweepResult> {
  const limit = options.limit ?? DEFAULT_LIMIT;
  const excluded = new Set((options.ignoreSheets ?? []).map(canon));
  const rules: MetadataRules = parseMetadata(options.metadata);
  const [baseGrids, nextGrids] = await Promise.all([sweepFile(basePath), sweepFile(nextPath)]);
  const coverage = coverageOf(compared);
  const tolerances = toleranceBandsOf(compared);
  const fallbackTolerance = options.tolerance ?? 0;

  const baseByCanon = new Map([...baseGrids.keys()].map((s) => [canon(s), s]));
  const nextByCanon = new Map([...nextGrids.keys()].map((s) => [canon(s), s]));

  const sheets: SweptSheet[] = [];
  const differences: SweepCell[] = [];
  const metadata: SweepCell[] = [];
  const toleratedCells: SweepCell[] = [];
  let metadataDiffering = 0;
  let totalTolerated = 0;
  let cellsSwept = 0;
  let cellsCompared = 0;
  let totalDifferences = 0;
  let totalGaps = 0;
  let totalReported = 0;
  let totalExcluded = 0;

  // Baseline order, then sheets the new file added -- the same ordering layer 1
  // reports in, so the two files read side by side.
  const names = [
    ...baseGrids.keys(),
    ...[...nextGrids.keys()].filter((s) => !baseByCanon.has(canon(s))),
  ];

  for (const sheet of names) {
    const base = baseGrids.get(sheet);
    const nextName = nextByCanon.get(canon(sheet));
    const next = nextName ? nextGrids.get(nextName) : undefined;

    // Nothing to compare a one-sided sheet against. Layer 1 already reports it
    // as added or removed; repeating every one of its cells as a difference
    // would bury the cells that genuinely went unchecked.
    if (!base || !next) {
      const only = base ?? next!;
      sheets.push({
        sheet,
        status: base ? 'removed' : 'added',
        cells: only.cells.size,
        compared: 0,
        differing: 0,
        gaps: 0,
        tolerated: 0,
        reshaped: false,
      });
      continue;
    }

    const cov = coverage.get(canon(sheet));
    const sheetExcluded = excluded.has(canon(sheet));
    const meta = metadataOn(sheet, [base.cells, next.cells], rules, textOf);
    const addresses = new Set([...base.cells.keys(), ...next.cells.keys()]);
    let differing = 0;
    let gaps = 0;
    let compared = 0;
    let tolerated = 0;

    for (const address of addresses) {
      const b = base.cells.get(address) ?? '';
      const n = next.cells.get(address) ?? '';
      const row = rowNumberOf(address);
      const column = colNumberOf(address);
      const { status, reason, rule } = classify(
        cov, address, row, column, sheetExcluded, meta.get(address),
      );
      if (status === 'compared') compared++;

      // Display form, so a consumer never sees the internal markers.
      const record = (): SweepCell => ({
        sheet, address, row, column, base: display(b), next: display(n), status,
        ...(reason ? { reason } : {}),
        ...(rule !== undefined ? { rule } : {}),
      });

      // Metadata is listed whether it differs or not: the point of the list is
      // to show what was skipped, and a reader cannot see that from silence.
      if (status === 'metadata') {
        metadata.push(record());
        if (!same(b, n)) metadataDiffering++;
        continue;
      }

      if (same(b, n)) continue;

      // A gap the reader has already said does not matter. Counted and listed
      // on its own rather than folded into the differences, so a tolerance
      // quiets the report without quietly editing what the run found.
      const band = bandAt(tolerances, sheet, row, column);
      const tol = toleranceIn(band, fallbackTolerance, column);
      if (withinTolerance(b, n, tol)) {
        tolerated++;
        if (toleratedCells.length < limit) {
          toleratedCells.push({
            ...record(),
            tolerance: tol,
            // Named so the report can say which table absorbed the drift.
            ...(band ? { table: band.label } : {}),
          });
        }
        continue;
      }

      differing++;
      if (status === 'gap') gaps++;
      else if (status === 'reported') totalReported++;
      else if (status === 'excluded') totalExcluded++;

      if (differences.length < limit) differences.push(record());
    }

    cellsSwept += addresses.size;
    cellsCompared += compared;
    totalDifferences += differing;
    totalGaps += gaps;
    totalTolerated += tolerated;

    sheets.push({
      sheet,
      status: 'swept',
      cells: addresses.size,
      compared,
      differing,
      gaps,
      tolerated,
      reshaped: base.rows !== next.rows || base.columns !== next.columns,
    });
  }

  // What the differences will do to the rest of the workbook once Excel
  // recalculates. The formulas come from the new report, since that is the
  // file someone will open; the seeds are every cell that differs.
  const formulas: FormulaCell[] = [];
  for (const [sheet, grid] of nextGrids) {
    for (const [address, token] of grid.cells) {
      if (!isFormula(token)) continue;
      formulas.push({
        sheet,
        row: rowNumberOf(address),
        col: colNumberOf(address),
        formula: parts(token).formula,
      });
    }
  }
  const impact = impactOf(
    formulas,
    differences.map((d) => ({ sheet: d.sheet, row: d.row, col: d.column })),
  );
  const bands = headerBandsOf(compared);
  // Every differing cell by address, so an affected cell can quote what its
  // cause holds on each side rather than only naming it.
  const differing = new Map(
    differences.map((d) => [`${canon(d.sheet)}|${d.address}`, d] as const),
  );
  const affected: AffectedCell[] = impact.affected.map((c) => {
    const column = columnAt(bands, c.sheet, c.row, c.col);
    const cause = differing.get(`${canon(c.via.sheet)}|${addressOf(c.via)}`);
    return {
      sheet: c.sheet,
      address: addressOf(c),
      ...(column !== undefined ? { column } : {}),
      indirect: c.indirect,
      via: {
        sheet: c.via.sheet,
        address: addressOf(c.via),
        ...(cause ? { base: cause.base, next: cause.next } : {}),
      },
    };
  });

  // Gaps first: they are the reason this layer exists. Then by position, so a
  // reader can follow the file.
  const rank = { gap: 0, reported: 1, excluded: 2, metadata: 3, compared: 4 };
  differences.sort(
    (a, b) =>
      rank[a.status] - rank[b.status] ||
      a.sheet.localeCompare(b.sheet) ||
      a.row - b.row ||
      a.column - b.column,
  );

  return {
    base: basePath,
    next: nextPath,
    sheets,
    cellsSwept,
    cellsCompared,
    differences,
    totalDifferences,
    totalGaps,
    totalReported,
    totalExcluded,
    metadata: metadata.sort(
      (a, b) => a.sheet.localeCompare(b.sheet) || a.row - b.row || a.column - b.column,
    ),
    metadataDiffering,
    tolerated: toleratedCells,
    totalTolerated,
    affected: affected.slice(0, limit),
    totalAffected: impact.affected.length,
    inertChanges: impact.inert,
  };
}

const colNumberOf = (address: string): number => {
  const letters = address.replace(/[^A-Z]/gi, '').toUpperCase();
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
};

/** Shortens a token so a wide formula does not wrap the report. */
const clip = (s: string, width = 44): string =>
  s.length <= width ? s : `${s.slice(0, width - 1)}…`;

const show = (s: string): string => (s === '' ? '∅' : clip(s));

/**
 * The coverage report. Written so the first thing read is the one number that
 * matters: whether anything differed where layer 1 was not looking.
 */
export function formatSweepReport(s: SweepResult, options: SweepOptions = {}): string {
  const limit = options.limit ?? DEFAULT_LIMIT;
  const out: string[] = [];
  const h = (t: string) => out.push('', t, '─'.repeat(Math.min(t.length, 72)));
  const pct = (n: number, of: number) => (of === 0 ? '0%' : `${Math.round((n / of) * 100)}%`);

  const swept = s.sheets.filter((x) => x.status === 'swept');
  out.push(
    `baseline  ${s.base}`,
    `actual    ${s.next}`,
    `swept     ${s.cellsSwept} cell(s) across ${swept.length} sheet(s) present in both`,
    `compared  ${s.cellsCompared} of those checked by layer 1 (${pct(s.cellsCompared, s.cellsSwept)})`,
  );
  if (s.metadata.length) {
    out.push(
      `metadata  ${s.metadata.length} read and set aside as run identity` +
      `${s.metadataDiffering ? `, ${s.metadataDiffering} differing` : ''}`,
    );
  }

  if (s.totalGaps === 0) {
    out.push(
      '',
      s.totalDifferences === 0
        ? '✓ No cell differs anywhere in either file.'
        : `✓ All ${s.totalDifferences} differing cell(s) were accounted for by layer 1 — ` +
          'nothing changed where nobody was looking.',
    );
  } else {
    h(`UNCHECKED DIFFERENCES (${s.totalGaps}) — nobody looked at these`);
    const missed = s.differences.filter((d) => d.status === 'gap');
    for (const d of missed) {
      out.push(`  ${d.sheet}!${d.address}  ${show(d.base)}  →  ${show(d.next)}   [${d.reason}]`);
    }
    if (s.totalGaps > missed.length) out.push(`  … and ${s.totalGaps - missed.length} more`);
  }

  // Everything below is a difference layer 1 has already accounted for. It is
  // listed so the totals reconcile, not because it needs acting on.
  const section = (
    status: SweepStatus,
    total: number,
    title: string,
  ) => {
    if (total <= 0) return;
    const rows = s.differences.filter((d) => d.status === status);
    h(`${title} (${total})`);
    for (const d of rows.slice(0, 20)) {
      out.push(`  ${d.sheet}!${d.address}  ${show(d.base)}  →  ${show(d.next)}`);
    }
    const shown = Math.min(rows.length, 20);
    if (total > shown) out.push(`  … and ${total - shown} more`);
  };

  if (s.metadataDiffering > 0) {
    h(`REPORT METADATA (${s.metadataDiffering}) — expected to differ, not judged`);
    for (const d of s.metadata.filter((m) => m.base !== m.next)) {
      out.push(`  ${d.sheet}!${d.address}  ${show(d.base)}  →  ${show(d.next)}   [${d.rule}]`);
    }
  }

  section('reported', s.totalReported, "IN AN ADDED OR REMOVED COLUMN — layer 1 reported the column");
  section('excluded', s.totalExcluded, 'IN AN EXCLUDED COLUMN OR ROW — skipped on purpose');
  section(
    'compared',
    s.totalDifferences - s.totalGaps - s.totalReported - s.totalExcluded,
    "ALREADY REPORTED — inside layer 1's compared set, see diff.txt",
  );

  if (s.totalAffected > 0) {
    h(`WILL RECALCULATE DIFFERENTLY (${s.totalAffected}) — downstream of the changes above`);
    out.push('  These cells hold formulas whose text has not changed, so nothing above');
    out.push('  reports them. They read a cell that did change, so Excel will produce a');
    out.push('  different number here the moment the file is opened. No value is shown');
    out.push('  because neither file stores one.');
    const bySheet = new Map<string, string[]>();
    for (const a of s.affected) {
      const list = bySheet.get(a.sheet) ?? [];
      list.push(a.address);
      bySheet.set(a.sheet, list);
    }
    for (const [sheet, cells] of bySheet) {
      const shown = cells.slice(0, 24).join(', ');
      out.push(`  ~ ${sheet}: ${shown}${cells.length > 24 ? `, … and ${cells.length - 24} more` : ''}`);
    }
    if (s.totalAffected > s.affected.length) {
      out.push(`  … and ${s.totalAffected - s.affected.length} more not listed`);
    }
  } else if (s.totalDifferences > 0) {
    h('WILL RECALCULATE DIFFERENTLY (0)');
    out.push('  No formula reads any of the cells that changed, so nothing else in the');
    out.push('  workbook moves when it is recalculated.');
  }

  const reshaped = swept.filter((x) => x.reshaped);
  if (reshaped.length) {
    h(`RESHAPED SHEETS (${reshaped.length}) — extent changed, so positions drift`);
    out.push('  A sheet that gained a row or column shifts every cell after it, and this');
    out.push('  layer compares by position. Read layer 1 first for what actually moved.');
    for (const x of reshaped) out.push(`  ~ ${x.sheet} — ${x.differing} differing, ${x.gaps} unchecked`);
  }

  const oneSided = s.sheets.filter((x) => x.status !== 'swept');
  if (oneSided.length) {
    h(`NOT SWEPT (${oneSided.length}) — present in one file only`);
    for (const x of oneSided) {
      out.push(`  ${x.status === 'added' ? '+' : '−'} ${x.sheet} — ${x.cells} cell(s), nothing to compare against`);
    }
  }

  // The forward-looking half: sheets that happen to match today but that no
  // one is checking, so a future change there would pass unnoticed.
  const blind = swept.filter((x) => x.compared === 0 && x.cells > 0).sort((a, b) => b.cells - a.cells);
  if (blind.length) {
    const cells = blind.reduce((n, x) => n + x.cells, 0);
    h(`NO COVERAGE (${blind.length} sheet(s), ${cells} cell(s)) — layer 1 checks nothing here`);
    out.push('  These match today. Nothing would report it if they stopped.');
    for (const x of blind.slice(0, 15)) {
      out.push(`  ? ${x.sheet} — ${x.cells} cell(s) swept, ${x.differing} differing`);
    }
    if (blind.length > 15) out.push(`  … and ${blind.length - 15} more`);
  }

  if (s.totalDifferences > limit) {
    out.push('', `Note: ${s.totalDifferences} cells differ; the first ${limit} are listed. Counts above are exact.`);
  }
  return out.join('\n');
}

/** One-line verdict, for logs and CLI output. */
export function summarizeSweep(s: SweepResult): string {
  // Metadata is reported separately and never counts here: a run whose only
  // movement is the report id is a clean run, and saying otherwise is what
  // teaches someone to stop reading the first line.
  const meta = s.metadataDiffering > 0 ? `, ${s.metadataDiffering} metadata` : '';
  // Same reasoning for tolerated cells: they moved, the reader has said a move
  // that size is immaterial, and reporting them as differences would make the
  // one line people read overstate what happened.
  const tol = s.totalTolerated > 0 ? `, ${s.totalTolerated} within tolerance` : '';
  if (s.totalDifferences === 0) return `every cell identical${tol}${meta}`;
  if (s.totalGaps === 0) {
    return `${s.totalDifferences} differing, all accounted for by layer 1${tol}${meta}`;
  }
  return `${s.totalGaps} differing cell(s) nobody checked (of ${s.totalDifferences})${tol}${meta}`;
}
