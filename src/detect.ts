import { extname } from 'node:path';
import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { parse } from 'csv-parse/sync';
import type { WorkbookSpec, WorkbookSheetSpec } from './types.js';
import { CSV_SHEET } from './reader-csv.js';
import { cellToken, headerName } from './reader-excel.js';
import { numToCol } from './a1.js';
import { openWorkbook } from './open-xlsx.js';

/**
 * Works out how a workbook is laid out, so a comparison needs no hand-written
 * spec. Detection is a pre-pass that emits an ordinary `WorkbookSpec`: the
 * comparison itself is unchanged, and the guess can be printed, saved and
 * edited when it gets something wrong.
 */

export interface DetectedTable {
  /** Name used in reports. The sheet's own name when it holds one table. */
  name: string;
  headerRow: number;
  /** 0 when the table runs to the last row. */
  endRow: number;
  /**
   * Columns the table occupies, as a range of letters: "H:J". Always set, and
   * only worth carrying into a spec when the sheet holds tables side by side.
   */
  columns: string;
  /** Where it sits in the grid, 0-based. Used to order and bound the tables. */
  region: Region;
  headers: string[];
  /** Null when no column or pair of columns identifies a row. */
  keyColumns: string[] | null;
  rows: number;
  /**
   * How many data rows hold something, per column, positionally aligned with
   * `headers`. A column written on a handful of rows out of hundreds is a
   * group heading rather than a column with gaps -- which is the shape
   * `fillKeyDown` exists for, and worth being able to point at rather than
   * assume.
   */
  filledPerColumn: number[];
  /** Whether the table's first data row holds something, per column. */
  firstRowFilled: boolean[];
}

export interface DetectedSheet {
  sheet: string;
  tables: DetectedTable[];
}

/**
 * Words that mean "this identifies the row" when a column name ends in one.
 *
 * `period` is here for the return period, which is the row identifier of every
 * exceedance table these reports print -- and is numeric, so without a word
 * saying otherwise it is discarded below as a measure. The cost of that was
 * the whole table matched by position: one return period added to the report
 * shifted every row under it, and a single configuration change read as
 * dozens of broken numbers.
 */
const KEY_WORDS = new Set([
  'id', 'no', 'nr', 'num', 'number', 'code', 'key', 'ref', 'reference',
  'period',
]);

/** Splits `PolicyId`, `policy_id` and `Policy Id` alike into words. */
const wordsOf = (name: string): string[] =>
  name
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .split(/[^A-Za-z0-9]+/)
    .filter(Boolean)
    .map((w) => w.toLowerCase());

function isKeyish(name: string): boolean {
  const words = wordsOf(name);
  if (!words.length) return false;
  return KEY_WORDS.has(words[words.length - 1]!) || KEY_WORDS.has(words[0]!);
}

const NUMERIC = /^-?[\d.,\s]+$/;

/** How far into a block to look for the header row, past any title rows. */
const HEADER_SEARCH_ROWS = 5;

const isBlank = (v: unknown): boolean =>
  v === null || v === undefined || String(v).trim() === '';

/** Flattens ExcelJS's several object-shaped cell values to something printable. */
function flat(v: unknown): string {
  if (v === null || v === undefined) return '';
  // An Invalid Date is an empty cell wearing a date format; see reader-excel.
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? '' : v.toISOString();
  if (typeof v === 'object') {
    const o = v as Record<string, any>;
    if ('richText' in o) return (o.richText as any[]).map((t) => t.text).join('');
    if ('text' in o) return String(o.text);
    if ('result' in o) return flat(o.result);
    if ('error' in o) return String(o.error);
    return '';
  }
  return String(v);
}

/**
 * Picks the columns that identify a row: every value present, and no value
 * repeated. A single column is preferred, then a pair -- beyond that a guess
 * is more likely to be wrong than useful, and saying so beats inventing one.
 */
export function detectKeyColumns(headers: string[], rows: string[][]): string[] | null {
  if (!rows.length) return null;

  const usable = headers
    .map((name, i) => ({ name, i }))
    .filter(({ name }) => name !== '');

  const values = (i: number) => rows.map((r) => (r[i] ?? '').trim());
  const complete = (vals: string[]) => vals.every((v) => v !== '');
  const distinct = (vals: string[]) => new Set(vals).size === vals.length;

  const singles = usable.filter(({ name, i }) => {
    const v = values(i);
    // Every row has to carry a value. Allowing gaps -- pairing the blanks by
    // their order of appearance, which layer 1 is willing to do -- was tried
    // against 109 real cases and took the tree from 4 failing to 12: a column
    // that identifies half its rows identifies none of them reliably, and a
    // wrong key pairs rows confidently.
    if (!complete(v) || !distinct(v)) return false;
    // A measure that happens to be distinct is not an identifier. `Amount`
    // is unique across three rows and useless as a key across three thousand,
    // so a numeric column only qualifies if its name says it identifies.
    if (v.every((x) => NUMERIC.test(x)) && !isKeyish(name)) return false;
    return true;
  });

  if (singles.length) {
    // A column named like an identifier beats one that merely happens to be
    // unique -- "Holder" is unique in a five-row sample and meaningless in a
    // five-thousand-row one.
    const named = singles.find(({ name }) => isKeyish(name));
    return [(named ?? singles[0]!).name];
  }

  // A composite key, from the leftmost few columns. Measures are excluded here
  // for the same reason as above: `Region` + `Amount` is distinct, and is still
  // not what identifies the row.
  const head = usable
    .filter(({ name, i }) => isKeyish(name) || !values(i).every((x) => NUMERIC.test(x)))
    .slice(0, 5);
  for (let a = 0; a < head.length; a++) {
    for (let b = a + 1; b < head.length; b++) {
      const first = head[a]!;
      const second = head[b]!;
      const va = values(first.i);
      const vb = values(second.i);
      if (!complete(va) || !complete(vb)) continue;
      const pair = va.map((v, r) => `${v}\u0000${vb[r]}`);
      if (distinct(pair)) return [first.name, second.name];
    }
  }
  return null;
}

/** A rectangle of the grid, 0-based and inclusive on every side. */
export interface Region {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

/**
 * How many blank columns it takes to separate two tables rather than decorate
 * one.
 *
 * A single blank column is a spacer, and a common one: a dimension breakdown
 * carries an unnamed column between its labels and its measures, and one report
 * here leaves a blank column in the middle of an eighteen-column table. Cutting
 * on one blank column fragments 186 blocks across this tree, most of them
 * wrongly. Requiring two cuts 67, and the ones it cuts are genuinely two tables
 * -- a definitions table in H:J beside a key-value block in A:B, with five
 * empty columns between them.
 */
const COLUMN_GAP = 2;

/** Whether any cell of a row within the given columns holds something. */
const rowUsed = (grid: string[][], r: number, c0: number, c1: number): boolean => {
  const row = grid[r] ?? [];
  for (let c = c0; c <= c1; c++) if (!isBlank(row[c])) return true;
  return false;
};

/** Whether any cell of a column within the given rows holds something. */
const colUsed = (grid: string[][], c: number, r0: number, r1: number): boolean => {
  for (let r = r0; r <= r1; r++) if (!isBlank((grid[r] ?? [])[c])) return true;
  return false;
};

/** Splits a region into the runs of rows that hold something. */
function byRows(grid: string[][], reg: Region): Region[] {
  const out: Region[] = [];
  let start = -1;
  for (let r = reg.r0; r <= reg.r1; r++) {
    const used = rowUsed(grid, r, reg.c0, reg.c1);
    if (used && start === -1) start = r;
    if (!used && start !== -1) {
      out.push({ ...reg, r0: start, r1: r - 1 });
      start = -1;
    }
  }
  if (start !== -1) out.push({ ...reg, r0: start, r1: reg.r1 });
  return out;
}

/**
 * Splits a region wherever COLUMN_GAP or more blank columns run through it.
 * Leading and trailing blank columns are trimmed off whatever is left.
 */
function byColumns(grid: string[][], reg: Region): Region[] {
  const out: Region[] = [];
  let start = -1;
  let gap = 0;
  for (let c = reg.c0; c <= reg.c1; c++) {
    if (colUsed(grid, c, reg.r0, reg.r1)) {
      start = start === -1 ? c : start;
      gap = 0;
      continue;
    }
    if (start === -1) continue;
    gap++;
    if (gap >= COLUMN_GAP) {
      out.push({ ...reg, c0: start, c1: c - gap });
      start = -1;
      gap = 0;
    }
  }
  if (start !== -1) out.push({ ...reg, c0: start, c1: reg.c1 });
  return out;
}

/**
 * Splits a sheet into the rectangles that hold a table.
 *
 * Blank rows separate an "output info" block from the data table below it, and
 * blank columns separate two tables printed side by side. Neither alone is
 * enough: cutting a sheet into rows leaves a definitions table in H:J fused to
 * the key-value block in A:B, and cutting it into columns leaves the info block
 * fused to the data. So both are applied, in turn, until nothing moves --
 * cutting a region by rows can expose a column gap that spanned the old region,
 * and the other way round.
 */
function regionsOf(grid: string[][], width: number): Region[] {
  let regions: Region[] = [{ r0: 0, r1: grid.length - 1, c0: 0, c1: width - 1 }];

  // Alternating passes converge quickly -- two are enough for every real sheet
  // seen here -- but the loop is bounded rather than trusted.
  for (let pass = 0; pass < 6; pass++) {
    const split = pass % 2 === 0 ? byRows : byColumns;
    const next = regions.flatMap((reg) => split(grid, reg));
    const settled =
      next.length === regions.length &&
      next.every((n, i) => {
        const o = regions[i]!;
        return n.r0 === o.r0 && n.r1 === o.r1 && n.c0 === o.c0 && n.c1 === o.c1;
      });
    regions = next;
    if (settled && pass > 0) break;
  }

  // Reading order: down the sheet, then across it.
  return regions.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
}

/**
 * Splits a sheet into tables. A table is a rectangle of the grid whose top row
 * reads as headers -- see `regionsOf` for how the rectangles are found.
 */
function tablesOnSheet(
  grid: string[][],
  sheetName: string,
  /** Cells formatted as a heading. Absent for CSV, which carries no formatting. */
  styled?: boolean[][],
  /**
   * Stored values only. A key has to come from these: the comparison never
   * sees a formula's result, so a column full of them cannot identify a row.
   */
  values?: string[][],
  /**
   * What each cell could be *called*, which is not the same as whether it holds
   * anything. A totals row of bare SUM formulas occupies twenty-five cells and
   * names none of them; counting those as names is what made the header search
   * prefer it to the real header row above. Absent for CSV, where a cell is its
   * own text.
   */
  names?: string[][],
): DetectedTable[] {
  const width = grid.reduce((w, row) => Math.max(w, row.length), 0);
  const tables: DetectedTable[] = [];

  for (const reg of regionsOf(grid, width)) {
    // A header row plus at least one row of data, across at least two columns.
    if (reg.r1 - reg.r0 < 1 || reg.c1 - reg.c0 < 1) continue;

    /** One row of the region, blanks included, so positions line up. */
    const cut = (r: number) => (grid[r] ?? []).slice(reg.c0, reg.c1 + 1);
    /** The same row, as names rather than as occupancy. */
    const nameCut = (r: number) => ((names ?? grid)[r] ?? []).slice(reg.c0, reg.c1 + 1);

    // The header is the *fullest* row near the top of the block, not
    // necessarily the first. Real reports put a title above the table -- often
    // several -- and a title occupies one cell. Taking the first row blindly
    // discards the whole block when it is a title, header and all.
    //
    // The row with the most cells that read as a *name* wins, not the fullest
    // row. A data row is often the wider of the two -- it fills every column,
    // where a header may leave some blank -- but most of what fills it is
    // measurements. Counting names instead separates them: a header row is
    // labels, a data row is numbers with a few labels down its left side.
    //
    // Without this, detection reads a row of postcodes as the column names,
    // and those pair against different numbers in the other file and report
    // the whole table as changed.
    // Formatting decides it when there is any, because these generators mark
    // their headers deliberately -- bold white on navy -- and nothing else on
    // the sheet is painted that way. Counting cells alone gets a small info
    // block wrong: its header holds two words and the data row beneath holds
    // three, so the data row wins and the columns end up named after values.
    let headerAt = -1;
    let best = 0;
    let bestStyled = 0;
    let fallback = reg.r0;
    let fallbackWidth = -1;
    const limit = Math.min(reg.r0 + HEADER_SEARCH_ROWS, reg.r1);

    for (let r = reg.r0; r <= limit; r++) {
      const cells = cut(r).filter((v) => !isBlank(v));
      // A row holding one cell names one column, which is not a table, and a
      // row holding none names nothing. Neither can be the header -- the check
      // below rejects exactly these -- so they must not be allowed to win the
      // search either. They used to: a section title is painted like a heading
      // and the key-value block beneath it is not, so a bold "Exchange Rate
      // Information" alone in column A beat the four labelled rows under it,
      // and the block was then thrown away for having a one-column header. The
      // whole table went unseen by layer 1, with nothing said about it.
      if (cells.length < 2) continue;
      // Strictly greater throughout, so the earliest row wins a tie -- a
      // header and its data rows are often equally full.
      if (cells.length > fallbackWidth) {
        fallbackWidth = cells.length;
        fallback = r;
      }
      const named = nameCut(r).filter((v) => !isBlank(v) && !NUMERIC.test(v)).length;
      if (!named) continue; // a row of numbers is data however it is painted

      const painted = (styled?.[r] ?? []).slice(reg.c0, reg.c1 + 1).filter(Boolean).length;
      if (painted > bestStyled || (painted === bestStyled && named > best)) {
        bestStyled = painted;
        best = named;
        headerAt = r;
      }
    }

    // Every candidate row was numbers. Keep the old behaviour rather than
    // dropping the block: a wrong header row still beats no comparison, and
    // the names it produces are visible in the report.
    if (headerAt === -1) headerAt = fallback;
    if (headerAt >= reg.r1) continue; // nothing left below it to be data

    // Some blocks have no header row at all, and picking one anyway does two
    // things wrong: the row picked stops being data, and the value column is
    // named after whatever value happens to sit in it. On a report info block
    // that name landed on the report's own name -- which differs by
    // construction between two runs -- so the column matched nothing in the
    // other file and all fifteen rows reported as one column removed and
    // another added, with the real changes buried among them.
    // A row that names fewer than two columns cannot be the header of a table,
    // and dropping the block for it is the wrong answer: the rows are still
    // there and still comparable, they just have no names of their own. That
    // happens wherever a block's top row is formulas the generator never
    // computed -- common enough that refusing cost 26 tables across this tree,
    // silently, which is the one outcome worse than naming them positionally.
    const namesEnough = (r: number) =>
      nameCut(r).filter((v) => !isBlank(v)).length >= 2;

    const noHeader =
      paintedByColumn(styled, grid, reg) ||
      readsAsData(grid, reg, headerAt) ||
      !namesEnough(headerAt);

    // `headerRow` is the row *above* the data throughout, so a block with no
    // header is one whose header row is the row above it -- 0 when it starts
    // at the top of the sheet. Nothing downstream needs a new concept.
    const top = noHeader ? reg.r0 : headerAt + 1;

    const source = values ?? grid;
    const body: string[][] = [];
    for (let r = top; r <= reg.r1; r++) {
      body.push((source[r] ?? []).slice(reg.c0, reg.c1 + 1));
    }

    // Named after themselves, which is what the reader does with a blank
    // header row, so both agree what a key column is called.
    //
    // Whether a column holds anything is asked of the grid rather than of
    // `body`. `body` carries stored values only, so a column of formulas the
    // generator never computed reads as empty there -- while the reader counts
    // a formula cell as data when it names a column after itself. Asking the
    // wrong one cost a four-row block of layer premiums its name and then its
    // place: eleven columns became one, and the block was dropped for being
    // too narrow to be a table.
    const holdsData = (i: number) => {
      for (let r = top; r <= reg.r1; r++) {
        if (!isBlank((grid[r] ?? [])[reg.c0 + i])) return true;
      }
      return false;
    };
    const headers = noHeader
      ? Array.from({ length: reg.c1 - reg.c0 + 1 }, (_, i) =>
          holdsData(i) ? `Column ${numToCol(reg.c0 + i + 1)}` : '',
        )
      : nameCut(headerAt).map((h) => String(h ?? '').trim());
    if (headers.filter((h) => h !== '').length < 2) continue;

    tables.push({
      name: '',
      headerRow: top,
      endRow: reg.r1 + 1,
      columns: `${numToCol(reg.c0 + 1)}:${numToCol(reg.c1 + 1)}`,
      region: reg,
      headers,
      keyColumns: detectKeyColumns(headers, body),
      rows: body.length,
      filledPerColumn: headers.map((_, i) => body.filter((r) => !isBlank(r[i])).length),
      firstRowFilled: headers.map((_, i) => !isBlank(body[0]?.[i])),
    });
  }

  // A sheet with one table is named after the sheet, so reports read as
  // "Policies" rather than "Policies · Table 1".
  tables.forEach((t, i) => {
    t.name = tables.length === 1 ? sheetName : `Table ${i + 1}`;

    // A table with nothing below it runs to the end of the sheet, so it keeps
    // working as its data grows. "Below" has to mean below *in its own
    // columns*: a table in H:J may have nothing under it while another one
    // fills A:BX further down, and letting either run to the bottom would
    // swallow the other.
    const covered = tables.some(
      (o) => o !== t && o.region.r0 > t.region.r1 &&
        o.region.c0 <= t.region.c1 && o.region.c1 >= t.region.c0,
    );
    if (!covered) t.endRow = 0;
  });
  return tables;
}

/**
 * Reads a worksheet into a plain grid of strings.
 *
 * Two things would otherwise read as blank when they are not.
 *
 * A merged range reports the master's value from every cell it covers, so a
 * banner like "PERSONAL LINES" merged across nineteen columns reads back as
 * nineteen identical headers. That is the widest row on the sheet, so the
 * header search picks it over the real header row underneath, and every column
 * ends up with the same name. Counting only the master leaves the banner one
 * cell wide, which is what it is.
 *
 * And a formula cell in a file the generator wrote holds no computed result, so
 * taking its value alone makes it invisible. A row of nothing but formulas then
 * looks like a blank row, which splits the block it sits in and strands the
 * real header row above the split -- on one report that left detection reading
 * a row of postcodes as the column names. The cell is described the same way
 * the reader describes it, so the two agree about what is there.
 */
/**
 * Whether a cell is painted the way these reports paint a heading: bold, or
 * filled with a colour chosen for it rather than inherited from the theme.
 */
function looksLikeHeading(cell: ExcelJS.Cell): boolean {
  if (cell.font?.bold) return true;
  const fill = cell.fill as { type?: string; fgColor?: { argb?: string } } | undefined;
  return Boolean(fill?.type === 'pattern' && fill.fgColor?.argb);
}

/**
 * Whether a block's painting marks out a column rather than a row.
 *
 * A header row is painted to stand out from its data -- bold white on navy
 * across the width of the table, and nothing beneath it. A key-value block is
 * the other shape: the label column is painted on every row, so the paint says
 * "this column holds the labels", not "this row is the header". Told apart by
 * asking whether every row is painted identically; if one differs, some row
 * stands out and the search that found it was right.
 *
 * A block nobody painted is excluded, and that exclusion is the point rather
 * than an oversight. Every row of an unstyled table is painted identically
 * too -- not at all -- so without it every report that arrives as a plain grid
 * would lose its column names.
 */
/**
 * Whether the chosen header row is really just another row of data.
 *
 * The paint heuristic is defeated by banded shading. A report that fills every
 * other row for legibility paints those rows exactly the way it paints a
 * heading, so the search takes the first banded row as the header. On one real
 * sheet that made the key column `US - Northeast` -- a region name, lifted from
 * the data -- and named the rest of the columns after that row's figures, which
 * then drifted between runs and reported the whole table as columns added and
 * removed.
 *
 * Told apart by shape rather than by paint. A header row is text where its data
 * is numeric; a data row has the same numeric columns as the row beneath it.
 * So: if the candidate's numeric columns are exactly the next row's, it is
 * data.
 *
 * The candidate must actually hold a number for this to fire. A header of plain
 * words above rows of plain words -- `Name | City` over `Ivanov | Sofia` --
 * matches on shape too, and there the heading is real; refusing it would strip
 * the names off every text table in the tree.
 */
function readsAsData(grid: string[][], reg: Region, headerAt: number): boolean {
  const numericAt = (r: number): string => {
    const row = (grid[r] ?? []).slice(reg.c0, reg.c1 + 1);
    return row.map((v, i) => (!isBlank(v) && NUMERIC.test(v) ? i : -1)).filter((i) => i >= 0).join(',');
  };
  const here = numericAt(headerAt);
  if (!here) return false; // no numbers of its own: a heading of words is a heading

  // The first row under it that holds anything, which is what it must differ
  // from to be a header at all.
  for (let r = headerAt + 1; r <= reg.r1; r++) {
    const row = (grid[r] ?? []).slice(reg.c0, reg.c1 + 1);
    if (row.every((v) => isBlank(v))) continue;
    return numericAt(r) === here;
  }
  return false;
}

function paintedByColumn(
  styled: boolean[][] | undefined,
  grid: string[][],
  reg: Region,
): boolean {
  if (!styled) return false; // CSV carries no formatting to read
  let signature: string | null = null;
  for (let r = reg.r0; r <= reg.r1; r++) {
    const row = (grid[r] ?? []).slice(reg.c0, reg.c1 + 1);
    if (row.every((v) => isBlank(v))) continue;
    const painted = (styled[r] ?? [])
      .slice(reg.c0, reg.c1 + 1)
      .map((on, i) => (on ? i : -1))
      .filter((i) => i >= 0)
      .join(',');
    if (!painted) return false;
    if (signature === null) signature = painted;
    else if (painted !== signature) return false;
  }
  return signature !== null;
}

/**
 * Two readings of the same sheet, because two different questions are being
 * asked of it.
 *
 * `text` is what a person sees: a formula cell contributes the text it would
 * produce, so a row of computed headings still reads as a heading row and does
 * not look blank.
 *
 * `values` is what the comparison will actually have to work with -- stored
 * values only, blank where a formula has no result. Key detection has to use
 * this one. Judging a key on `text` picks columns that look full and turn out
 * empty, and every row keyed on one is dropped for having a blank key: one
 * sheet detected a key, then compared none of its 888 rows.
 */
function gridOf(ws: ExcelJS.Worksheet): {
  text: string[][];
  values: string[][];
  names: string[][];
  styled: boolean[][];
} {
  const text: string[][] = [];
  const values: string[][] = [];
  const names: string[][] = [];
  const styled: boolean[][] = [];
  const width = Math.max(ws.columnCount, 1);

  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const cells: string[] = [];
    const stored: string[] = [];
    const named: string[] = [];
    const marks: boolean[] = [];
    for (let c = 1; c <= width; c++) {
      const cell = row.getCell(c);
      if (cell.isMerged && cell.master !== cell) {
        cells.push(''); stored.push(''); named.push(''); marks.push(false);
        continue;
      }
      const value = flat(cell.value).trim();
      const shown = value || cellToken(cell);
      cells.push(shown);
      stored.push(cell.formula ? '' : value);
      // What this cell could be called, which is not the same as whether it
      // holds anything. A totals row of bare SUM formulas occupies its cells
      // and names none of them.
      named.push(value || headerName(cell));
      marks.push(shown !== '' && looksLikeHeading(cell));
    }
    text.push(cells);
    values.push(stored);
    names.push(named);
    styled.push(marks);
  }
  return { text, values, names, styled };
}

export async function detectWorkbook(path: string): Promise<DetectedSheet[]> {
  const ext = extname(path).toLowerCase();

  if (['.csv', '.tsv', '.txt'].includes(ext)) {
    const text = (await readFile(path)).toString('utf8').replace(/^﻿/, '');
    const grid: string[][] = parse(text, {
      relax_column_count: true,
      skip_empty_lines: true,
      bom: true,
      trim: true,
    });
    return [{ sheet: CSV_SHEET, tables: tablesOnSheet(grid, CSV_SHEET) }];
  }

  const wb = await openWorkbook(path);
  return wb.worksheets.map((ws) => {
    const { text, values, names, styled } = gridOf(ws);
    return { sheet: ws.name, tables: tablesOnSheet(text, ws.name, styled, values, names) };
  });
}

/** Turns a detection into the spec the comparison actually runs on. */
export function specFromDetection(sheets: DetectedSheet[]): WorkbookSpec {
  const out: Record<string, WorkbookSheetSpec> = {};

  for (const s of sheets) {
    if (!s.tables.length) continue;

    if (s.tables.length === 1) {
      const t = s.tables[0]!;
      out[s.sheet] = {
        headerRow: t.headerRow,
        ...(t.keyColumns ? { keyColumns: t.keyColumns } : {}),
      };
      continue;
    }

    // A column bound is only written when the sheet actually puts tables side
    // by side -- which is exactly when two of them share any row. Writing
    // "A:Q" on every table of a stacked sheet would freeze a width that grows,
    // and tell a reader of the spec nothing.
    const overlap = (a: Region, b: Region) => a.r0 <= b.r1 && b.r0 <= a.r1;
    const sideBySide = s.tables.some((t) =>
      s.tables.some((o) => o !== t && overlap(t.region, o.region)),
    );

    const tables: Record<string, WorkbookSheetSpec> = {};
    for (const t of s.tables) {
      tables[t.name] = {
        headerRow: t.headerRow,
        ...(t.endRow ? { endRow: t.endRow } : {}),
        ...(sideBySide ? { columns: t.columns } : {}),
        ...(t.keyColumns ? { keyColumns: t.keyColumns } : {}),
      };
    }
    out[s.sheet] = { tables };
  }

  return { sheets: out };
}

/** Detection for a file, as a ready-to-edit spec. */
export async function detectSpec(path: string): Promise<WorkbookSpec> {
  return specFromDetection(await detectWorkbook(path));
}
