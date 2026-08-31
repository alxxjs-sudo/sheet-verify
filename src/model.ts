import type {
  Cell, CellValue, ResolvedSpec, Row, SheetModel, SheetSpec,
} from './types.js';
import { numToCol, toHeaderRef, toR1C1 } from './a1.js';

export const KEY_SEP = '␟';

/**
 * The rectangle a table actually occupies, as an A1 range: "H2:J22".
 *
 * Read off the model rather than the spec, so it says what was *read* and not
 * what was asked for. A spec with no `endRow` runs to the bottom of the sheet
 * and a spec with no `columns` spans its whole width; neither is a useful
 * answer to "what did this comparison cover". The header row is included when
 * there was one, because it is part of the table even though its cells are
 * compared as column names rather than as values.
 *
 * Empty when the table has no columns at all, which is a table that could not
 * be read and has nothing to report.
 */
const POSITIONAL_HEADER = /^Column [A-Z]+$/;

export function tableRange(model: SheetModel, headerRow: number): string {
  const cols = [...model.headerIndex.values()];
  if (!cols.length) return '';

  const first = Math.min(...cols);
  const last = Math.max(...cols);

  /** Row number of the first row found from `start`, stepping by `step`. */
  const rowFrom = (start: number, step: number): number | null => {
    for (let i = start; i >= 0 && i < model.order.length; i += step) {
      const row = model.rows.get(model.order[i]!);
      const address = row && Object.values(row)[0]?.address;
      const m = address && /^[A-Z]+(\d+)$/.exec(address);
      if (m) return Number(m[1]);
    }
    return null;
  };

  // Walking back from the end covers a duplicate key, which is left out of
  // `rows` while keeping its place in `order`.
  const bottom = rowFrom(model.order.length - 1, -1);
  const firstData = rowFrom(0, 1);

  // A header row belongs to the table only when it named the columns. Where it
  // named nothing -- `headerRow` 0, or the blank separator row that stands in
  // for one above a block with no headers of its own -- every column is named
  // after itself and the table is its data alone. Including a row the
  // comparison never read would overstate what was covered by exactly one row.
  const named = model.headers.some((h) => !POSITIONAL_HEADER.test(h));
  const top = headerRow > 0 && named ? headerRow : (firstData ?? Math.max(headerRow, 1));

  return `${numToCol(first)}${top}:${numToCol(last)}${Math.max(bottom ?? top, top)}`;
}

/**
 * Marks the nth row sharing a key, e.g. "Total" then "Total ×2".
 *
 * Real reports repeat a key legitimately: a breakdown carries one "Total" row
 * per group, and every one of them has the same blanks in the columns that
 * identify a row. Dropping those rows -- which is what this used to do -- threw
 * away the totals, the very lines most people read first. Numbering them
 * instead pairs the nth "Total" in the baseline with the nth in the new report,
 * which is right whenever the groups are in the same order.
 */
export const KEY_OCCURRENCE = ' ×';

/**
 * Tolerance applied when a config names none.
 *
 * A thousandth of a unit. What it absorbs is recalculation noise -- a total
 * rebuilt in a different order lands a few bits away from the one stored last
 * month, a gap of about 1e-7 on a figure in the hundreds of millions -- while
 * staying well under a cent, so a change anyone wrote down on purpose is still
 * reported.
 *
 * This is not the hidden float slack that used to live in `equalValues`. That
 * one was invisible and scaled with the value, so nobody could say what it had
 * swallowed. This is a plain number, written down, reported wherever it
 * applies -- the run counts those cells and lists them -- and overridden by
 * `tolerance` in any meta.json. Set it to 0 there for exactness.
 */
export const DEFAULT_TOLERANCE = 0.001;

/**
 * Relative tolerance applied when a config names none: off.
 *
 * `tolerance` asks how many units apart two numbers are. That question has no
 * single right answer across a report whose figures run from 0.0002 to 1.3e11,
 * because Excel stores 15 significant digits and nothing more: at 1.3e11 the
 * last digit it can hold sits at the thousandths, so a total rebuilt in a
 * different order lands whole thousandths away and a flat 0.001 calls it a
 * difference. At 5,000,000 that same 0.001 is a hundred thousand times looser
 * than the format's own precision, and a real error of a cent goes unreported.
 *
 * Measured on this project's own cases, the effect is stark. Of the cells whose
 * gap is pure recalculation drift, the ABSOLUTE gaps span fourteen orders of
 * magnitude while the RELATIVE gaps span two, all of them between 1e-16 and
 * 5.4e-14 -- which is to say they are all the same phenomenon, and only the
 * magnitude of the number they sit on makes them look different. Judged by
 * `tolerance` alone, five drifting cells were absorbed in silence while another
 * with LESS relative drift was reported as a difference, decided by nothing but
 * where the decimal point fell.
 *
 * So a cell may now also be forgiven for being close in proportion:
 *
 *     |a - b| <= max(tolerance, relativeTolerance * max(|a|, |b|))
 *
 * `tolerance` stays as the floor, because a proportion of almost-nothing is
 * meaningless and a value near zero needs an absolute answer.
 *
 * This is deliberately NOT the hidden float slack that used to live in
 * `equalValues`, though the arithmetic is a cousin of it. That one was on for
 * everybody, was never written down, and could not be seen doing its work --
 * "nobody could say what it had swallowed" was the reason it went. This one
 * defaults to 0, so a config that does not ask for it compares exactly as
 * before; it is written in a meta.json where it can be read; and every cell it
 * forgives is counted and listed with the allowance that forgave it, which for
 * a relative rule is computed per cell rather than per column.
 */
export const DEFAULT_RELATIVE_TOLERANCE = 0;

/**
 * The gap two numbers are allowed, given both rules. The larger of the two
 * wins: they are alternatives, not conditions, and a cell close enough by
 * either measure is close enough.
 */
export function allowance(a: number, b: number, tol: number, rel: number): number {
  if (!(rel > 0)) return tol;
  return Math.max(tol, rel * Math.max(Math.abs(a), Math.abs(b)));
}

export function resolveSpec(spec: SheetSpec): ResolvedSpec {
  if (!spec.keyColumns?.length && !spec.matchRowsByPosition) {
    throw new Error(
      'sheet-verify: `keyColumns` is required. Row identity cannot be inferred; ' +
      'without it a diff degrades to positional comparison, which is the problem ' +
      'this library exists to avoid. Set matchRowsByPosition to accept that ' +
      'trade deliberately for a table that has no key.',
    );
  }
  const tol =
    typeof spec.tolerance === 'number'
      ? { '*': spec.tolerance }
      : { '*': DEFAULT_TOLERANCE, ...(spec.tolerance ?? {}) };
  const rel =
    typeof spec.relativeTolerance === 'number'
      ? { '*': spec.relativeTolerance }
      : { '*': DEFAULT_RELATIVE_TOLERANCE, ...(spec.relativeTolerance ?? {}) };

  return {
    sheet: spec.sheet ?? 0,
    headerRow: spec.headerRow ?? 1,
    endRow: spec.endRow ?? 0,
    columns: spec.columns ?? '',
    keyColumns: spec.keyColumns ?? [],
    matchRowsByPosition: spec.matchRowsByPosition ?? false,
    fillKeyDown: spec.fillKeyDown ?? false,
    keySeparator: spec.keySeparator ?? KEY_SEP,
    tolerance: tol,
    relativeTolerance: rel,
    ignoreColumns: spec.ignoreColumns ?? [],
    ignoreRows: spec.ignoreRows ?? [],
    metadataCells: spec.metadataCells ?? [],
    compareFormulas: spec.compareFormulas ?? true,
    formulaMode: spec.formulaMode ?? 'header',
    requireCachedValues: spec.requireCachedValues ?? true,
    trimStrings: spec.trimStrings ?? true,
    looseHeaders: spec.looseHeaders ?? true,
    strictSchema: spec.strictSchema ?? false,
    invariants: spec.invariants ?? [],
    csv: {
      delimiter: spec.csv?.delimiter ?? '',
      numeric: spec.csv?.numeric ?? 'auto',
      strictDialect: spec.csv?.strictDialect ?? false,
    },
  };
}

/**
 * A header that is a number, as a name.
 *
 * Reports head columns with computed figures -- a loss threshold, a return
 * period, a currency rate -- and those recalculate. Two files then hold the
 * same column under names that differ only in the last digits:
 * `788321.400221` against `788321.4002209998`, a gap of 1.16e-10.
 *
 * Compared as text those are two different columns, so layer 1 reports one
 * removed and one added and every cell beneath them is reported twice, with
 * the real findings buried among hundreds of rows of noise. The irony is that
 * a tolerance would have forgiven the same drift instantly had the number been
 * a *value* rather than a *name*.
 *
 * Twelve significant figures because Excel keeps about fifteen and drift shows
 * up past twelve. Two genuinely different headers agreeing to twelve figures
 * would be merged, which is possible and has never been seen: column headings
 * are not distinguished by their thirteenth digit.
 */
const NUMERIC_NAME = /^-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?$/;
const HEADER_FIGURES = 12;

export function canonHeader(name: string, loose: boolean): string {
  const t = name.trim();
  // Only under `looseHeaders`, which is what asks for representation to be
  // forgiven. Turning it off still means matching exactly what is written.
  if (loose && NUMERIC_NAME.test(t)) {
    const n = Number(t);
    if (Number.isFinite(n)) return n.toPrecision(HEADER_FIGURES);
  }
  return loose ? t.toLowerCase().replace(/\s+/g, ' ') : t;
}

/**
 * A row key reduced for `ignoreRows` matching. Trailing colons go too, since a
 * key-value block writes "Report ID" or "Report ID:" interchangeably.
 */
export function canonRowKey(key: string): string {
  return key.trim().toLowerCase().replace(/\s+/g, ' ').replace(/:$/, '');
}

/**
 * A test for `ignoreRows`, matching the way `canonRowKey` normalises and
 * allowing `*` to stand for a run of text.
 *
 * The wildcard exists because a metadata label folded in from `metadata` may
 * carry one -- every report type spells its own name differently, and listing
 * each spelling is how a config quietly stops covering the newest one.
 */
export function rowKeyMatcher(patterns: string[]): (key: string) => boolean {
  const exact = new Set<string>();
  const globs: RegExp[] = [];
  for (const pattern of patterns) {
    const p = canonRowKey(pattern);
    if (!p.includes('*')) { exact.add(p); continue; }
    const body = p
      .split('*')
      .map((part) => part.replace(/[.+?^${}()|[\]\\]/g, '\\$&'))
      .join('.*?');
    globs.push(new RegExp(`^${body}$`));
  }
  if (!globs.length) return (key) => exact.has(canonRowKey(key));
  return (key) => {
    const k = canonRowKey(key);
    return exact.has(k) || globs.some((g) => g.test(k));
  };
}

/**
 * Whether two cell values are the same. Shared, deliberately: the comparer
 * and the cell ledger are two views of one run, and a ledger that applies a
 * stricter rule reports differences the report says do not exist.
 */
export function equalValues(
  a: CellValue,
  b: CellValue,
  tol: number,
  rel = 0,
): boolean {
  if (a === null && b === null) return true;
  if (a === null || b === null) return false;
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if (typeof a === 'number' && typeof b === 'number') {
    if (Number.isNaN(a) && Number.isNaN(b)) return true;
    // Two rules, either of which forgives: `tolerance` in units and
    // `relativeTolerance` in proportion. See DEFAULT_RELATIVE_TOLERANCE for
    // why one number cannot cover a report spanning 0.0002 to 1.3e11.
    //
    // Both are the reader's to set and both are reported. Neither is a hidden
    // allowance: `relativeTolerance` is 0 unless a config asks for it, and a
    // cell either forgives is counted and listed with the allowance that
    // applied to it. Judging which differences matter stays the reader's job.
    return Math.abs(a - b) <= allowance(a, b, tol, rel);
  }
  if (typeof a === 'boolean' || typeof b === 'boolean') return a === b;
  return String(a) === String(b);
}

export function toleranceFor(spec: ResolvedSpec, column: string): number {
  return spec.tolerance[column] ?? spec.tolerance['*'] ?? 0;
}

export function relativeToleranceFor(spec: ResolvedSpec, column: string): number {
  return spec.relativeTolerance[column] ?? spec.relativeTolerance['*'] ?? 0;
}

export interface RawCell {
  address: string;
  kind: Cell['kind'];
  value: CellValue;
  formula: string | null;
}

/**
 * Shared assembly step. Readers produce headers plus a per-row cell map;
 * this turns that into the keyed, normalised model the comparer consumes.
 */
export function buildModel(args: {
  source: string;
  sheet: string;
  headers: { name: string; colNum: number }[];
  rows: { rowNum: number; cells: Map<number, RawCell> }[];
  spec: ResolvedSpec;
}): SheetModel {
  const { source, sheet, spec } = args;

  const seen = new Map<string, number>();
  const duplicateHeaders: string[] = [];
  const headers: string[] = [];
  const headerIndex = new Map<string, number>();
  const headerByCol = new Map<number, string>();

  for (const h of args.headers) {
    const canon = canonHeader(h.name, spec.looseHeaders);
    if (!canon) continue;
    const count = (seen.get(canon) ?? 0) + 1;
    seen.set(canon, count);
    // Disambiguate rather than silently dropping: a duplicated header is a
    // real defect, but the rest of the sheet should still be comparable.
    const name = count === 1 ? h.name.trim() : `${h.name.trim()} (#${count})`;
    if (count === 2) duplicateHeaders.push(h.name.trim());
    headers.push(name);
    headerIndex.set(name, h.colNum);
    headerByCol.set(h.colNum, name);
  }

  const rows = new Map<string, Row>();
  const order: string[] = [];
  const dupes = new Set<string>();
  const uncached: string[] = [];
  const occurrences = new Map<string, number>();
  /** Last non-blank value seen in each key column. See `fillKeyDown`. */
  const carried = new Map<string, string>();

  let ordinal = 0;

  for (const r of args.rows) {
    const rec: Row = {};
    for (const name of headers) {
      const colNum = headerIndex.get(name)!;
      const raw = r.cells.get(colNum);
      rec[name] = raw
        ? finalize(raw, r.rowNum, colNum, headerByCol, spec, uncached)
        : { address: '', kind: 'empty', value: null, formula: null, r1c1: null, headerRef: null };
    }

    let key: string;
    if (spec.matchRowsByPosition) {
      // Counted from the top of the table rather than taken from the sheet row
      // number, so a table that starts lower in one file still lines up. Blank
      // rows take an ordinal too: skipping them would shift everything under
      // them out of step with the other side, which is the one thing
      // positional matching cannot survive.
      key = `#${++ordinal}`;
    } else {
      const raw = spec.keyColumns.map((k) => {
        const c = rec[k];
        return c && c.value !== null && c.value !== undefined ? String(c.value).trim() : '';
      });

      // A spacer row must not inherit the group above it -- it would take that
      // group's key and collide with the row that legitimately holds it. Only
      // a row with something in it is treated as belonging to the heading.
      const empty = Object.values(rec).every(
        (c) => c.value === null || c.value === undefined || String(c.value).trim() === '',
      );
      if (empty) continue;

      const parts = spec.fillKeyDown
        ? raw.map((p, i) => {
            const column = spec.keyColumns[i]!;
            if (p !== '') { carried.set(column, p); return p; }
            return carried.get(column) ?? '';
          })
        : raw;

      if (parts.every((p) => p === '')) continue; // nothing identifies this row
      key = parts.join(spec.keySeparator);
    }

    // A repeated key is numbered rather than discarded, so the nth row bearing
    // it lines up with the nth on the other side. See KEY_OCCURRENCE.
    const n = (occurrences.get(key) ?? 0) + 1;
    occurrences.set(key, n);
    if (n > 1) dupes.add(key);
    const unique = n === 1 ? key : `${key}${KEY_OCCURRENCE}${n}`;
    rows.set(unique, rec);
    order.push(unique);
  }

  // Trailing blank rows are an artefact of the file, not of the report.
  //
  // Excel's used range outlives its contents: a sheet edited down to thirteen
  // rows still reports thirty if something once occupied them, and the two
  // files being compared have no reason to agree on that number. Under
  // positional matching every one of those phantom rows takes an ordinal, so a
  // disclaimer identical in both arrives as seventeen removed rows.
  //
  // Only the trailing run goes. A blank row *between* two populated ones is
  // load-bearing: dropping it would shift everything beneath it out of step
  // with the other side, which is the one thing positional matching cannot
  // survive. Keyed tables never get here -- they discard blank rows outright,
  // having no key to file them under.
  if (spec.matchRowsByPosition) {
    const blank = (key: string) =>
      Object.values(rows.get(key) ?? {}).every(
        (c) => c.value === null || c.value === undefined || String(c.value).trim() === '',
      );
    while (order.length && blank(order[order.length - 1]!)) rows.delete(order.pop()!);
  }

  return {
    source, sheet, headers, headerIndex, rows,
    order: order.filter((k) => rows.has(k)),
    duplicateKeys: [...dupes],
    duplicateHeaders,
    uncachedFormulaCells: uncached,
  };
}

function finalize(
  raw: RawCell,
  rowNum: number,
  colNum: number,
  headerByCol: ReadonlyMap<number, string>,
  spec: ResolvedSpec,
  uncached: string[],
): Cell {
  let value = raw.value;
  if (spec.trimStrings && typeof value === 'string') value = value.trim();
  if (value === '') value = null;

  if (raw.formula) {
    if (value === null || value === undefined) uncached.push(raw.address);
    // Excel strips whitespace around a formula when it saves, so a generator
    // that wrote " IF(...)" comes back as "IF(...)" from any file a person has
    // opened. Both compute the same thing, and reporting it as a calculation
    // change would flag every such formula on every run. Only the ends are
    // touched -- spacing inside the text can sit within a string literal,
    // where it is part of the output.
    const formula = raw.formula.trim();
    return {
      address: raw.address,
      kind: 'formula',
      value,
      formula,
      r1c1: toR1C1(formula, rowNum, colNum),
      headerRef: toHeaderRef(formula, rowNum, colNum, headerByCol),
    };
  }

  return {
    address: raw.address,
    kind: value === null ? 'empty' : raw.kind,
    value,
    formula: null,
    r1c1: null,
    headerRef: null,
  };
}

/** Splits a composite key back into its parts, for readable reporting. */
export function displayKey(key: string, spec: ResolvedSpec): string {
  return key.split(spec.keySeparator).join(' / ');
}
