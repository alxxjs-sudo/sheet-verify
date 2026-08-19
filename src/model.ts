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

/** Canonical form used to match header names across releases. */
export function canonHeader(name: string, loose: boolean): string {
  const t = name.trim();
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
export function equalValues(a: CellValue, b: CellValue, tol: number): boolean {
  if (a === null && b === null) return true;
  if (a === null || b === null) return false;
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if (typeof a === 'number' && typeof b === 'number') {
    if (Number.isNaN(a) && Number.isNaN(b)) return true;
    // Only `tolerance` absorbs a gap. It defaults to DEFAULT_TOLERANCE rather
    // than to zero, and every cell it forgives is counted and listed.
    //
    // This used to carry slack of scale x 1e-12 as well, on the reasoning that
    // Excel keeps 15 significant digits so a smaller gap is rounding from a
    // different order of operations rather than a change. True as far as it
    // goes, and the wrong trade: a cell whose stored values differ is a cell
    // that differs, and deciding on the reader's behalf that they did not want
    // to know leaves the report unable to show the full reach of a change.
    // Judging which differences matter is the reader's job; `tolerance` is how
    // they say so, per column, deliberately.
    return Math.abs(a - b) <= tol;
  }
  if (typeof a === 'boolean' || typeof b === 'boolean') return a === b;
  return String(a) === String(b);
}

export function toleranceFor(spec: ResolvedSpec, column: string): number {
  return spec.tolerance[column] ?? spec.tolerance['*'] ?? 0;
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
