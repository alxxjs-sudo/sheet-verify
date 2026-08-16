import type {
  Cell, CellValue, ResolvedSpec, Row, SheetModel, SheetSpec,
} from './types.js';
import { toHeaderRef, toR1C1 } from './a1.js';

export const KEY_SEP = '␟';

export function resolveSpec(spec: SheetSpec): ResolvedSpec {
  if (!spec.keyColumns?.length) {
    throw new Error(
      'sheet-verify: `keyColumns` is required. Row identity cannot be inferred; ' +
      'without it a diff degrades to positional comparison, which is the problem ' +
      'this library exists to avoid.',
    );
  }
  const tol =
    typeof spec.tolerance === 'number'
      ? { '*': spec.tolerance }
      : { '*': 0, ...(spec.tolerance ?? {}) };

  return {
    sheet: spec.sheet ?? 0,
    headerRow: spec.headerRow ?? 1,
    keyColumns: spec.keyColumns,
    keySeparator: spec.keySeparator ?? KEY_SEP,
    tolerance: tol,
    ignoreColumns: spec.ignoreColumns ?? [],
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

  for (const r of args.rows) {
    const rec: Row = {};
    for (const name of headers) {
      const colNum = headerIndex.get(name)!;
      const raw = r.cells.get(colNum);
      rec[name] = raw
        ? finalize(raw, r.rowNum, colNum, headerByCol, spec, uncached)
        : { address: '', kind: 'empty', value: null, formula: null, r1c1: null, headerRef: null };
    }

    const parts = spec.keyColumns.map((k) => {
      const c = rec[k];
      return c && c.value !== null && c.value !== undefined ? String(c.value) : '';
    });
    if (parts.every((p) => p === '')) continue; // blank row
    const key = parts.join(spec.keySeparator);

    if (rows.has(key)) dupes.add(key);
    else { rows.set(key, rec); order.push(key); }
  }

  for (const k of dupes) rows.delete(k);

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
    return {
      address: raw.address,
      kind: 'formula',
      value,
      formula: raw.formula,
      r1c1: toR1C1(raw.formula, rowNum, colNum),
      headerRef: toHeaderRef(raw.formula, rowNum, colNum, headerByCol),
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
