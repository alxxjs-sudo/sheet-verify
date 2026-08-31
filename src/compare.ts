import type {
  Cell, CellValue, DiffResult, FormulaDiff, InvariantFailure,
  MovedColumn, ResolvedSpec, SheetModel, TypeDiff, ValueDiff,
} from './types.js';
import {
  canonHeader, equalValues, relativeToleranceFor, rowKeyMatcher, toleranceFor,
} from './model.js';
import { dialectDrift } from './reader-csv.js';

/** Same-row column dependencies, read off the header-resolved formula. */
const SAME_ROW_REF = /\[([^\]]+)\]@row(?![+\-])/g;

function dependencies(cell: Cell): string[] {
  if (!cell.headerRef) return [];
  const out: string[] = [];
  for (const m of cell.headerRef.matchAll(SAME_ROW_REF)) out.push(m[1]!);
  return out;
}


function normalizedFormula(cell: Cell, mode: ResolvedSpec['formulaMode']): string | null {
  return mode === 'a1' ? cell.formula : mode === 'r1c1' ? cell.r1c1 : cell.headerRef;
}

/** Matches headers across the two models, honouring loose matching. */
function pairColumns(base: SheetModel, next: SheetModel, spec: ResolvedSpec) {
  const canon = (h: string) => canonHeader(h, spec.looseHeaders);
  const nextByCanon = new Map(next.headers.map((h) => [canon(h), h]));
  const baseByCanon = new Map(base.headers.map((h) => [canon(h), h]));

  const shared: { base: string; next: string }[] = [];
  for (const h of base.headers) {
    const other = nextByCanon.get(canon(h));
    if (other !== undefined) shared.push({ base: h, next: other });
  }
  const added = next.headers.filter((h) => !baseByCanon.has(canon(h)));
  const removed = base.headers.filter((h) => !nextByCanon.has(canon(h)));
  return { shared, added, removed };
}

export function compare(base: SheetModel, next: SheetModel, spec: ResolvedSpec): DiffResult {
  const errors: string[] = [];

  // A repeated column name is disambiguated by position -- "Value" pairs with
  // "Value", "Value (#2)" with "Value (#2)". That is sound whenever both files
  // lay their columns out the same way, and real reports repeat column groups
  // routinely: a currency block listing Name and Abbreviation a dozen times
  // over is a normal shape, not a defect. It only becomes a guess when the two
  // layouts disagree, because then the Nth "Value" need not be the same column.
  const canon = (h: string) => canonHeader(h, spec.looseHeaders);
  const sameLayout =
    base.headers.length === next.headers.length &&
    base.headers.every((h, i) => canon(h) === canon(next.headers[i] ?? ''));

  // --- structural trust checks -------------------------------------------
  for (const [label, m] of [['baseline', base], ['actual', next]] as const) {
    if (m.duplicateHeaders.length && !sameLayout)
      errors.push(`${label}: duplicate header(s) ${m.duplicateHeaders.map((h) => `"${h}"`).join(', ')} — the two files order their columns differently, so the repeated names could not be matched up reliably`);
    // A repeated key used to mean those rows were dropped, which was worth
    // failing over. They are now numbered and compared in order of appearance,
    // so it is worth saying and not worth failing: a breakdown carries one
    // "Total" row per group and every one of them has the same blank key.
    for (const k of spec.keyColumns)
      if (!m.headers.some((h) => canonHeader(h, spec.looseHeaders) === canonHeader(k, spec.looseHeaders)))
        errors.push(`${label}: key column "${k}" not found. Headers: ${m.headers.join(', ')}`);
  }
  if (spec.requireCachedValues && next.uncachedFormulaCells.length) {
    const n = next.uncachedFormulaCells.length;
    errors.push(
      `actual: ${n} formula cell(s) have no cached value (e.g. ${next.uncachedFormulaCells.slice(0, 3).join(', ')}). ` +
      `Value comparison silently passes on these. Either the generator must write cached results, or set requireCachedValues:false and rely on formula comparison.`,
    );
  }
  const drift = dialectDrift(base, next);
  if (drift.length) {
    const msg = `CSV dialect: ${drift.join('; ')}`;
    if (spec.csv.strictDialect) errors.push(msg);
  }

  // --- layer 1: schema ----------------------------------------------------
  const { shared, added, removed } = pairColumns(base, next, spec);
  const ignore = new Set(spec.ignoreColumns.map((c) => canonHeader(c, spec.looseHeaders)));
  const compared = shared.filter((p) => !ignore.has(canonHeader(p.base, spec.looseHeaders)));

  const moved: MovedColumn[] = shared
    .map((p) => ({
      column: p.next,
      from: base.headers.indexOf(p.base) + 1,
      to: next.headers.indexOf(p.next) + 1,
    }))
    .filter((m) => m.from !== m.to);

  // --- layer 2: row population -------------------------------------------
  // ignoreRows drops rows by key, the row-wise counterpart of ignoreColumns.
  // A key-value block holds its per-run values in rows, so a timestamp there
  // cannot be excluded by column.
  // Matched loosely, the way header names are: whether the sheet writes
  // "Report ID", "Report Id " or "REPORT ID:" is a styling decision, and a
  // config that has to guess which is a config that silently stops working.
  const ignoredRow = rowKeyMatcher(spec.ignoreRows);

  const addedRows = next.order.filter((k) => !base.rows.has(k) && !ignoredRow(k));
  const removedRows = base.order.filter((k) => !next.rows.has(k) && !ignoredRow(k));
  const sharedRows = base.order.filter((k) => next.rows.has(k) && !ignoredRow(k));

  // --- layers 3 & 4: values, types, formulas ------------------------------
  // Report metadata named by address. Compared to nothing, reported by layer 2
  // instead, which lists it with both values under "not verified".
  const metadataCells = new Set(spec.metadataCells.map((a) => a.toUpperCase()));
  const isMetadata = (c: Cell | undefined) =>
    metadataCells.size > 0 && !!c?.address
    && metadataCells.has(c.address.replace(/\$/g, '').toUpperCase());

  const values: ValueDiff[] = [];
  const types: TypeDiff[] = [];
  const formulas: FormulaDiff[] = [];

  for (const key of sharedRows) {
    const b = base.rows.get(key)!;
    const n = next.rows.get(key)!;
    const changedInRow = new Set<string>();
    const pending: ValueDiff[] = [];

    for (const pair of compared) {
      const bc = b[pair.base];
      const nc = n[pair.next];
      if (!bc || !nc) continue;
      if (isMetadata(bc) || isMetadata(nc)) continue;

      let formulaChanged = false;
      if (spec.compareFormulas && (bc.formula || nc.formula)) {
        const bf = normalizedFormula(bc, spec.formulaMode);
        const nf = normalizedFormula(nc, spec.formulaMode);
        if (bf !== nf) {
          formulaChanged = true;
          formulas.push({
            key, column: pair.next, address: nc.address,
            base: bf, next: nf, baseA1: bc.formula, nextA1: nc.formula,
          });
        }
      }

      const tol = toleranceFor(spec, pair.next);
      const rel = relativeToleranceFor(spec, pair.next);
      if (!equalValues(bc.value, nc.value, tol, rel)) {
        changedInRow.add(pair.next);
        const delta =
          typeof bc.value === 'number' && typeof nc.value === 'number'
            ? Math.abs(bc.value - nc.value)
            : undefined;
        pending.push({
          key, column: pair.next, address: nc.address,
          base: bc.value, next: nc.value, delta,
          formulaChanged, rootCause: false,
        });
      } else if (bc.kind !== nc.kind && bc.kind !== 'empty' && nc.kind !== 'empty') {
        // Equal once rendered, but a different type. "100" vs 100 is a real
        // defect that string comparison alone would wave through.
        types.push({
          key, column: pair.next, address: nc.address,
          baseKind: bc.kind, nextKind: nc.kind, value: nc.value,
        });
      }
    }

    // A change is a root cause when nothing it depends on also changed.
    for (const d of pending) {
      const deps = dependencies(n[d.column]!);
      d.rootCause = d.formulaChanged || !deps.some((dep) => changedInRow.has(dep));
      values.push(d);
    }
  }

  // --- invariants ---------------------------------------------------------
  const invariants: InvariantFailure[] = [];
  for (const inv of spec.invariants) {
    try {
      const fails = inv.check(next);
      if (fails) invariants.push(...fails);
    } catch (e) {
      invariants.push({ invariant: inv.name, detail: `threw: ${(e as Error).message}` });
    }
  }

  const defects = values.length + types.length + formulas.length + invariants.length + errors.length;
  const schemaChanged = added.length + removed.length + moved.length + addedRows.length + removedRows.length;

  return {
    ok: defects === 0 && (!spec.strictSchema || schemaChanged === 0),
    reviewOnly: defects === 0 && schemaChanged > 0,
    base: { source: base.source, sheet: base.sheet, rows: base.rows.size, columns: base.headers.length },
    next: { source: next.source, sheet: next.sheet, rows: next.rows.size, columns: next.headers.length },
    schema: { added, removed, moved, compared: compared.map((p) => p.next) },
    rows: {
      added: addedRows, removed: removedRows, compared: sharedRows.length,
      duplicateKeysBase: base.duplicateKeys, duplicateKeysNext: next.duplicateKeys,
    },
    values, types, formulas, invariants, errors,
  };
}
