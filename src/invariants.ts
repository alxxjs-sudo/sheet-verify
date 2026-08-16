/**
 * Baseline-free assertions.
 *
 * A golden-file comparison proves the new output matches the old one, never
 * that either is correct: an error present in the baseline is invisible
 * forever. These run against the actual output alone and close that gap.
 */
import type { Invariant, InvariantFailure, Row } from './types.js';

const EXCEL_ERRORS = /^#(REF|DIV\/0|VALUE|NAME\?|N\/A|NULL|NUM|SPILL|CALC|GETTING_DATA)!?$/i;

/** Flags #REF!, #DIV/0!, #VALUE! and friends anywhere in the sheet. */
export function noErrorValues(): Invariant {
  return {
    name: 'no-error-values',
    check(m) {
      const fails: InvariantFailure[] = [];
      for (const [key, row] of m.rows) {
        for (const [column, cell] of Object.entries(row)) {
          const v = cell.value;
          if (cell.kind === 'error' || (typeof v === 'string' && EXCEL_ERRORS.test(v.trim()))) {
            fails.push({
              invariant: 'no-error-values', key, column, address: cell.address,
              detail: `cell contains ${String(v)}`,
            });
          }
        }
      }
      return fails;
    },
  };
}

/** Requires the named columns to be non-empty on every row. */
export function notBlank(...columns: string[]): Invariant {
  return {
    name: 'not-blank',
    check(m) {
      const fails: InvariantFailure[] = [];
      for (const [key, row] of m.rows) {
        for (const column of columns) {
          const cell = row[column];
          if (!cell) continue;
          if (cell.value === null || cell.value === '') {
            fails.push({ invariant: 'not-blank', key, column, address: cell.address, detail: 'is empty' });
          }
        }
      }
      return fails;
    },
  };
}

/** Numeric bounds, inclusive. Non-numeric values are reported too. */
export function inRange(column: string, min: number, max: number): Invariant {
  return {
    name: `in-range(${column})`,
    check(m) {
      const fails: InvariantFailure[] = [];
      for (const [key, row] of m.rows) {
        const cell = row[column];
        if (!cell || cell.value === null) continue;
        const v = cell.value;
        if (typeof v !== 'number') {
          fails.push({ invariant: `in-range(${column})`, key, column, address: cell.address, detail: `expected a number, got ${JSON.stringify(v)}` });
        } else if (v < min || v > max) {
          fails.push({ invariant: `in-range(${column})`, key, column, address: cell.address, detail: `${v} outside [${min}, ${max}]` });
        }
      }
      return fails;
    },
  };
}

/** Requires values in a column to be distinct across the sheet. */
export function unique(column: string): Invariant {
  return {
    name: `unique(${column})`,
    check(m) {
      const seen = new Map<string, string>();
      const fails: InvariantFailure[] = [];
      for (const [key, row] of m.rows) {
        const cell = row[column];
        if (!cell || cell.value === null) continue;
        const v = String(cell.value);
        const first = seen.get(v);
        if (first !== undefined) {
          fails.push({ invariant: `unique(${column})`, key, column, address: cell.address, detail: `duplicate of row ${first}` });
        } else seen.set(v, key);
      }
      return fails;
    },
  };
}

/**
 * Recomputes a column from the others and checks the sheet agrees.
 *
 * This is the invariant that actually catches calculation bugs a baseline
 * shares, because it asserts the arithmetic independently of the generator.
 */
export function derived(
  column: string,
  compute: (row: Readonly<Row>, num: (col: string) => number) => number | null,
  tolerance = 1e-9,
): Invariant {
  return {
    name: `derived(${column})`,
    check(m) {
      const fails: InvariantFailure[] = [];
      const num = (row: Row) => (col: string) => {
        const v = row[col]?.value;
        return typeof v === 'number' ? v : Number.NaN;
      };
      for (const [key, row] of m.rows) {
        const cell = row[column];
        if (!cell || cell.value === null) continue;
        let expected: number | null;
        try {
          expected = compute(row, num(row));
        } catch (e) {
          fails.push({ invariant: `derived(${column})`, key, column, address: cell.address, detail: `compute threw: ${(e as Error).message}` });
          continue;
        }
        if (expected === null || Number.isNaN(expected)) continue;
        const actual = cell.value;
        if (typeof actual !== 'number') {
          fails.push({ invariant: `derived(${column})`, key, column, address: cell.address, detail: `expected a number, got ${JSON.stringify(actual)}` });
        } else if (Math.abs(actual - expected) > tolerance) {
          fails.push({
            invariant: `derived(${column})`, key, column, address: cell.address,
            detail: `sheet says ${actual}, recomputed ${expected} (Δ ${Math.abs(actual - expected)})`,
          });
        }
      }
      return fails;
    },
  };
}

/** Every formula cell must carry a cached result, else values are untested. */
export function formulasHaveResults(): Invariant {
  return {
    name: 'formulas-have-results',
    check(m) {
      return m.uncachedFormulaCells.map((address) => ({
        invariant: 'formulas-have-results', address,
        detail: 'formula cell has no cached value',
      }));
    },
  };
}
