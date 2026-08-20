import { numToCol, referencesOf, type RefRect } from './a1.js';

/**
 * What will come out different once Excel recalculates.
 *
 * The reports carry no calculated results, so a formula whose inputs moved
 * looks untouched: its text is the same on both sides and nothing is stored to
 * compare. The comparison reports the input that moved and stops there, which
 * is correct but leaves the obvious question unanswered -- the total two rows
 * down reads off that input, so what happens to it?
 *
 * This answers that without evaluating anything. A formula's result can only
 * move if its own text changed or something it reads changed, so following the
 * references outward from every cell that differs reaches every cell that will
 * recalculate differently. No arithmetic is done and none is needed.
 *
 * Three honest limits. It says *will differ*, not *by how much* -- there is no
 * number to report. It over-approximates: `IF(A1>0,B1,C1)` counts as reading
 * all three, so a change in the branch not taken still marks the cell. Naming a
 * cell that turns out to match is the cheaper mistake.
 *
 * And it can *under*-approximate, which is the expensive one. References are
 * found by reading the formula's text, so a function that computes which cell
 * it reads defeats it: `OFFSET($A$1, MATCH(...), 0)` textually references
 * `$A$1` while actually reading wherever the arithmetic lands, and `INDIRECT`
 * is worse. A change reaching a cell that way is not reported here. This is not
 * hypothetical -- 62,406 of the 379,959 formulas in one real tree are OFFSET.
 * Recalculating the files instead of reasoning about them has no such gap, and
 * is what `recalc.ts` is for.
 */

export interface CellId {
  sheet: string;
  row: number;
  col: number;
}

/** A formula cell, as the graph needs it. */
export interface FormulaCell extends CellId {
  formula: string;
}

/**
 * A formula cell that will recalculate, and the cell it reads that makes it do
 * so.
 *
 * Naming the cause is what makes the finding usable. "This cell recalculates"
 * asks the reader to go looking; "this cell reads Report Info!AT59, which went
 * from 35000000 to Unlimited" is the whole story in one row, which is what a
 * reader who opened the two files and found a difference the tool appeared to
 * miss actually needs.
 */
export interface AffectedRef extends CellId {
  /**
   * The cell whose change reaches this one. A differing cell when `indirect`
   * is false; another formula cell, itself recalculating, when it is true.
   */
  via: CellId;
  /** True when the cell it reads is itself a formula rather than a difference. */
  indirect: boolean;
}

export interface ImpactResult {
  /** Formula cells that read something which differs, directly or through others. */
  affected: AffectedRef[];
  /** How many of those were reached only through another formula. */
  indirect: number;
  /** Cells that differ and are read by nothing -- the changes with no consequence. */
  inert: number;
}

/** Excel forbids two sheets differing only in case. */
const canon = (s: string) => s.trim().toLowerCase();
const idOf = (c: CellId) => `${canon(c.sheet)}|${c.row}|${c.col}`;

export const addressOf = (c: CellId) => `${numToCol(c.col)}${c.row}`;

/**
 * Rectangles wider than this are not expanded into individual cells. A whole
 * column reference would otherwise put a million entries in the index; there
 * are few of them, so they are checked by scanning instead.
 */
const EXPAND_LIMIT = 4096;

const area = (r: RefRect) => (r.toRow - r.fromRow + 1) * (r.toCol - r.fromCol + 1);

const covers = (r: RefRect, sheet: string, row: number, col: number) =>
  canon(r.sheet ?? sheet) === canon(sheet) &&
  row >= r.fromRow && row <= r.toRow && col >= r.fromCol && col <= r.toCol;

/**
 * Follows references outward from the cells that differ, and returns the
 * formula cells they reach.
 */
export function impactOf(formulas: FormulaCell[], changed: CellId[]): ImpactResult {
  // cell -> formulas that read it, for the references small enough to expand.
  const readers = new Map<string, FormulaCell[]>();
  // Formulas whose references were too wide to expand, checked by scanning.
  const wide: { cell: FormulaCell; rects: RefRect[] }[] = [];

  for (const f of formulas) {
    let rects: RefRect[];
    try { rects = referencesOf(f.formula); } catch { continue; }
    const big: RefRect[] = [];

    for (const r of rects) {
      if (area(r) > EXPAND_LIMIT) { big.push(r); continue; }
      const sheet = r.sheet ?? f.sheet;
      for (let row = r.fromRow; row <= r.toRow; row++) {
        for (let col = r.fromCol; col <= r.toCol; col++) {
          const key = idOf({ sheet, row, col });
          const list = readers.get(key);
          if (list) list.push(f);
          else readers.set(key, [f]);
        }
      }
    }
    if (big.length) wide.push({ cell: f, rects: big });
  }

  const seeds = new Set(changed.map(idOf));
  const dirty = new Set(seeds);
  const affected: AffectedRef[] = [];
  const affectedIds = new Set<string>();
  let indirect = 0;
  const reachedBySeed = new Set<string>();

  const queue: CellId[] = [...changed];
  while (queue.length) {
    const cell = queue.pop()!;
    const direct = seeds.has(idOf(cell));

    const hits = readers.get(idOf(cell)) ?? [];
    // A wide reference is rare, so scanning them per popped cell is cheaper
    // than putting a million keys in the index.
    for (const w of wide) {
      if (w.rects.some((r) => covers(r, cell.sheet, cell.row, cell.col))) hits.push(w.cell);
    }

    for (const f of hits) {
      const id = idOf(f);
      if (dirty.has(id)) continue;
      dirty.add(id);
      affectedIds.add(id);
      affected.push({
        sheet: f.sheet,
        row: f.row,
        col: f.col,
        via: { sheet: cell.sheet, row: cell.row, col: cell.col },
        indirect: !direct,
      });
      if (direct) reachedBySeed.add(id);
      else indirect++;
      queue.push(f);
    }
  }

  // A change nothing reads has no consequence beyond itself -- worth counting,
  // because it says the rest of the workbook is unaffected by it.
  const inert = changed.filter((c) => {
    if ((readers.get(idOf(c)) ?? []).length) return false;
    return !wide.some((w) => w.rects.some((r) => covers(r, c.sheet, c.row, c.col)));
  }).length;

  return { affected, indirect, inert };
}
