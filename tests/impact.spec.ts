import { test, expect } from '@playwright/test';
import { referencesOf } from '../src/a1.js';
import { impactOf, addressOf, type FormulaCell } from '../src/impact.js';

/**
 * Following a change outward to the cells that read it. The reports store no
 * calculated results, so a total whose inputs moved looks untouched -- its
 * formula is identical and there is no number to compare. This is what says
 * "and these will come out different".
 */

const rects = (formula: string) =>
  referencesOf(formula).map((r) => `${r.sheet ?? ''}${r.fromCol},${r.fromRow}:${r.toCol},${r.toRow}`);

test.describe('referencesOf', () => {
  test('reads a plain reference and a range', () => {
    expect(rects('A1')).toEqual(['1,1:1,1']);
    expect(rects('SUM(B2:B5)')).toEqual(['2,2:2,5']);
  });

  test('keeps the whole range, not just its corners', () => {
    // A change in the middle of a SUM has to reach the SUM.
    const [r] = referencesOf('SUM($AO$29:$AO$38)');
    expect(r!.fromRow).toBe(29);
    expect(r!.toRow).toBe(38);
  });

  test('carries the sheet name, quoted or not', () => {
    expect(rects("SUM('Report Info'!$AO$29:$AO$38)")).toEqual(['Report Info41,29:41,38']);
    expect(rects('Summary!C7')).toEqual(['Summary3,7:3,7']);
  });

  test('is not fooled by text, function names or error values', () => {
    // LOG10 ends in something that reads like a reference; "A1" inside a
    // string is text; #REF! is not a cell.
    expect(rects('LOG10(B4)')).toEqual(['2,4:2,4']);
    expect(rects('IF(A1="B2","C3",D4)')).toEqual(['1,1:1,1', '4,4:4,4']);
    expect(rects('#REF!')).toEqual([]);
  });

  test('reads every branch of an IF, not only the one taken', () => {
    // Over-approximating is deliberate: calling a cell unaffected when it is
    // not is the expensive mistake.
    expect(rects('IF(A1>0,B1,C1)')).toEqual(['1,1:1,1', '2,1:2,1', '3,1:3,1']);
  });

  test('handles the real shape from these reports', () => {
    const f = 'IF(OR(D14<=0,NOT(ISNUMBER(D14)),NOT(ISNUMBER(D15))),"-",(D15-D14)/D14)';
    expect(new Set(rects(f))).toEqual(new Set(['4,14:4,14', '4,15:4,15']));
  });
});

test.describe('impactOf', () => {
  const cell = (sheet: string, row: number, col: number, formula: string): FormulaCell =>
    ({ sheet, row, col, formula });

  test('a changed input reaches the total that reads it', () => {
    const formulas = [cell('Totals', 16, 4, '(D15-D14)/D14')];

    const r = impactOf(formulas, [{ sheet: 'Totals', row: 15, col: 4 }]);

    expect(r.affected).toHaveLength(1);
    expect(addressOf(r.affected[0]!)).toBe('D16');
  });

  test('and onward through a chain of formulas', () => {
    const formulas = [
      cell('S', 2, 1, 'A1*2'),      // reads the changed cell
      cell('S', 3, 1, 'A2*2'),      // reads the one above
      cell('S', 4, 1, 'A3*2'),
    ];

    const r = impactOf(formulas, [{ sheet: 'S', row: 1, col: 1 }]);

    expect(r.affected.map(addressOf).sort()).toEqual(['A2', 'A3', 'A4']);
    // Two of the three were reached only through another formula.
    expect(r.indirect).toBe(2);
  });

  test('a change nothing reads moves nothing', () => {
    const formulas = [cell('S', 2, 1, 'A1*2')];

    const r = impactOf(formulas, [{ sheet: 'S', row: 9, col: 9 }]);

    expect(r.affected).toHaveLength(0);
    expect(r.inert).toBe(1);
  });

  test('follows a reference across sheets', () => {
    const formulas = [cell('Summary', 5, 2, "SUM('Report Info'!$A$1:$A$10)")];

    const r = impactOf(formulas, [{ sheet: 'Report Info', row: 4, col: 1 }]);

    expect(r.affected.map(addressOf)).toEqual(['B5']);
  });

  test('does not confuse sheets that merely share a cell address', () => {
    const formulas = [cell('Summary', 5, 2, 'A4*2')];

    // Same address, different sheet: must not match.
    const r = impactOf(formulas, [{ sheet: 'Report Info', row: 4, col: 1 }]);

    expect(r.affected).toHaveLength(0);
  });

  test('a whole-column reference is followed without expanding a million cells', () => {
    const formulas = [cell('S', 1, 5, 'SUM(A1:A1048576)')];

    const r = impactOf(formulas, [{ sheet: 'S', row: 900000, col: 1 }]);

    expect(r.affected.map(addressOf)).toEqual(['E1']);
  });

  test('a cycle terminates instead of looping', () => {
    const formulas = [cell('S', 1, 1, 'A2'), cell('S', 2, 1, 'A1')];

    const r = impactOf(formulas, [{ sheet: 'S', row: 1, col: 1 }]);

    // A2 reads the changed cell. A1 is the changed cell itself, already
    // reported as a difference, so it is not listed again here.
    expect(r.affected.map(addressOf)).toEqual(['A2']);
  });

  test('a cell already reported as differing is not listed twice', () => {
    // The formula in B1 changed *and* it reads A1, which also changed. It is
    // in the differences already; repeating it here would double-count.
    const formulas = [cell('S', 1, 2, 'A1*2'), cell('S', 1, 3, 'B1*2')];

    const r = impactOf(formulas, [
      { sheet: 'S', row: 1, col: 1 },
      { sheet: 'S', row: 1, col: 2 },
    ]);

    expect(r.affected.map(addressOf)).toEqual(['C1']);
  });
});
