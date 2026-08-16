import { test, expect } from '@playwright/test';
import { toR1C1, toHeaderRef, colToNum, numToCol } from '../src/a1.js';

const headers = new Map<number, string>([
  [1, 'PolicyId'], [2, 'Holder'], [3, 'Region'], [4, 'Sum Insured'],
  [5, 'Premium'], [6, 'Rate'], [7, 'Annual Cost'], [8, 'Commission'],
]);

test.describe('column letters', () => {
  test('round-trips through the 26-boundary and up to XFD', () => {
    for (const [n, s] of [[1, 'A'], [26, 'Z'], [27, 'AA'], [52, 'AZ'], [702, 'ZZ'], [703, 'AAA'], [16384, 'XFD']] as const) {
      expect(numToCol(n)).toBe(s);
      expect(colToNum(s)).toBe(n);
    }
  });
});

test.describe('R1C1', () => {
  test('relative refs become offsets from the host cell', () => {
    // G2 = D2*F2  ->  three left, one left
    expect(toR1C1('D2*F2', 2, 7)).toBe('RC[-3]*RC[-1]');
  });

  test('same column or row collapses to bare R / C', () => {
    expect(toR1C1('G1', 2, 7)).toBe('R[-1]C');
    expect(toR1C1('A2', 2, 7)).toBe('RC[-6]');
  });

  test('absolute refs keep their absolute position', () => {
    expect(toR1C1('$D$2', 5, 7)).toBe('R2C4');
    // Mixed: absolute row, relative column.
    expect(toR1C1('D$2', 5, 7)).toBe('R2C[-3]');
    // Mixed the other way: relative row (2 is three above 5), absolute column.
    expect(toR1C1('$D2', 5, 7)).toBe('R[-3]C4');
  });
});

test.describe('header resolution', () => {
  test('same-row refs become header names', () => {
    expect(toHeaderRef('D2*F2', 2, 7, headers)).toBe('[Sum Insured]@row*[Rate]@row');
  });

  test('the shifted and unshifted layouts normalise identically', () => {
    const before = toHeaderRef('D2*E2', 2, 6, new Map([[4, 'Sum Insured'], [5, 'Rate'], [6, 'Annual Cost']]));
    const after = toHeaderRef('D2*F2', 2, 7, headers);
    expect(after).toBe(before);
  });

  test('ranges resolve at both ends', () => {
    expect(toHeaderRef('SUM(D2:D7)', 8, 4, headers)).toBe('SUM([Sum Insured]@row-6:[Sum Insured]@row-1)');
  });

  test('off-row refs carry a signed offset', () => {
    expect(toHeaderRef('D1', 2, 7, headers)).toBe('[Sum Insured]@row-1');
    expect(toHeaderRef('D3', 2, 7, headers)).toBe('[Sum Insured]@row+1');
  });
});

test.describe('scanner does not corrupt non-references', () => {
  // Each of these breaks a naive /[A-Z]+[0-9]+/ substitution.
  test('function names containing digits survive', () => {
    expect(toHeaderRef('LOG10(D2)', 2, 7, headers)).toBe('LOG10([Sum Insured]@row)');
    expect(toR1C1('LOG10(D2)', 2, 7)).toBe('LOG10(RC[-3])');
  });

  test('string literals are untouched', () => {
    expect(toHeaderRef('IF(D2>0,"see A1 for detail","")', 2, 7, headers))
      .toBe('IF([Sum Insured]@row>0,"see A1 for detail","")');
  });

  test('doubled quotes inside literals do not end the string early', () => {
    expect(toHeaderRef('CONCAT("a""B2""c",D2)', 2, 7, headers))
      .toBe('CONCAT("a""B2""c",[Sum Insured]@row)');
  });

  test('sheet-qualified refs stay in A1 form', () => {
    expect(toHeaderRef('Rates!B2*D2', 2, 7, headers)).toBe('Rates!B2*[Sum Insured]@row');
    expect(toHeaderRef("'Rate Table'!B2*D2", 2, 7, headers)).toBe("'Rate Table'!B2*[Sum Insured]@row");
  });

  test('structured table references are left alone', () => {
    expect(toHeaderRef('SUM(Table1[Amount])', 2, 7, headers)).toBe('SUM(Table1[Amount])');
  });

  test('error literals are preserved', () => {
    expect(toHeaderRef('IFERROR(D2,#N/A)', 2, 7, headers)).toBe('IFERROR([Sum Insured]@row,#N/A)');
  });

  test('boolean and defined names are not references', () => {
    expect(toHeaderRef('IF(TRUE,TaxRate,D2)', 2, 7, headers)).toBe('IF(TRUE,TaxRate,[Sum Insured]@row)');
  });

  test('out-of-range coordinates are treated as names', () => {
    expect(toHeaderRef('ZZZZ1', 2, 7, headers)).toBe('ZZZZ1');
    expect(toHeaderRef('A99999999', 2, 7, headers)).toBe('A99999999');
  });

  test('columns beyond the header row fall back to a stable placeholder', () => {
    expect(toHeaderRef('Z2', 2, 7, headers)).toBe('[col26]@row');
  });
});
