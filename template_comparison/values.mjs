/**
 * Reading a template cell and a captured value as the same thing.
 *
 * The two sources do not speak the same dialect as the sheet. A request body
 * carries raw JSON; a screen carries whatever the browser rendered, which is
 * abbreviated, whitespace-collapsed and sometimes missing its currency. None of
 * that is a defect, and none of it should be reported as one -- but neither
 * should it be waved through, because "close enough" applied loosely is how a
 * real difference gets absorbed.
 *
 * Each rule here is written to be as tight as the source allows and no tighter.
 */

/** The multiplier each abbreviation stands for. */
const UNIT = { K: 1e3, M: 1e6, B: 1e9, T: 1e12 };

const MONEY = /^\s*(?:([A-Za-z]{3})\s+)?(-?[\d,]*\.?\d+)\s*([KMBT])?\s*$/;

/**
 * Reads "USD 18.00M", "BRL 85000000", "4.86M" or "1,234.5" into a currency, a
 * number, and the half-step its own precision implies.
 *
 * The half-step is the point. "4.86M" is not 4,860,000 -- it is every value
 * that rounds to 4.86M, a band 5,000 wide, and the underlying figure was
 * 4,862,069. Reading it as an exact number would report a correct template as
 * wrong; ignoring the precision entirely would let 4,900,000 through. The band
 * is exactly what the screen actually said.
 *
 * Returns null for anything that is not a number.
 */
export function readMoney(raw) {
  if (raw == null) return null;
  if (typeof raw === 'number') return { currency: null, value: raw, half: 0 };
  const m = MONEY.exec(String(raw));
  if (!m) return null;
  const [, currency, digits, suffix] = m;
  const unit = suffix ? UNIT[suffix.toUpperCase()] : 1;
  const plain = digits.replace(/,/g, '');
  const decimals = plain.includes('.') ? plain.split('.')[1].length : 0;
  return {
    currency: currency ? currency.toUpperCase() : null,
    value: Number(plain) * unit,
    // A figure written to n decimals of its unit stands for a band half a step
    // wide either side. Written in full, the band is nothing.
    half: (suffix || decimals ? 0.5 * 10 ** -decimals * unit : 0),
  };
}

/**
 * A comparator for a money column, for use as a column's `compare`.
 *
 * Handles both conventions these templates use: the whole figure in one cell
 * ("BRL 85000000"), and the amount in one cell with its currency in a sibling
 * column. Pass that sibling's name to check the currency as well; without it
 * the currency simply is not checked, which is the honest outcome when the
 * source never stated one.
 */
export function money(currencyColumn) {
  return (got, want, cell) => {
    const blankGot = got == null || got === '';
    const blankWant = want == null || want === '';
    if (blankGot && blankWant) return true;
    if (blankGot !== blankWant) return false;

    const a = readMoney(got);
    const b = readMoney(want);
    if (!a || !b) return false;

    // The currency lives in the cell for some columns and in a sibling for
    // others. Only compared when both sides actually carry one: an override the
    // screen shows as "18.00M" states no currency, and inventing one to compare
    // against would be checking our own guess.
    const mine = a.currency ?? (currencyColumn ? String(cell(currencyColumn) ?? '').trim().toUpperCase() || null : null);
    if (mine && b.currency && mine !== b.currency) return false;

    return Math.abs(a.value - b.value) <= a.half + b.half;
  };
}

/**
 * Runs of whitespace read as one.
 *
 * A browser collapses whitespace before it paints, so a template cell holding
 * "Layer 2 -  85M XS 35M" and a screen showing "Layer 2 - 85M XS 35M" are the
 * same observation -- the screen had no way to show the difference. Applied
 * only to sources that were read off a rendered page; a request body carries
 * the string exactly and is compared exactly.
 */
export const collapse = (v) => (typeof v === 'string' ? v.replace(/\s+/g, ' ').trim() : v);
