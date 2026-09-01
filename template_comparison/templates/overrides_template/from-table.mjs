/**
 * The overrides table on screen, read against the template it produced.
 *
 * This capture is self-describing: it carries a `headers` map from its own
 * field names to the exact column titles the template writes.
 *
 *   { "headers": { "override-premium": "Edison Treaty Premium", ... },
 *     "rows": [ { "override-premium": "4.86M", ... } ] }
 *
 * So the column list is taken from the capture rather than written out here.
 * Transcribing it would only create a second place to be wrong, and a field the
 * capture stops carrying is caught by coverage -- which reports the template
 * column nobody checked, which is the failure that actually matters.
 *
 * Two adjustments are needed between what the screen shows and what the sheet
 * holds, and both are properties of the screen rather than defects:
 *
 *   - The editable columns are titled with a trailing " *" in the sheet and
 *     without it on screen, so a name is resolved to whichever the sheet has.
 *   - Money is abbreviated ("USD 4.86M" for 4,862,069) and sometimes shown
 *     without its currency. Those columns are compared through `money`, which
 *     reads the abbreviation as the band of values it actually stands for.
 */
import { money } from '../../values.mjs';

/** The sheet titles an editable column "X *"; the screen calls it "X". */
const resolve = (name, columns) => (columns.has(name) ? name : `${name} *`);

/**
 * Money columns, and where each keeps its currency.
 *
 * A value is either whole in one cell ("BRL 85000000") or split into an amount
 * and a sibling currency column. Naming the sibling lets the currency be
 * checked too; where the screen states no currency, it simply is not checked
 * rather than being compared against a guess.
 */
const MONEY = {
  'GCMP Layer Limits': null,
  'MetaRisk Treaty Occurrence Limit': 'MetaRisk Treaty Occurrence Limit Currency',
  'MetaRisk Treaty Aggregate Limit': 'MetaRisk Treaty Aggregate Limit Currency',
  'MetaRisk Treaty Premium': 'MetaRisk Treaty Premium Currency',
  'Edison Treaty Occurrence Limit': 'Edison Treaty Occurrence Limit Currency',
  'Edison Treaty Aggregate Limit': 'Edison Treaty Aggregate Limit Currency',
  'Edison Treaty Premium': 'Edison Treaty Premium Currency',
  'Edison Layer Occurrence Limit': 'Edison Layer Occurrence Limit Currency',
  'Edison Layer Aggregate Limit': 'Edison Layer Aggregate Limit Currency',
};

/** Every column this capture speaks for, named as the sheet names it. */
export const columnsFor = (data, columns) =>
  Object.values(data.headers ?? {}).map((title) => {
    const name = resolve(title, columns);
    return title in MONEY ? { name, compare: money(MONEY[title]) } : name;
  });

/** One row per captured row, keyed the way the descriptor's `key` reads it. */
export function projectRowsFor(keyField) {
  return (data, columns) => {
    const headers = data.headers ?? {};
    const out = new Map();
    for (const row of data.rows ?? []) {
      const want = {};
      for (const [field, title] of Object.entries(headers)) want[resolve(title, columns)] = row[field];
      out.set(String(row[keyField]), want);
    }
    return out;
  };
}

/** What the capture says it was taken of, so a mismatched pair is caught. */
export const declares = (data) => data.targetType ?? null;

/**
 * The page says how many rows it was showing. A capture taken mid-scroll would
 * be a partial source silently compared as a complete one.
 */
export const isComplete = (data) => !data.page || data.page.shown === data.page.total;
