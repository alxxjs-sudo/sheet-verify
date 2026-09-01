/**
 * Projects the treaties-template request payload into the rows the template
 * should hold, so the two can be compared as spreadsheets rather than as
 * JSON-against-cells.
 *
 * Row identity is the treaty id (template column L). Payload order is NOT row
 * order -- measured on the captured pair, 36 of 68 rows line up positionally
 * and the rest have shuffled -- so nothing here may depend on position.
 */

/**
 * Value of a `{ value, currency }` money field, or a plain number.
 *
 * A limitless layer arrives as the STRING "Infinity" -- JSON has no literal for
 * the number, so the API sends the word -- and the template writes it as
 * "Unlimited". That is not cosmetic: the column's validation rule is
 * `OR(Q4="Unlimited", ISNUMBER(Q4))`, so "Unlimited" is a value Excel accepts
 * there by name and anything else spelling it would be rejected on upload.
 *
 * Seen on `aggregateLimit` once and `perRiskLimit` three times in the captured
 * payload, so it is a shape to handle rather than a one-off.
 */
const money = (v) => {
  const n = v && typeof v === 'object' ? v.value : v;
  return n === 'Infinity' || n === Infinity ? 'Unlimited' : n;
};

/** Payload holds percentages; the sheet holds fractions. */
const pct = (v) => (v == null ? null : v / 100);

const yesNo = (v) => (v === true ? 'Yes' : v === false ? 'No' : null);

/** Treaty Type is an enum in the payload and a label in the sheet. */
const TREATY_TYPE = { 0: 'Occurrence' };

/**
 * Participating Limit, which the template computes rather than receives: the
 * occurrence limit times the share taken of it. Verified against all 68 rows
 * of the captured pair.
 */
function participatingLimit(t, e) {
  const limit = money(t.effectiveLimit);
  const share = e.participation;
  if (typeof limit !== 'number' || typeof share !== 'number') return null;
  return limit * (share / 100);
}

/**
 * One row per treaty, keyed by treaty id, in the template's own column names.
 * Only the left block (B..AF) is produced: everything from AG rightwards is
 * ROLePlay reference data the payload does not carry.
 */
export function projectRows(payload) {
  const rows = new Map();

  for (const p of payload.participations) {
    for (const t of p.treaties) {
      // The editable values live on `t.participation`, not on the treaty. In an
      // untouched template the two agree on every field; they diverge the
      // moment somebody edits one, which is the whole point of the round trip.
      const e = t.participation ?? {};

      rows.set(String(t.id), {
        'Include': yesNo(e.included),
        'View Of Risk Id': p.id,
        'Company Name': p.companyName,
        'Edison Program Name': p.edisonProgramName,
        'View Of Risk': p.viewOfRiskType?.name ?? null,
        'Effective Date': p.effectiveDate,
        'Office Regions': p.officeRegions,
        'Currency': p.currency,
        'Approved By': p.approvedBy,
        'Created At': p.createdAt,
        'Treaty Id': t.id,
        'Treaty Name': t.name,
        'Treaty Type': TREATY_TYPE[t.type] ?? null,
        'Treaty Occurrence Limit': money(t.effectiveLimit),
        'Treaty Occurrence Retention': money(t.effectiveRetention),
        'Treaty Aggregate Limit': money(t.aggregateLimit),
        'Treaty Aggregate Retention': money(t.aggregateRetention),
        'Treaty Number of Reinstatements': t.numberOfReinstatements,
        'Treaty Reinstatement Description': t.gcmpReinstatementDescription,
        'Treaty Premium': money(e.premium ?? t.premium),
        'Treaty GCMP Premium': money(t.gcmpPremium),
        'Treaty Non Concurrent GCMP Premium': money(t.gcmpNonConcurrentPremium),
        'Treaty Placement': pct(t.placement),
        'Treaty GCMP Placement': pct(t.gcmpPlacement),
        'Treaty Non Concurrent GCMP Placement': pct(t.gcmpNonConcurrentPlacement),
        'Treaty Participation': pct(e.participation),
        'Treaty Computed Participation': pct(e.computedParticipation),
        'Treaty Cap Limit': money(e.capLimit),
        'Treaty Discount': pct(e.discount),
        'Treaty Tags': e.tags?.length ? e.tags.join(', ') : null,
        // The one column the payload does not carry: the template works it out.
        // Occurrence limit x participation, which holds on all 68 rows of the
        // captured pair. Projecting it rather than leaving it out means the
        // comparison checks the template's own arithmetic and not merely that
        // it copied the request across.
        //
        // Null when the limit is not a number -- a limitless layer reads
        // "Unlimited", and a share of unlimited is not a figure to invent.
        'Treaty Participating Limit': participatingLimit(t, e),
      });
    }
  }
  return rows;
}

/** The columns above, in the order the template writes them. */
export const COLUMNS = [
  'Include', 'View Of Risk Id', 'Company Name', 'Edison Program Name',
  'View Of Risk', 'Effective Date', 'Office Regions', 'Currency', 'Approved By',
  'Created At', 'Treaty Id', 'Treaty Name', 'Treaty Type',
  'Treaty Occurrence Limit', 'Treaty Occurrence Retention',
  'Treaty Aggregate Limit', 'Treaty Aggregate Retention',
  'Treaty Number of Reinstatements', 'Treaty Reinstatement Description',
  'Treaty Premium', 'Treaty GCMP Premium', 'Treaty Non Concurrent GCMP Premium',
  'Treaty Placement', 'Treaty GCMP Placement',
  'Treaty Non Concurrent GCMP Placement', 'Treaty Participation',
  'Treaty Computed Participation', 'Treaty Cap Limit', 'Treaty Discount',
  'Treaty Tags', 'Treaty Participating Limit',
];
