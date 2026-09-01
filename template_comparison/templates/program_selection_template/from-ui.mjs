/**
 * Projects the UI table capture into the rows the template should hold.
 *
 * The sibling of project.mjs, which does the same from the request payload. Both
 * emit the SAME row shape, so one writer and one comparison serve both: the only
 * thing that differs is where the numbers came from.
 *
 * That matters more than it looks. The payload is what the client sent; the UI
 * capture is what a person could see. Checking the template against both means a
 * bug that corrupts the request is caught by one and a bug that misrenders the
 * screen is caught by the other -- where a single source would call either one
 * "correct" because it never had a second opinion.
 *
 * Row identity is Edison Program Name + Treaty Name. Not the treaty id, which
 * the template has and the UI does not display; and not company + treaty name,
 * which collides -- one company can carry the same treaty name in several
 * programs (12 of 68 rows in the first capture).
 */

/** The UI states percentages as percentages; the sheet holds fractions. */
const pct = (v) => (v == null ? null : v / 100);

const yesNo = (v) => (v === true ? 'Yes' : v === false ? 'No' : null);

export const KEY_COLUMNS = ['Edison Program Name', 'Treaty Name'];

/**
 * @param table the parsed table_data.json
 * @returns Map keyed by "<program name>|<treaty name>"
 */
export function projectTableRows(table) {
  const rows = new Map();

  for (const p of table.programs ?? []) {
    for (const t of p.treaties ?? []) {
      rows.set(`${p.edisonProgramName}|${t.name}`, {
        'Include': yesNo(t.include),
        'Company Name': p.companyName,
        'Edison Program Name': p.edisonProgramName,
        'View Of Risk': p.viewOfRisk,
        'Effective Date': p.effectiveDate,
        'Office Regions': p.officeRegions,
        'Currency': p.currency,
        'Approved By': p.approvedBy,
        'Treaty Name': t.name,
        'Treaty Type': t.type,
        'Treaty Occurrence Limit': t.occurrenceLimit,
        'Treaty Occurrence Retention': t.occurrenceRetention,
        'Treaty Aggregate Limit': t.aggregateLimit,
        'Treaty Aggregate Retention': t.aggregateRetention,
        'Treaty Number of Reinstatements': t.numberOfReinstatements,
        'Treaty Reinstatement Description': t.reinstatementDescription,
        'Treaty Premium': t.premium,
        'Treaty GCMP Premium': t.gcmpPremium,
        'Treaty Non Concurrent GCMP Premium': t.gcmpNonConcurrentPremium,
        'Treaty Placement': pct(t.placement),
        'Treaty GCMP Placement': pct(t.gcmpPlacement),
        'Treaty Non Concurrent GCMP Placement': pct(t.gcmpNonConcurrentPlacement),
        'Treaty Participation': pct(t.participation),
        'Treaty Computed Participation': pct(t.computedParticipation),
        'Treaty Cap Limit': t.capLimit,
        'Treaty Discount': pct(t.discount),
        'Treaty Tags': t.tags?.length ? t.tags.join(', ') : null,
      });
    }
  }
  return rows;
}

/**
 * The columns the UI can account for.
 *
 * Four of the template's thirty-one are missing, and each for a reason worth
 * knowing rather than working around:
 *
 *   View Of Risk Id            an internal id the table never shows
 *   Created At                 shown, but the capture does not carry it yet
 *   Treaty Id                  not displayed anywhere in the UI
 *   Treaty Participating Limit computed by the template, not shown on screen
 *
 * All four are covered by the payload comparison, so nothing goes unverified --
 * it is simply verified from the one source that can see it.
 */
export const COLUMNS = [
  'Include', 'Company Name', 'Edison Program Name', 'View Of Risk',
  'Effective Date', 'Office Regions', 'Currency', 'Approved By',
  'Treaty Name', 'Treaty Type',
  'Treaty Occurrence Limit', 'Treaty Occurrence Retention',
  'Treaty Aggregate Limit', 'Treaty Aggregate Retention',
  'Treaty Number of Reinstatements', 'Treaty Reinstatement Description',
  'Treaty Premium', 'Treaty GCMP Premium', 'Treaty Non Concurrent GCMP Premium',
  'Treaty Placement', 'Treaty GCMP Placement',
  'Treaty Non Concurrent GCMP Placement', 'Treaty Participation',
  'Treaty Computed Participation', 'Treaty Cap Limit', 'Treaty Discount',
  'Treaty Tags',
];
