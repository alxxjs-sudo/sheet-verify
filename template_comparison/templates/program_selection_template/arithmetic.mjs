/**
 * The ROLePlay columns that can be checked without a source, because they are
 * arithmetic on columns beside them.
 *
 * No capture carries these figures. The download request does not have them,
 * the table on screen does not show them, and the only endpoint in the flow --
 * report-creation-treaties-template -- answers with a job ticket
 * (`{topic, downloadName, status, id}`), not with data. The server computes
 * them and writes them straight into the workbook, so there is nothing to
 * compare them against.
 *
 * Except each other. These are standard reinsurance identities, and every one
 * of them holds exactly on all three captured ROLePlay templates -- 17 rows,
 * worst relative gap 3.93e-16, which is IEEE754 rounding and nothing else.
 *
 * Two kinds, and the difference is the whole point of separating them:
 *
 *   ANCHORED -- every input traces back to a column verified against the
 *   payload, so the result is verified too. `Treaty Premium`, `Treaty
 *   Occurrence Limit` and `Treaty GCMP Premium` are payload-checked, which is
 *   what does the anchoring.
 *
 *   CONSISTENCY ONLY -- the rule ties several unverified figures together
 *   without pinning any of them down. Three equations in four unknowns says
 *   nothing about whether `Expected Loss` is right; it says that if it is
 *   wrong, `LOL`, `Loss Ratio` and `CV` are wrong in exactly the matching way.
 *   Worth having -- it catches one figure moving without the others -- but it
 *   is not verification, and reporting it as such would be the overstatement
 *   this tool exists to avoid.
 *
 * All are `optional`: a template downloaded without ROLePlay data has none of
 * these columns, and a rule that cannot run is skipped WITHOUT its column being
 * counted as covered.
 */

const div = (a, b) => (Number(b) === 0 ? null : Number(a) / Number(b));

export const ARITHMETIC = [
  // --- anchored to payload-verified columns -------------------------------
  {
    column: 'Limit',
    from: ['Treaty Occurrence Limit'],
    value: (limit) => limit,
    optional: true,
  },
  {
    column: 'Modeled Deposit Premium @ 100% Placed',
    from: ['Treaty Premium'],
    value: (premium) => premium,
    optional: true,
  },
  {
    // Rate on line: what the layer costs per unit of cover.
    column: 'Modeled ROL',
    from: ['Treaty Premium', 'Treaty Occurrence Limit'],
    value: (premium, limit) => div(premium, limit),
    optional: true,
  },

  // --- consistency only ----------------------------------------------------
  {
    // Loss on line.
    column: 'LOL',
    from: ['Expected Loss @ 100% Placed', 'Treaty Occurrence Limit'],
    value: (loss, limit) => div(loss, limit),
    anchored: false,
    optional: true,
  },
  {
    column: 'Loss Ratio',
    from: ['Expected Loss @ 100% Placed', 'Treaty Premium'],
    value: (loss, premium) => div(loss, premium),
    anchored: false,
    optional: true,
  },
  {
    // Coefficient of variation.
    column: 'CV',
    from: ['Standard Deviation @ 100% Placed', 'Expected Loss @ 100% Placed'],
    value: (sd, loss) => div(sd, loss),
    anchored: false,
    optional: true,
  },
  {
    column: 'Modeled Expected Premium @ 100% Placed',
    from: ['Modeled Deposit Premium @ 100% Placed', 'Modeled Reinstatement Premium @ 100% Placed'],
    value: (deposit, reinstatement) => Number(deposit) + Number(reinstatement),
    anchored: false,
    optional: true,
  },
  {
    column: 'GCMP Deposit Premium @ 100% Placed',
    from: ['GCMP USD Rate', 'Treaty GCMP Premium'],
    value: (rate, premium) => Number(rate) * Number(premium),
    anchored: false,
    optional: true,
  },
];

/** The columns these rules speak for, so they stop being declared unverifiable. */
export const COVERED_BY_ARITHMETIC = ARITHMETIC.map((r) => r.column);
