/**
 * What the overrides download request can prove.
 *
 * Very little, and that is worth stating plainly rather than papering over. The
 * request body is:
 *
 *   { "targetType": "metaRiskTreaty", "targetIds": [96900, 96898, ...],
 *     "timezoneOffset": -180 }
 *
 * -- a list of ids and nothing else. Every value in the workbook is assembled
 * server-side from those ids, so the request cannot be checked against a single
 * cell. Verifying the values needs the RESPONSE captured, which nothing does
 * yet; the table capture stands in for it in the meantime.
 *
 * What the request CAN prove is membership, and membership is not a small
 * thing: it is the check that catches a download returning rows nobody asked
 * for, or quietly dropping one that was selected. That failure is invisible to
 * a value comparison, because every row it does return is correct.
 */

/** No columns: this source speaks to which rows exist, not to what is in them. */
export const COLUMNS = [];

/** The ids asked for, as the identity of the rows expected back. */
export const projectRows = (data) => new Map(
  (data.targetIds ?? []).map((id) => [String(id), {}]),
);

/** What the request says it was for, so a mismatched capture is caught. */
export const declares = (data) => data.targetType ?? null;
