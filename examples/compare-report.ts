/**
 * Compares the two generated releases.
 *
 *   npm run build
 *   node --experimental-strip-types examples/generate-report.ts examples/out/baseline.xlsx
 *   node --experimental-strip-types examples/generate-report.ts examples/out/actual.xlsx --next
 *   node --experimental-strip-types examples/compare-report.ts
 *
 * Every sheet holds two tables: an output-info block on rows 1-6 and the data
 * table from row 8 down. Each is declared separately, so the info block stops
 * where the data begins instead of swallowing it.
 */
// Built output, the way a consumer imports it. Run `npm run build` first.
import {
  verifyWorkbook, formatWorkbookReport, summarizeWorkbook, invariants as inv,
} from '../dist/index.js';

const BASE = 'examples/out/baseline.xlsx';
const NEXT = 'examples/out/actual.xlsx';

/** Every sheet carries the same info block, so it lives in the defaults. */
const info = {
  headerRow: 1,
  keyColumns: ['Field'],
  // Written afresh on every run: a difference here is never a defect.
  ignoreRows: ['Generated At'],
};

const d = await verifyWorkbook(BASE, NEXT, {
  defaults: {
    headerRow: 8,
    invariants: [inv.noErrorValues()],
  },
  sheets: {
    Policies: {
      tables: { Info: info, Detail: { keyColumns: ['PolicyId'] } },
    },
    Premiums: {
      tables: { Info: info, Detail: { keyColumns: ['PolicyId', 'Period'] } },
    },
    Claims: {
      tables: { Info: info, Detail: { keyColumns: ['ClaimId'] } },
    },
    Commissions: {
      tables: { Info: info, Detail: { keyColumns: ['AgentId'] } },
    },
    Regions: {
      tables: { Info: info, Detail: { keyColumns: ['Region'] } },
    },
  },
});

console.log('\n' + summarizeWorkbook(d) + '\n');
console.log(formatWorkbookReport(d));
