/**
 * Runs the five-sheet report as a case.
 *
 *   npm run build && npm run example
 *
 * The case folder ends up holding everything about this one report type:
 *
 *   examples/cases/monthly-policy-export/
 *     golden.xlsx   the output the new report is judged against
 *     actual.xlsx   the new report, copied in
 *     diff.txt      the differences, human-readable
 *     diff.json     the same, structured
 *     cells.csv     every cell compared, and its verdict
 *
 * In real use `golden.xlsx` is committed and reviewed in the pull request;
 * here it is generated so the example is self-contained.
 */
import { rm, readFile } from 'node:fs/promises';
import { join } from 'node:path';
// Built output, the way a consumer imports it. Run `npm run build` first.
import { runCase, invariants as inv } from '../dist/index.js';
import { generate } from './generate-report.ts';

const CASE_DIR = join('examples', 'cases', 'monthly-policy-export');

/** Every sheet carries the same info block, so it lives in one constant. */
const info = {
  headerRow: 1,
  keyColumns: ['Field'],
  // Rewritten on every run: a difference here is never a defect.
  ignoreRows: ['Generated At'],
};

const spec = {
  defaults: { headerRow: 8, invariants: [inv.noErrorValues()] },
  sheets: {
    Policies: { tables: { Info: info, Detail: { keyColumns: ['PolicyId'] } } },
    Premiums: { tables: { Info: info, Detail: { keyColumns: ['PolicyId', 'Period'] } } },
    Claims: { tables: { Info: info, Detail: { keyColumns: ['ClaimId'] } } },
    Commissions: { tables: { Info: info, Detail: { keyColumns: ['AgentId'] } } },
    Regions: { tables: { Info: info, Detail: { keyColumns: ['Region'] } } },
  },
};

// Start from an empty case folder so the example is repeatable.
await rm(CASE_DIR, { recursive: true, force: true });

// First run: no golden output yet, so the release under test establishes it.
const golden = await generate(join('examples', 'out', 'release-4.2.0.xlsx'), false);
const first = await runCase(golden, CASE_DIR, spec);
console.log(`first run  → ${first.summary}\n`);

// Second run: the following release, judged against what the first blessed.
const actual = await generate(join('examples', 'out', 'release-4.3.0.xlsx'), true);
const result = await runCase(actual, CASE_DIR, spec);

console.log(`case       ${result.name}`);
console.log(`verdict    ${result.ok ? 'PASS' : 'FAIL'} — ${result.summary}\n`);
console.log(await readFile(result.files.diffText, 'utf8'));

const cells = (await readFile(result.files.cells, 'utf8')).trim().split('\n');
const [header, ...body] = cells;
const status = (line: string) => line.split(',')[4];
const tally = new Map<string, number>();
for (const line of body) tally.set(status(line)!, (tally.get(status(line)!) ?? 0) + 1);

console.log(`cells.csv  ${body.length} cells recorded`);
for (const [s, n] of [...tally].sort((a, b) => b[1] - a[1])) {
  console.log(`           ${String(n).padStart(5)}  ${s}`);
}
console.log(`\nfiles      ${Object.values(result.files).join('\n           ')}`);
