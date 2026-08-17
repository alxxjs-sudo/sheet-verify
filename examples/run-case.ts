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
 *     cells.xlsx    a formatted table, one row per differing cell
 *
 * In real use `golden.xlsx` is committed and reviewed in the pull request;
 * here it is generated so the example is self-contained.
 */
import { rm, readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
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

// Start from an empty case folder so the example is repeatable. A file open
// in Excel or an editor cannot be unlinked on Windows, which is not worth
// failing the run over -- everything written below overwrites in place.
await rm(CASE_DIR, { recursive: true, force: true }).catch(() => {});

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

// The ledger is a formatted worksheet, so read it back the same way any
// consumer would rather than parsing it as text.
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(result.files.cells);
const ws = wb.getWorksheet('Cells')!;

const tally = new Map<string, number>();
for (let r = 2; r <= ws.rowCount; r++) {
  const status = String(ws.getRow(r).getCell(5).value ?? '');
  tally.set(status, (tally.get(status) ?? 0) + 1);
}

console.log(`cells.xlsx  ${ws.rowCount - 1} differing cell(s) recorded`);
for (const [s, n] of [...tally].sort((a, b) => b[1] - a[1])) {
  console.log(`            ${String(n).padStart(4)}  ${s}`);
}
// compared.xlsx keeps every cell that was checked, split a tab per table.
const all = new ExcelJS.Workbook();
await all.xlsx.readFile(result.files.compared);
const total = all.worksheets.reduce((n, ws) => n + ws.rowCount - 1, 0);

console.log(`\ncompared.xlsx  ${total} cells checked across ${all.worksheets.length} tabs`);
for (const ws of all.worksheets) {
  console.log(`            ${String(ws.rowCount - 1).padStart(4)}  ${ws.name}`);
}

console.log(`\nfiles       ${Object.values(result.files).join('\n            ')}`);
