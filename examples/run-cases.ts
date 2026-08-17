/**
 * Runs every example case and prints what each one produced.
 *
 *   npm run build && npm run example
 *
 * Each case leaves a folder behind holding everything about it:
 *
 *   examples/cases/<name>/
 *     golden.xlsx   the output the new report is judged against
 *     actual.xlsx   the new report, copied in
 *     diff.txt            human-readable summary
 *     diff.json           the same, structured
 *     differences.xlsx    one row per differing cell
 *     compared.xlsx       every cell checked, one worksheet per table
 *
 * In real use `golden.xlsx` is committed and reviewed in the pull request;
 * here it is generated so the examples are self-contained.
 */
import { rm } from 'node:fs/promises';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
// Built output, the way a consumer imports it. Run `npm run build` first.
import { runCase } from '../dist/index.js';
import { generate } from './generate-report.ts';
import { CASES, SPEC } from './cases.ts';

const ROOT = join('examples', 'cases');
const OUT = join('examples', 'out');

const pad = (s: string, n: number) => s.padEnd(n);

let failures = 0;

for (const c of CASES) {
  const dir = join(ROOT, c.name);
  // A file open in Excel cannot be unlinked on Windows, which is not worth
  // failing over -- everything below overwrites in place.
  await rm(dir, { recursive: true, force: true }).catch(() => {});

  const spec = c.spec ?? SPEC;
  const golden = await generate(join(OUT, `${c.name}-golden.xlsx`), c.golden);
  const actual = await generate(join(OUT, `${c.name}-actual.xlsx`), c.actual);

  // First run establishes the golden output; the second is the real comparison.
  await runCase(golden, dir, spec);
  const result = await runCase(actual, dir, spec);

  const verdict = result.ok ? 'PASS' : 'FAIL';
  const asExpected = (result.ok ? 'pass' : 'fail') === c.expect;
  if (!asExpected) failures++;

  console.log(`\n${pad(c.name, 24)} ${verdict}${asExpected ? '' : '  ← UNEXPECTED'}`);
  console.log(`${pad('', 24)} ${c.about}`);
  console.log(`${pad('', 24)} ${result.summary}`);

  // Tally the ledger so the shape of each outcome is visible at a glance.
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(result.files.differences);
  const ws = wb.getWorksheet("Differences")!;
  const tally = new Map<string, number>();
  for (let r = 2; r <= ws.rowCount; r++) {
    const status = String(ws.getRow(r).getCell(5).value ?? '');
    tally.set(status, (tally.get(status) ?? 0) + 1);
  }
  const cells = [...tally].sort((a, b) => b[1] - a[1])
    .map(([s, n]) => `${n} ${s}`).join(', ');
  console.log(`${pad('', 24)} differences.xlsx: ${cells || 'no differing cells'}`);
}

console.log(`\n${CASES.length} cases run, all in ${ROOT}`);
if (failures) {
  console.error(`${failures} case(s) did not behave as documented`);
  process.exitCode = 1;
}
