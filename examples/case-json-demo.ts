/**
 * When you need a case.json, and what goes in it.
 *
 *   npm run build && npm run example:case-json
 *
 * Builds a report containing a Summary sheet whose rows are identified by
 * three columns -- Region + Band + Quarter. Detection tries single columns and
 * then pairs, so it finds no key there and reports the sheet as NOT COMPARED
 * rather than guessing one. A defect is planted on that sheet, so you can see
 * it stay hidden and then be found.
 *
 * The real CLI is invoked as a subprocess at each step, so what you see here
 * is exactly what you would get running it yourself.
 */
import { mkdir, rm, writeFile, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { detectSpec } from '../dist/index.js';
import { generate } from './generate-report.ts';

const ROOT = join('examples', 'cli-cases');
const CASE = join(ROOT, 'summary-needs-a-key');

const rule = (n: number, title: string) => {
  console.log(`\n${'─'.repeat(72)}\n${n}. ${title}\n${'─'.repeat(72)}`);
};

/** Runs the CLI the way you would, and shows what it printed. */
function cli(...args: string[]): string {
  const r = spawnSync('node', ['dist/cli.js', ...args], { encoding: 'utf8' });
  const out = `${r.stdout ?? ''}${r.stderr ?? ''}`.trimEnd();
  console.log(out);
  console.log(`  [exit code ${r.status}]`);
  return out;
}

await rm(ROOT, { recursive: true, force: true }).catch(() => {});
await mkdir(CASE, { recursive: true });

rule(1, 'Two files in a case folder, as usual');

await generate(join(CASE, 'golden.xlsx'), { summarySheet: true });
await generate(join(CASE, 'actual.xlsx'), { summarySheet: true, summaryDefect: true });
console.log(`  ${CASE}/golden.xlsx`);
console.log(`  ${CASE}/actual.xlsx`);
console.log('\n  A premium on the Summary sheet differs between them: 88000 -> 91500.');

rule(2, 'Run it — and the Summary sheet is skipped, not compared');

cli(CASE);

console.log(`
  Read that carefully. The run PASSED, and it is not a clean bill of health:
  "1 sheet not compared" means detection could not identify a row on Summary,
  so it refused to compare it rather than pair rows arbitrarily. The planted
  defect is sitting there unreported.`);

rule(3, 'Ask what it detected');

console.log('  node dist/cli.js --print-spec ' + CASE + '\n');
console.log('  It prints every sheet; here is the one that matters:\n');

const detected = await detectSpec(join(CASE, 'golden.xlsx'));
console.log(
  JSON.stringify({ Summary: detected.sheets!['Summary'] }, null, 2).replace(/^/gm, '    '),
);
console.log(`
  Table 1 is the info block and was keyed by "Field". Table 2 is the data, and
  it has a headerRow but NO keyColumns -- every other table in the file has one.
  That is the gap.`);

rule(4, 'Write a case.json naming the key');

const caseJson = {
  sheets: {
    Summary: { keyColumns: ['Region', 'Band', 'Quarter'] },
  },
};
await writeFile(join(CASE, 'case.json'), JSON.stringify(caseJson, null, 2) + '\n', 'utf8');
console.log(`  ${join(CASE, 'case.json')}\n`);
console.log((await readFile(join(CASE, 'case.json'), 'utf8')).replace(/^/gm, '  '));
console.log(`  Only the part that was wrong -- three lines, not the whole spec.

  Note it names the key on the SHEET rather than on Table 2. A per-table value
  wins over a sheet-level one, so Table 1 keeps the "Field" key detection gave
  it, and only the table that had none picks this up. Naming the table
  explicitly works too:

    { "sheets": { "Summary": { "tables": { "Table 2": { "keyColumns": [...] } } } } }`);

rule(5, 'Run it again — the defect surfaces');

cli(CASE);

const report = await readFile(join(CASE, 'result', 'diff.txt'), 'utf8');
console.log(`
  ${(/^ +\d+ tables compared.*$/m.exec(report)?.[0] ?? '').trim()} — one more than before, and the run now FAILS
  on the defect that was invisible in step 2:

${(/^ {4}\S.*(?:88000|91500).*$/m.exec(report)?.[0] ?? '').replace(/^ {4}/, '    ')}

  The lesson: "not compared" is work to do, never a pass. Check the
  tables-compared count against what your report actually contains.

  Everything is left in ${CASE} to look through.`);
