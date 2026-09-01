/**
 * Verifies downloaded templates against the request and the screen that
 * produced them.
 *
 *   node template_comparison/cli.mjs [folder]
 *
 * Run on demand, over a tree the automation filled, the same way the report
 * comparison is run. It reads only; results are written beside each case.
 *
 * The tree, with a folder per template kind and a folder per case:
 *
 *   comparison/templates/
 *     program_selection_template/
 *       <case>/
 *         template/template.xlsx      what the app produced
 *         data/payload_data.json      what the client sent
 *         data/table_data.json        what the screen showed
 *         results/report.md           written by this run
 *
 * A template kind is named by its folder, and the folder must match a
 * descriptor under ./templates/. A folder with no descriptor is reported rather
 * than skipped: a case nobody is checking looks exactly like a case that passed.
 */
import { readdir, mkdir, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { join, relative } from 'node:path';
import { compareCase } from './compare.mjs';
import { report } from './report.mjs';

const ROOT = process.argv[2] ?? 'comparison/templates';
const HERE = new URL('.', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1');

const dirs = async (p) =>
  (await readdir(p, { withFileTypes: true })).filter((e) => e.isDirectory()).map((e) => e.name);

if (!existsSync(ROOT)) {
  console.error(`template-verify: no folder at ${ROOT}`);
  process.exit(1);
}

let cases = 0;
let failing = 0;
const unknown = [];

for (const kind of await dirs(ROOT)) {
  const descriptorPath = join(HERE, 'templates', kind, 'index.mjs');
  if (!existsSync(descriptorPath)) {
    unknown.push(kind);
    continue;
  }
  const descriptor = (await import(`./templates/${kind}/index.mjs`)).default;

  for (const name of await dirs(join(ROOT, kind))) {
    const caseDir = join(ROOT, kind, name);
    if (!existsSync(join(caseDir, 'template'))) continue;
    cases++;

    let outcome;
    try {
      outcome = await compareCase(caseDir, descriptor);
    } catch (e) {
      failing++;
      console.log(`x ${kind} · ${name}`);
      console.log(`    ${e.message}`);
      continue;
    }

    const { lines, markdown } = report(kind, name, outcome);
    if (!outcome.ok) failing++;
    for (const line of lines) console.log(line);

    await mkdir(join(caseDir, 'results'), { recursive: true });
    await writeFile(join(caseDir, 'results', 'report.md'), markdown, 'utf8');
    console.log(`    ${relative('.', join(caseDir, 'results', 'report.md')).replace(/\\/g, '/')}`);
  }
}

if (unknown.length) {
  console.log('');
  console.log(`${unknown.length} template folder(s) with no descriptor under templates/:`);
  for (const u of unknown) console.log(`  ${u}`);
}

console.log('');
console.log(`${cases} case(s), ${failing} failing`);
process.exit(failing ? 1 : 0);
