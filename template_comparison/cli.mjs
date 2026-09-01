/**
 * Verifies downloaded templates against the request and the screen that
 * produced them.
 *
 *   npm run compare:templates              verify ./comparison/templates
 *   npm run compare:templates -- <folder>  verify somewhere else
 *   npm run bless:templates                take the current downloads as the contract
 *   npm run clean:templates                clear the results folders
 *
 * <folder> is the templates root, holding a folder per kind. Naming one kind's
 * folder works too -- it is the obvious thing to reach for when only one of them
 * is being worked on, and reading it as a root would report every case in it as
 * a kind nobody has a descriptor for.
 *
 * Run on demand, over a tree the automation filled, the same way the report
 * comparison is run. The two inputs are read and never written to; results are
 * written beside each case.
 *
 * The tree, with a folder per template kind and a folder per case:
 *
 *   comparison/templates/
 *     program_selection_template/
 *       <case>/
 *         current/template.xlsx       what the app produced
 *         golden/template.xlsx        what it produced when this was blessed
 *         data/payload_data.json      what the client sent
 *         data/table_data.json        what the screen showed
 *         results/report.md           written by this run
 *
 * A template kind is named by its folder, and the folder must match a
 * descriptor under ./templates/. A folder with no descriptor is reported rather
 * than skipped: a case nobody is checking looks exactly like a case that passed.
 */
import { readdir, mkdir, writeFile, rm, copyFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { basename, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { compareCase, folders, openTemplate } from './compare.mjs';
import { report } from './report.mjs';

const args = process.argv.slice(2);
const CLEAN = args.includes('--clean');
const BLESS = args.includes('--bless');
const ROOT = args.find((a) => !a.startsWith('-')) ?? 'comparison/templates';
// fileURLToPath and not `.pathname`, which is percent-encoded: a checkout under
// a folder named "OneDrive - MMC" gave a HERE of "OneDrive%20-%20MMC", so no
// descriptor was ever found and every kind was reported as one nobody wrote.
const HERE = fileURLToPath(new URL('.', import.meta.url));

const dirs = async (p) =>
  (await readdir(p, { withFileTypes: true })).filter((e) => e.isDirectory()).map((e) => e.name);

if (!existsSync(ROOT)) {
  console.error(`template-verify: no folder at ${ROOT}`);
  process.exit(1);
}

const descriptorFor = (kind) => join(HERE, 'templates', kind, 'index.mjs');

// The root holds a folder per kind. Naming one kind's folder directly is the
// obvious thing to reach for when only one of them is being worked on, so take
// that as well -- read as a root it would report every case inside it as a kind
// nobody wrote a descriptor for, which looks like a missing descriptor and is
// really a path one level too deep.
const named = basename(resolve(ROOT));
const kinds = existsSync(descriptorFor(named))
  ? [{ kind: named, dir: ROOT }]
  : (await dirs(ROOT)).map((kind) => ({ kind, dir: join(ROOT, kind) }));

let cases = 0;
let failing = 0;
const unknown = [];

for (const { kind, dir } of kinds) {
  if (!existsSync(descriptorFor(kind))) {
    unknown.push(kind);
    continue;
  }
  const descriptor = (await import(`./templates/${kind}/index.mjs`)).default;

  for (const name of await dirs(dir)) {
    const caseDir = join(dir, name);
    const where = folders(caseDir);
    if (!where) continue;
    cases++;

    // Take the current download as the new contract. Separate from a run so it
    // is always deliberate: blessing a drift you have not read is how a wrong
    // figure becomes the thing everything else is measured against.
    if (BLESS) {
      if (!where.golden) {
        console.log(`- ${kind} · ${name}: nothing to bless, this case has no golden/`);
        continue;
      }
      const { file } = await openTemplate(where.current);
      await mkdir(where.golden, { recursive: true });
      await copyFile(join(where.current, file), join(where.golden, file));
      console.log(`  blessed ${relative('.', join(where.golden, file)).replace(/\\/g, '/')}`);
      continue;
    }

    // A run overwrites what it produces, so this is not needed for a correct
    // comparison -- it is needed for an honest one. A case that stops failing
    // leaves its old report sitting there, and a case that is renamed leaves a
    // whole results folder behind. Both read as current.
    if (CLEAN) {
      const results = join(caseDir, 'results');
      if (existsSync(results)) {
        await rm(results, { recursive: true, force: true });
        console.log(`  removed ${relative('.', results).replace(/\\/g, '/')}`);
      }
      continue;
    }

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
  console.log(`  (expected one of: ${(await dirs(join(HERE, 'templates'))).join(', ')})`);
}

console.log('');
if (BLESS) {
  console.log(`${cases} case(s) blessed`);
  process.exit(0);
}
if (CLEAN) {
  console.log(`${cases} case(s) cleared`);
  process.exit(0);
}
console.log(`${cases} case(s), ${failing} failing`);

// A run that recognised nothing printed "0 case(s), 0 failing" and exited 0,
// which is the same thing a clean run says. Anything unrecognised, and anything
// at all that ran nothing, is a failed run.
if (unknown.length || cases === 0) process.exit(1);
process.exit(failing ? 1 : 0);
