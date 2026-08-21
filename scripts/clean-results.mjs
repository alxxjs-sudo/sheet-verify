#!/usr/bin/env node
/**
 * Deletes every `results/` folder under the cases root.
 *
 * A run overwrites the artefacts it produces, so this is not needed to get a
 * correct comparison. It is needed to get an honest one: a case that stops
 * failing leaves its old differences.xlsx sitting there, and a case that is
 * renamed or removed leaves a whole results folder behind. Both read as current.
 * Clearing first means everything present was produced by the run you just did.
 *
 * Only directories named exactly `results` are removed, and only ones that sit
 * beside a golden file -- the two inputs and any config are never touched, and
 * neither is a folder called `results` somewhere else in the tree.
 *
 *   npm run clean            clear ./output_comparison
 *   npm run clean -- --dry   list what would go, delete nothing
 *   npm run clean -- path/to/cases
 */
import { readdir, rm, stat } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';

const DEFAULT_ROOT = 'output_comparison';
const RESULT_DIR = 'results';


const args = process.argv.slice(2);
const dry = args.some((a) => a === '--dry' || a === '--dry-run');
const root = resolve(args.find((a) => !a.startsWith('-')) ?? DEFAULT_ROOT);
/**
 * A run started with `--results <name>` writes somewhere else, and those
 * folders go just as stale as the default one. Naming it here means a clean
 * clears what the last run actually wrote rather than what it usually writes.
 */
const named = args.find((a) => a.startsWith('--results='));
const RESULT_DIRS = new Set([RESULT_DIR, ...(named ? [named.slice(10)] : [])]);

const isDir = async (p) => stat(p).then((s) => s.isDirectory(), () => false);

if (!(await isDir(root))) {
  console.error(`sheet-verify: no such folder\n  ${root}`);
  process.exit(1);
}

/**
 * A results folder counts only when a golden sits beside it -- either a
 * golden file, or a `golden/` folder holding one, which is how a download
 * keeps the name the source system gave it. Anything else called `results`
 * belongs to whoever put it there.
 */
const GOLDEN_FILE = /^(golden|baseline|expected|before)\./i;
const GOLDEN_DIR = /^(golden|baseline|expected|before)$/i;

async function* resultsUnder(dir) {
  let entries;
  try { entries = await readdir(dir, { withFileTypes: true }); } catch { return; }

  const names = entries.map((e) => e.name);
  const beside = entries.some(
    (e) => (e.isFile() && GOLDEN_FILE.test(e.name)) || (e.isDirectory() && GOLDEN_DIR.test(e.name)),
  );
  if (beside) {
    for (const name of RESULT_DIRS) if (names.includes(name)) yield join(dir, name);
  }
  for (const e of entries) {
    if (e.isDirectory() && !RESULT_DIRS.has(e.name)) yield* resultsUnder(join(dir, e.name));
  }
}

const found = [];
for await (const p of resultsUnder(root)) found.push(p);

/**
 * The run summary belongs to the whole run rather than to any case, so it lives
 * in a results folder at the root -- where the walk above, which only counts a
 * results folder with a golden beside it, never reaches it. A summary left from
 * the last run reads as current just as loudly as a stale results/ folder does.
 */
for (const name of [...RESULT_DIRS, '!summary']) {
  const p = join(root, name);
  if (existsSync(p)) found.push(p);
}

if (!found.length) {
  console.log(`nothing to clear under ${relative(process.cwd(), root) || root}`);
  process.exit(0);
}

/**
 * Delete first, report second, and never stop at the first refusal.
 *
 * This printed "removed" *before* attempting the delete, so a folder Windows
 * refused to release was reported as gone. Worse, the throw took the whole
 * script with it, leaving every later folder untouched and unmentioned -- a
 * clean that says it cleared the tree and cleared half of it is worse than one
 * that fails outright, because the run afterwards looks trustworthy.
 *
 * Files here get locked in ordinary use: Excel holds a workbook the moment it
 * is open, and OneDrive holds one while it syncs.
 */
const failures = [];
let removed = 0;

for (const p of found) {
  const shown = relative(process.cwd(), p);
  if (dry) {
    console.log(`would remove  ${shown}`);
    removed++;
    continue;
  }
  try {
    // Retry, because the usual failure here is transient. Windows marks a file
    // delete-pending until the last handle closes, so the files inside go and
    // the rmdir that follows fails with EPERM -- which is exactly the shape of
    // the error reported from a real tree. Node retries EBUSY, EMFILE, ENFILE,
    // ENOTEMPTY and EPERM, which is the whole set that matters here.
    await rm(p, { recursive: true, force: true, maxRetries: 8, retryDelay: 150 });
    console.log(`removed  ${shown}`);
    removed++;
  } catch (e) {
    // Keep going. One locked folder should cost that folder, not the rest.
    failures.push({ shown, why: e.code ?? e.message });
  }
}

console.log(
  `\n${removed} item(s)${dry ? ' would be removed' : ' removed'}` +
  `${dry ? ' — re-run without --dry to do it' : ''}`,
);

if (failures.length) {
  console.error(`\n${failures.length} item(s) could NOT be removed:`);
  for (const f of failures) console.error(`  ${f.shown}  (${f.why})`);
  console.error(
    'Something is holding them open. Close any workbook from these folders in'
    + ' Excel, and if the tree is inside OneDrive, pause syncing or move the tree'
    + ' out of it: OneDrive keeps a handle while it uploads, and can restore what'
    + ' you delete.'
    + '\n\nThese folders still hold the PREVIOUS run. Delete them before trusting'
    + ' what the next run reports.',
  );
  process.exit(1);
}
