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
import { join, relative, resolve, sep } from 'node:path';

const DEFAULT_ROOT = 'output_comparison';
const RESULT_DIR = 'results';
/** Summaries: one at the tree root, and one inside each report type's folder. */
const SUMMARY_DIR = '!summary';


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
    if (!e.isDirectory() || RESULT_DIRS.has(e.name)) continue;
    // A report type's own summary sits in its folder. Taken whole and not
    // descended into: there are no cases inside it, and it goes stale exactly
    // as loudly as a results/ folder does.
    if (e.name === SUMMARY_DIR) { yield join(dir, e.name); continue; }
    yield* resultsUnder(join(dir, e.name));
  }
}

// A Set: the walk above already reaches the summary folder at the root, and
// the explicit pass below names it too. Removing it twice would count it twice.
const seen = new Set();
for await (const p of resultsUnder(root)) seen.add(p);

/**
 * The run summary belongs to the whole run rather than to any case, so it lives
 * in a results folder at the root -- where the walk above, which only counts a
 * results folder with a golden beside it, never reaches it. A summary left from
 * the last run reads as current just as loudly as a stale results/ folder does.
 */
for (const name of [...RESULT_DIRS, SUMMARY_DIR]) {
  const p = join(root, name);
  if (existsSync(p)) seen.add(p);
}

const found = [...seen].sort();

if (!found.length) {
  console.log(`nothing to clear under ${relative(process.cwd(), root) || root}`);
  process.exit(0);
}

/**
 * What a clean actually recovers. Worth measuring rather than guessing: a
 * results folder is mostly compared.xlsx, which grows with the report, and on
 * a tree of forty cases the answer is gigabytes rather than the megabytes
 * people expect.
 */
async function bytesUnder(dir) {
  let entries;
  try { entries = await readdir(dir, { withFileTypes: true }); } catch { return 0; }
  let total = 0;
  for (const e of entries) {
    const p = join(dir, e.name);
    total += e.isDirectory() ? await bytesUnder(p) : await stat(p).then((s) => s.size, () => 0);
  }
  return total;
}

function size(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  const units = ['KB', 'MB', 'GB'];
  let n = bytes / 1024;
  let i = 0;
  while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
  return `${n < 10 ? n.toFixed(1) : Math.round(n)} ${units[i]}`;
}

/**
 * The report type a folder belongs to. Forty lines reading "removed
 * <path>/results" said nothing forty times; grouped, the same clean is a dozen
 * lines showing at a glance which types were cleared and which were not.
 *
 * A results folder sits inside a case, which sits inside its report type -- so
 * the type is two levels up, at whatever depth the tree happens to nest. The
 * summary folder and a results folder at the very root belong to no type.
 */
function groupOf(p) {
  const parts = relative(root, p).split(sep);
  const type = RESULT_DIRS.has(parts[parts.length - 1]) ? parts.slice(0, -2) : parts;
  return type.join('/') || '(root)';
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
const groups = new Map();
let freed = 0;
let removed = 0;

for (const p of found) {
  const shown = relative(process.cwd(), p);
  const bytes = await bytesUnder(p);
  if (dry) {
    console.log(`would remove  ${shown}  (${size(bytes)})`);
    freed += bytes;
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
    const group = groups.get(groupOf(p)) ?? { folders: 0, bytes: 0 };
    group.folders++;
    group.bytes += bytes;
    groups.set(groupOf(p), group);
    freed += bytes;
    removed++;
  } catch (e) {
    // Keep going. One locked folder should cost that folder, not the rest.
    failures.push({ shown, why: e.code ?? e.message });
  }
}

if (groups.size) {
  const width = Math.max(...[...groups.keys()].map((k) => k.length));
  for (const [name, g] of [...groups].sort((a, b) => a[0].localeCompare(b[0]))) {
    console.log(
      `  ${name.padEnd(width)}  ${String(g.folders).padStart(3)}` +
      ` folder${g.folders === 1 ? ' ' : 's'}  ${size(g.bytes).padStart(8)}`,
    );
  }
}

console.log(
  `\n${removed} item(s)${dry ? ' would be removed' : ' removed'}, ${size(freed)}` +
  `${dry ? ' — re-run without --dry to do it' : ' freed'}`,
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
