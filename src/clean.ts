import { readdir, rm, stat } from 'node:fs/promises';
import { join, relative, sep } from 'node:path';

/**
 * Clears every `results/` folder under a cases tree, and every `!summary/`.
 *
 * A run overwrites the artefacts it produces, so this is not needed to get a
 * correct comparison. It is needed to get an honest one: a case that stops
 * failing leaves its old differences.xlsx sitting there, and a case that is
 * renamed or removed leaves a whole results folder behind. Both read as
 * current. Clearing first means everything present was produced by the run you
 * just did.
 *
 * Only directories named exactly `results` are removed, and only ones that sit
 * beside a golden file -- the two inputs and any config are never touched, and
 * neither is a folder called `results` somewhere else in the tree.
 *
 * This lives in `src/` rather than in `scripts/` because `scripts/` is not in
 * the published package. A tree kept in another repository is exactly where
 * stale results pile up, and a maintenance command you cannot run there is not
 * a maintenance command.
 */

/**
 * A results folder counts only when a golden sits beside it -- either a golden
 * file, or a `golden/` folder holding one, which is how a download keeps the
 * name the source system gave it. Anything else called `results` belongs to
 * whoever put it there.
 */
const GOLDEN_FILE = /^(golden|baseline|expected|before)\./i;
const GOLDEN_DIR = /^(golden|baseline|expected|before)$/i;

export interface CleanTarget {
  path: string;
  bytes: number;
  /** The report type it belongs to, for grouping the log. */
  group: string;
}

export interface CleanOutcome {
  targets: CleanTarget[];
  removed: number;
  freed: number;
  groups: { name: string; folders: number; bytes: number }[];
  failures: { path: string; why: string }[];
}

export interface CleanOptions {
  /** List what would go, delete nothing. */
  dry?: boolean;
  /**
   * The results folder this run writes into, from `--results`. Those go just as
   * stale as the default one, so a clean clears what the last run actually
   * wrote rather than what it usually writes.
   */
  resultDir?: string;
  /** The summary folder name, at the root and inside each report type. */
  summaryDir?: string;
}

/**
 * What a clean actually recovers. Worth measuring rather than guessing: a
 * results folder is mostly compared.xlsx, which grows with the report, and on
 * a tree of forty cases the answer is gigabytes rather than the megabytes
 * people expect.
 */
async function bytesUnder(dir: string): Promise<number> {
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return 0;
  }
  let total = 0;
  for (const e of entries) {
    const p = join(dir, e.name);
    total += e.isDirectory() ? await bytesUnder(p) : await stat(p).then((s) => s.size, () => 0);
  }
  return total;
}

export function formatSize(bytes: number): string {
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
function groupOf(root: string, p: string, resultDirs: Set<string>): string {
  const parts = relative(root, p).split(sep);
  const type = resultDirs.has(parts[parts.length - 1]!) ? parts.slice(0, -2) : parts;
  return type.join('/') || '(root)';
}

async function* resultsUnder(
  dir: string,
  resultDirs: Set<string>,
  summaryDir: string,
): AsyncGenerator<string> {
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  const names = entries.map((e) => e.name);
  const beside = entries.some(
    (e) => (e.isFile() && GOLDEN_FILE.test(e.name)) || (e.isDirectory() && GOLDEN_DIR.test(e.name)),
  );
  if (beside) {
    for (const name of resultDirs) if (names.includes(name)) yield join(dir, name);
  }
  for (const e of entries) {
    if (!e.isDirectory() || resultDirs.has(e.name)) continue;
    // A report type's own summary sits in its folder. Taken whole and not
    // descended into: there are no cases inside it, and it goes stale exactly
    // as loudly as a results/ folder does.
    if (e.name === summaryDir) { yield join(dir, e.name); continue; }
    yield* resultsUnder(join(dir, e.name), resultDirs, summaryDir);
  }
}

/** Everything a clean would remove, measured, in tree order. */
export async function findStale(root: string, options: CleanOptions = {}): Promise<CleanTarget[]> {
  const summaryDir = options.summaryDir ?? '!summary';
  const resultDirs = new Set(['results', ...(options.resultDir ? [options.resultDir] : [])]);

  // A Set: the walk below already reaches the summary folder at the root, and
  // the explicit pass after it names the same folder. Removing it twice would
  // count it twice.
  const seen = new Set<string>();
  for await (const p of resultsUnder(root, resultDirs, summaryDir)) seen.add(p);

  // The run summary belongs to the whole run rather than to any case, so it
  // lives at the root -- where the walk above, which only counts a results
  // folder with a golden beside it, never reaches it. A summary left from the
  // last run reads as current just as loudly as a stale results/ folder does.
  for (const name of [...resultDirs, summaryDir]) {
    const p = join(root, name);
    if (await stat(p).then(() => true, () => false)) seen.add(p);
  }

  const targets: CleanTarget[] = [];
  for (const path of [...seen].sort()) {
    targets.push({ path, bytes: await bytesUnder(path), group: groupOf(root, path, resultDirs) });
  }
  return targets;
}

/**
 * Delete first, report second, and never stop at the first refusal.
 *
 * This once printed "removed" *before* attempting the delete, so a folder
 * Windows refused to release was reported as gone. Worse, the throw took the
 * whole command with it, leaving every later folder untouched and unmentioned
 * -- a clean that says it cleared the tree and cleared half of it is worse than
 * one that fails outright, because the run afterwards looks trustworthy.
 *
 * Files here get locked in ordinary use: Excel holds a workbook the moment it
 * is open, and OneDrive holds one while it syncs.
 */
export async function clean(root: string, options: CleanOptions = {}): Promise<CleanOutcome> {
  const targets = await findStale(root, options);
  const outcome: CleanOutcome = {
    targets, removed: 0, freed: 0, groups: [], failures: [],
  };
  if (options.dry) {
    outcome.removed = targets.length;
    outcome.freed = targets.reduce((n, t) => n + t.bytes, 0);
    return outcome;
  }

  const byGroup = new Map<string, { folders: number; bytes: number }>();
  for (const t of targets) {
    try {
      // Retry, because the usual failure here is transient. Windows marks a
      // file delete-pending until the last handle closes, so the files inside
      // go and the rmdir that follows fails with EPERM -- exactly the shape of
      // the error reported from a real tree. Node retries EBUSY, EMFILE,
      // ENFILE, ENOTEMPTY and EPERM, which is the whole set that matters here.
      await rm(t.path, { recursive: true, force: true, maxRetries: 8, retryDelay: 150 });
      const g = byGroup.get(t.group) ?? { folders: 0, bytes: 0 };
      g.folders++;
      g.bytes += t.bytes;
      byGroup.set(t.group, g);
      outcome.removed++;
      outcome.freed += t.bytes;
    } catch (e) {
      // Keep going. One locked folder should cost that folder, not the rest.
      const err = e as NodeJS.ErrnoException;
      outcome.failures.push({ path: t.path, why: err.code ?? err.message });
    }
  }

  outcome.groups = [...byGroup]
    .map(([name, g]) => ({ name, ...g }))
    .sort((a, b) => a.name.localeCompare(b.name));
  return outcome;
}
