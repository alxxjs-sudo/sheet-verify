#!/usr/bin/env node
import { readdir, readFile, rm, stat } from 'node:fs/promises';
import { basename, dirname, extname, join, relative, resolve, sep } from 'node:path';
import { runCase, type CaseOptions } from './case.js';
import { detectSpec } from './detect.js';
import { summarizeSweep } from './sweep.js';
import { makeBare, cachedValueState } from './bare.js';
import { mergeSheetSpec } from './workbook.js';
import type { LedgerScope } from './ledger.js';
import type { WorkbookSheetSpec, WorkbookSpec } from './types.js';

/**
 * Compares report folders. Put the golden output and the new report in a
 * folder, run this, read the result.
 *
 *   output_comparison/
 *     case_001/
 *       golden.xlsx     <- the output you trust
 *       actual.xlsx     <- the output under test
 *       results/        <- written by this command
 */

const DEFAULT_ROOT = 'output_comparison';
const RESULT_DIR = 'results';
const SPREADSHEET = new Set(['.xlsx', '.xlsm', '.csv', '.tsv', '.txt']);

const GOLDEN = /^(golden|baseline|expected|before)\b/i;
const ACTUAL = /^(actual|new|current|after|report)\b/i;

interface Args {
  target: string;
  bless: boolean;
  printSpec: boolean;
  ledger: LedgerScope;
  sweep: boolean;
  bare: boolean;
  help: boolean;
}

function parseArgs(argv: string[]): Args {
  const args: Args = {
    target: '', bless: false, printSpec: false, ledger: 'differences',
    sweep: true, bare: false, help: false,
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!;
    if (a === '-h' || a === '--help') args.help = true;
    else if (a === '--bless' || a === '--update') args.bless = true;
    else if (a === '--print-spec') args.printSpec = true;
    else if (a === '--no-sweep') args.sweep = false;
    else if (a === '--bare') args.bare = true;
    else if (a === '--ledger') args.ledger = argv[++i] as LedgerScope;
    else if (a.startsWith('--ledger=')) args.ledger = a.slice(9) as LedgerScope;
    else if (!a.startsWith('-')) args.target ||= a;
  }
  return args;
}

const HELP = `
sheet-verify — compare a generated report against a golden output

USAGE
  sheet-verify [folder] [options]

  With no folder, looks in ./${DEFAULT_ROOT} and runs every case beneath it.
  Give any folder to run only what is under that one.

LAYOUT
  A case is any folder holding a golden file and the report to compare. Group
  them however suits you -- every folder above a case is just a grouping.

  ${DEFAULT_ROOT}/
    meta.json                        applies to every case below
    global_standard_cat_report/
      meta.json                      applies to this report type
      case_001/
        golden.xlsx    the output you trust    (golden|baseline|expected|before)
        actual.xlsx    the output under test   (actual|new|current|after|report)
        case.json      optional; only this case
        results/       written by this command
          report.md           everything the run found — start here
          diff.json           the same, structured
          differences.xlsx    one row per differing cell; absent if none differed
          compared.xlsx       every cell checked, a worksheet per table

  A pair can sit in folders instead, which is what a downloader produces when
  it keeps the name the source system gave each file -- that name carries the
  download's timestamp, and renaming it to golden.xlsx would throw that away:

    case_001/
      golden/case_1%1786955263151.xlsx     one spreadsheet, any name
      current/case_1%1786957329031.xlsx    same, and the folder says which
      results/

  Same folder names as the file forms: golden|baseline|expected|before, and
  current|actual|new|after|report. Two spreadsheets in one of them stops the
  run rather than picking one -- with names like those, a guess is a coin toss.
  Re-blessing writes the new golden under the name it arrived with and removes
  the file it replaced.

  Which cases run can be written down instead of typed. In any meta.json:

    { "cases": ["comparison_report/**", "!comparison_report/case_002"] }

  Paths are relative to the file's own folder; naming a folder takes what is
  inside it; a leading ! excludes. Cases left out are counted with the results,
  never dropped in silence.

TWO LAYERS
  Layer 1 aligns columns by header name and rows by business key, which is what
  lets it survive schema drift. It only compares tables it could key, and says
  so when it cannot -- so "not compared" is a coverage gap, not a pass.

  Layer 2 sweeps every cell of both files by address. It needs no keys, so it
  reaches everywhere, and reports the number that matters: cells that differ
  OUTSIDE layer 1's compared set. Zero means the gaps hid nothing.

  Layer 2 never decides the outcome. Being positional it lights up whenever the
  layout moves. Both layers write into report.md, in reading order.

  Neither layer judges report metadata -- the name, the id, whoever generated
  it and when. Those differ between any two runs by construction, so comparing
  them would fail every run on nothing. Name them under "metadata" in
  meta.json; report.md then lists them, with both values, under
  "Not verified, on purpose".

  Configuration is inherited: every meta.json from the root down applies, and
  the case's own case.json wins. Settings a whole report type shares are
  written once; only a case that differs -- an extra sheet, say -- needs a file.

  Two keys only name things. "reportType" in a type's meta.json and "label" in
  a case's case.json title the report and head the case in this log, so a
  failure reads as what it is:

    ✗ Validation Report · case_003 · a peril column added between two others

  "label" is deliberately not inherited -- one written above a folder would
  describe every case under it identically.

  Sheets, header rows and row keys are detected from the files, so no
  configuration is needed to start. CSV works the same way.

  Detection never invents a row key. If nothing identifies a row on some
  table, that table is NOT COMPARED and is reported as such -- a wrong key
  would pair rows arbitrarily and produce a confident wrong answer. Watch for
  "not compared" in the summary, and name the key in case.json:

    { "sheets": { "Summary": { "keyColumns": ["Region", "Band"] } } }

OPENED IN EXCEL?
  Reports arrive from the generator with formulas but no calculated results --
  Excel works them out on open. Opening one and SAVING it writes them all in.

  That is fine until it happens to one side of a pair: a bare golden against a
  saved report shows every formula in the file as a value change. In production
  neither side is ever opened, so this only bites when a test edit forces a
  file through Excel.

    sheet-verify --bare [folder]     put the files back as the generator wrote
                                     them, leaving your edit as the only change

  A bare file still SHOWS numbers when you open it. It carries no stored
  results, and fullCalcOnLoad tells Excel to work them out on the way in, so
  what is on screen is calculated live and the file itself is untouched. Seeing
  values is the flag working, not a sign the file needs re-baring.

  Because Excel did that calculation, it treats the workbook as modified and
  offers to save on close even when you have changed nothing. Say no. Saying
  yes is what writes the results in and skews the pair.

  To skip Protected View entirely, Unblock-File the report before opening it.

OPTIONS
  --bless            replace the golden output with the new report and pass
  --bare             strip calculated results from every file, in place, and
                     exit. Run after editing a report by hand
  --print-spec       print the detected layout as JSON and exit, so it can be
                     saved as case.json and edited
  --ledger <scope>   all | differences | none      (default: differences)
  --no-sweep         skip layer 2. Saves a second parse of both files
  -h, --help         this text

EXIT CODE
  0  every case matched
  1  a case differed, or could not be run
`;

interface Case {
  name: string;
  dir: string;
  golden: string;
  actual: string;
}

const isSpreadsheet = (f: string) => SPREADSHEET.has(extname(f).toLowerCase());

/**
 * A pair kept in folders rather than in file names:
 *
 *   case_001/golden/case_1%1786955263151.xlsx
 *   case_001/current/case_1%1786957329031.xlsx
 *
 * Which is what a downloader produces when it keeps the name the source system
 * gave the file. That name carries the download's timestamp, and renaming it to
 * `golden.xlsx` to satisfy this tool would throw that away, so the role moves
 * to the folder and the file keeps what it came with.
 *
 * Returns the one spreadsheet inside, the reason there isn't one, or undefined
 * when no such folder exists and the flat layout should be tried instead.
 */
async function roleFolder(dir: string, roles: string[]): Promise<string | undefined | Problem> {
  let found: string | undefined;
  let entries: string[] = [];
  for (const role of roles) {
    try {
      const inside = (await readdir(join(dir, role), { withFileTypes: true }))
        .filter((e) => e.isFile() && isSpreadsheet(e.name))
        .map((e) => e.name);
      if (!found) { found = role; entries = inside; }
    } catch { /* no folder by that name */ }
  }
  if (!found) return undefined;

  if (!entries.length) return { problem: `${found}/ holds no .xlsx, .xlsm or .csv file` };
  // Picking one would be a guess, and the whole point of the folder is that
  // the file names no longer say which is which.
  if (entries.length > 1) {
    return {
      problem: `${found}/ holds ${entries.length} spreadsheets [${entries.sort().join(', ')}]`
        + ' — it must hold exactly one',
    };
  }
  return join(found, entries[0]!);
}

interface Problem { problem: string }
const isProblem = (v: unknown): v is Problem =>
  typeof v === 'object' && v !== null && 'problem' in v;

/** Folder names that say which side a file is, when the file name does not. */
const GOLDEN_DIRS = ['golden', 'baseline', 'expected', 'before'];
const ACTUAL_DIRS = ['current', 'actual', 'new', 'after', 'report'];

/**
 * Finds the two inputs in a case folder. Named files win; failing that, if the
 * folder holds exactly two spreadsheets the older-named one is not guessed at
 * -- an explicit error beats comparing them the wrong way round. Failing that,
 * a `golden/` and `current/` pair of folders, which is how a download keeps the
 * name the source system gave each file.
 */
async function readCase(dir: string, root: string): Promise<Case | string> {
  let entries: string[];
  try {
    entries = (await readdir(dir, { withFileTypes: true }))
      .filter((e) => e.isFile() && isSpreadsheet(e.name))
      .map((e) => e.name);
  } catch {
    return `cannot read folder: ${dir}`;
  }

  const golden = entries.find((f) => GOLDEN.test(basename(f, extname(f))));
  if (!golden) {
    const [g, a] = await Promise.all([
      roleFolder(dir, GOLDEN_DIRS),
      roleFolder(dir, ACTUAL_DIRS),
    ]);
    if (g || a) {
      if (isProblem(g)) return g.problem;
      if (isProblem(a)) return a.problem;
      if (!g) return 'a current/ folder is here with no golden/ folder beside it';
      if (!a) return 'a golden/ folder is here with no current/ folder beside it';
      const name = DISPLAY(relative(root, dir)) || basename(dir);
      return { name, dir, golden: join(dir, g), actual: join(dir, a) };
    }
    return entries.length
      ? `no golden file. Rename one of [${entries.join(', ')}] to golden${extname(entries[0]!)}`
      : 'folder holds no .xlsx, .xlsm or .csv files';
  }

  const rest = entries.filter((f) => f !== golden);
  const actual = rest.find((f) => ACTUAL.test(basename(f, extname(f))))
    ?? (rest.length === 1 ? rest[0] : undefined);

  if (!actual) {
    return rest.length
      ? `cannot tell which file to compare. Rename one of [${rest.join(', ')}] to actual${extname(golden)}`
      : `only ${golden} is here. Add the report to compare against it`;
  }

  // The path, not the folder name: case_001 will exist under every report type.
  const name = DISPLAY(relative(root, dir)) || basename(dir);
  return { name, dir, golden: join(dir, golden), actual: join(dir, actual) };
}

const DISPLAY = (p: string) => p.split(sep).join('/');

/**
 * Every case folder at any depth. A folder holding a golden file is a case;
 * anything else is a grouping, so reports can be filed by kind:
 *
 *   output_comparison/reports/global_standard_cat/case_001/
 *   output_comparison/analyses/marginal/case_003/
 */
async function findCases(
  root: string,
  dir: string = root,
  found: Case[] = [],
  problems: string[] = [],
  broken: string[] = [],
): Promise<{ cases: Case[]; problems: string[]; broken: string[] }> {
  const here = await readCase(dir, root);
  if (typeof here !== 'string') {
    found.push(here);
    return { cases: found, problems, broken };
  }

  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    problems.push(`  ${DISPLAY(relative(root, dir)) || '.'}: cannot read folder`);
    return { cases: found, problems, broken };
  }

  const subdirs = entries
    .filter((e) => e.isDirectory() && e.name !== RESULT_DIR)
    .sort((a, b) => a.name.localeCompare(b.name));

  // A folder holding spreadsheets, or a golden/current pair of folders, was
  // meant to be a case. Say why it is not one, and stop: descending would
  // report its golden/ and current/ as two broken cases of their own and bury
  // the one sentence that explains the folder.
  const role = (name: string) =>
    GOLDEN_DIRS.includes(name.toLowerCase()) || ACTUAL_DIRS.includes(name.toLowerCase());
  const meant = subdirs.some((e) => role(e.name))
    || entries.some((e) => e.isFile() && isSpreadsheet(e.name));

  const line = `  ${DISPLAY(relative(root, dir)) || '.'}: ${here}`;
  // A folder that was meant to be a case is reported whatever else the run
  // finds. Reporting it only when nothing runs is how a case that stopped
  // being a case goes unnoticed: the total reads 38 instead of 39, and a
  // number nobody was watching is the whole failure this tool exists to catch.
  if (meant) {
    broken.push(line);
    return { cases: found, problems, broken };
  }
  // A leaf folder that is not a case is worth reporting; a grouping is not.
  if (!subdirs.length) problems.push(line);

  for (const s of subdirs) await findCases(root, join(dir, s.name), found, problems, broken);
  return { cases: found, problems, broken };
}

const exists = (p: string) => stat(p).then(() => true, () => false);

/**
 * The root of the comparison tree, which is not necessarily the folder asked
 * for: running one report type must apply the same configuration, and produce
 * the same case names, as running everything. So the root is the outermost
 * ancestor carrying a meta.json, and the target only decides what to run.
 */
export async function findRoot(target: string): Promise<string> {
  let root = target;
  let at = target;
  for (let i = 0; i < 16; i++) {
    const parent = dirname(at);
    if (parent === at) break;
    at = parent;
    if (await exists(join(at, 'meta.json'))) root = at;
  }
  return root;
}

/**
 * Fields a config file may carry. The first group changes the comparison; the
 * second is labelling, kept so a case can describe itself and where it came
 * from.
 */
const CONFIG_KEYS = new Set([
  'defaults', 'sheets', 'ignoreSheets', 'metadata', 'strictSheets',
  'matchUnkeyedRowsByPosition', 'cases',
  'label', 'reportType', 'source', '//',
]);

/**
 * Catches a config that will quietly do nothing.
 *
 * Per-sheet settings belong under `sheets`, and writing them at the top level
 * is the easy mistake -- `{ "Occupancy": { "keyColumns": [...] } }` parses, is
 * accepted, and has no effect whatsoever. Nothing downstream would ever
 * mention it, and the sheet would go on being compared without the key.
 */
function unknownKeys(spec: WorkbookSpec): string[] {
  // Anything starting with "//" is a note. JSON has no comments, so that is
  // the convention people reach for, and a config worth annotating usually
  // needs more than one note in it.
  return Object.keys(spec).filter((k) => !CONFIG_KEYS.has(k) && !k.startsWith('//'));
}

async function readJson(path: string): Promise<WorkbookSpec | null> {
  let spec: WorkbookSpec;
  try {
    // Windows editors and PowerShell write a UTF-8 BOM by default, and
    // JSON.parse refuses it. Config files get hand-edited, so the first save
    // from the wrong tool would otherwise fail the run with a parse error
    // pointing at a character nobody can see.
    const text = (await readFile(path, 'utf8')).replace(/^﻿/, '');
    spec = JSON.parse(text) as WorkbookSpec;
  } catch (e) {
    // A malformed config is a mistake worth stopping for, but a missing one
    // is the normal case.
    if ((e as NodeJS.ErrnoException).code === 'ENOENT') return null;
    throw new Error(`${DISPLAY(path)} is not valid JSON: ${(e as Error).message}`);
  }

  // Checked after parsing, and outside the catch, so this reads as what it is
  // rather than being re-labelled a syntax error.
  const stray = unknownKeys(spec);
  if (stray.length) {
    throw new Error(
      `${DISPLAY(path)} has ${stray.length === 1 ? 'a setting' : 'settings'} that would do ` +
      `nothing: ${stray.map((k) => JSON.stringify(k)).join(', ')}.\n` +
      '  Per-sheet settings go inside "sheets". If you meant a sheet named ' +
      `${JSON.stringify(stray[0])}, write:\n` +
      `    { "sheets": { ${JSON.stringify(stray[0])}: { … } } }\n` +
      `  Recognised at the top level: ${[...CONFIG_KEYS].join(', ')}`,
    );
  }
  return spec;
}

export interface ConfigLayer {
  path: string;
  spec: WorkbookSpec;
}

/**
 * Configuration inherited down the tree: a meta.json at any level applies to
 * every case beneath it, and the case's own case.json wins. So the settings a
 * whole report type shares are written once, and only a case that is genuinely
 * odd -- an extra sheet, say -- needs a file of its own.
 */
export async function configLayers(root: string, dir: string): Promise<ConfigLayer[]> {
  const rel = relative(root, dir);
  const steps = rel ? rel.split(sep) : [];
  const layers: ConfigLayer[] = [];

  let at = root;
  for (const step of [null, ...steps]) {
    if (step !== null) at = join(at, step);
    const path = join(at, 'meta.json');
    const spec = await readJson(path);
    if (spec) layers.push({ path, spec });
  }

  const own = join(dir, 'case.json');
  const spec = await readJson(own);
  if (spec) layers.push({ path: own, spec });

  return layers;
}

function mergeSpecs(base: WorkbookSpec, given: WorkbookSpec): WorkbookSpec {
  const sheets: Record<string, WorkbookSheetSpec> = { ...base.sheets };
  for (const [name, override] of Object.entries(given.sheets ?? {})) {
    const found = Object.keys(sheets).find((k) => k.toLowerCase() === name.toLowerCase());
    sheets[found ?? name] = mergeSheetSpec(found ? sheets[found] : undefined, override);
  }
  const merged: WorkbookSpec = { ...base, ...given, sheets };
  if (base.defaults || given.defaults) {
    merged.defaults = mergeSheetSpec(base.defaults, given.defaults);
  }
  if (base.ignoreSheets || given.ignoreSheets) {
    // Additive: a type excluding its glossary and a case excluding one more
    // should end up excluding both.
    merged.ignoreSheets = [...new Set([...(base.ignoreSheets ?? []), ...(given.ignoreSheets ?? [])])];
  }
  if (base.metadata || given.metadata) {
    // Additive for the same reason: a report type naming its own header block
    // should not lose the labels every report shares.
    merged.metadata = [...new Set([...(base.metadata ?? []), ...(given.metadata ?? [])])];
  }
  return merged;
}

/**
 * Names the one mistake that makes a comparison meaningless: two files that
 * disagree about whether formulas carry results. Returns null when the pair is
 * consistent, which is the normal case and needs no comment.
 */
async function calcSkew(c: Case): Promise<string | null> {
  const xlsx = (p: string) => /\.xlsm?$|\.xlsx$/i.test(p);
  if (!xlsx(c.golden) || !xlsx(c.actual)) return null;

  try {
    const [g, a] = await Promise.all([
      readFile(c.golden).then(cachedValueState),
      readFile(c.actual).then(cachedValueState),
    ]);
    if (g.cached === 0 && a.cached === 0) return null;
    if (g.cached > 0 && a.cached > 0) return null;

    const [bare, saved] = g.cached === 0 ? ['golden', 'report'] : ['report', 'golden'];
    const n = Math.max(g.cached, a.cached);
    return (
      `${bare} has no calculated results, ${saved} has ${n} — one of these was opened ` +
      'in Excel and saved. Every formula will read as a value change. Fix with --bare'
    );
  } catch {
    // Diagnosing the inputs must never be what stops a run.
    return null;
  }
}

/** Detection, then every configuration layer over it in order. */
async function specFor(c: Case, layers: ConfigLayer[]): Promise<WorkbookSpec> {
  return layers.reduce((acc, l) => mergeSpecs(acc, l.spec), await detectSpec(c.golden));
}

/**
 * One `cases` pattern as a test. `*` stops at a path separator and `**` does
 * not, so "reports/*" is one level and "reports/**" is everything under it.
 * A pattern that names a folder selects what is inside it.
 */
function casePattern(pattern: string): (path: string) => boolean {
  const body = pattern
    .split('**')
    .map((part) => part.replace(/[.+?^${}()|[\]\\]/g, '\\$&').split('*').join('[^/]*'))
    .join('.*');
  const re = new RegExp(`^${body}$`);
  return (path) => {
    // A pattern matching any folder on the way down selects what is under it,
    // so "comparison_report" needs no trailing wildcard to mean the cases
    // inside it.
    const segments = path.split('/');
    for (let i = 1; i <= segments.length; i++) {
      if (re.test(segments.slice(0, i).join('/'))) return true;
    }
    return false;
  };
}

/**
 * Whether a case is selected by the `cases` lists above it. Each file that
 * carries one narrows the set further, so a type can say which of its cases
 * are live and the root can say which types are.
 */
function selects(layers: ConfigLayer[], dir: string): boolean {
  for (const layer of layers) {
    const patterns = layer.spec.cases;
    if (!patterns?.length) continue;
    const from = DISPLAY(relative(dirname(layer.path), dir));
    // A case.json sits in the case itself, with nothing beneath it to choose
    // between. Selecting from there would only be a case switching itself off,
    // which is what an empty results folder and a missing report already fail
    // to explain to whoever goes looking.
    if (from === '') continue;
    const include = patterns.filter((p) => !p.startsWith('!'));
    const exclude = patterns.filter((p) => p.startsWith('!')).map((p) => p.slice(1));
    if (include.length && !include.some((p) => casePattern(p)(from))) return false;
    if (exclude.some((p) => casePattern(p)(from))) return false;
  }
  return true;
}

/**
 * How a case announces itself in the log: what kind of report it is, which
 * case, and what that case is for. The path moves to the line below, because
 * it is how the folder is found rather than what the run is about.
 *
 * A tree that names neither reads exactly as it always did -- the path alone,
 * with nothing invented to fill the space.
 */
function headline(
  c: Case,
  reportType: string | undefined,
  label: string | undefined,
): { head: string; path?: string } {
  const name = basename(c.dir);
  // A label repeating the folder name says nothing twice. Cases are commonly
  // labelled that way before anyone has written a real description, and
  // "case_001 · case_001" would make the log worse for having the feature.
  const said = label && label !== name ? label : undefined;
  if (!reportType && !said) return { head: c.name };
  return {
    head: [reportType, name, said].filter(Boolean).join(' · '),
    // A case sitting at the root of the tree is its own path, and printing
    // "case_001" under "case_001 · what it tests" says nothing twice.
    path: c.name === name ? undefined : c.name,
  };
}

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log(HELP.trim());
    return 0;
  }

  const target = resolve(args.target || DEFAULT_ROOT);
  try {
    await stat(target);
  } catch {
    console.error(`sheet-verify: no such folder: ${target}`);
    console.error(`\nCreate it and put a case inside:\n  ${DEFAULT_ROOT}/case_001/golden.xlsx\n  ${DEFAULT_ROOT}/case_001/actual.xlsx`);
    return 1;
  }

  const root = await findRoot(target);

  let found: Case[];
  let problems: string[];
  let broken: string[];
  try {
    ({ cases: found, problems, broken } = await findCases(root, target));
  } catch (e) {
    console.error(`sheet-verify: ${(e as Error).message}`);
    return 1;
  }

  if (!found.length) {
    console.error(`sheet-verify: no runnable cases under ${target}`);
    if (broken.length) console.error(broken.join('\n'));
    if (problems.length) console.error(problems.join('\n'));
    console.error(
      '\nA case is any folder holding a golden output and the report to compare,'
      + '\nnamed by file:'
      + `\n  ${DEFAULT_ROOT}/reports/<type>/case_001/golden.xlsx`
      + `\n  ${DEFAULT_ROOT}/reports/<type>/case_001/actual.xlsx`
      + '\nor named by folder, one spreadsheet in each:'
      + `\n  ${DEFAULT_ROOT}/reports/<type>/case_001/golden/<any name>.xlsx`
      + `\n  ${DEFAULT_ROOT}/reports/<type>/case_001/current/<any name>.xlsx`,
    );
    return 1;
  }

  // What the tree says to run. Applied before anything is read, so a case set
  // aside costs nothing, and counted afterwards, so it is never set aside
  // quietly -- a report that stops being checked without anyone noticing is
  // the failure this whole tool exists to prevent.
  const everything = found;
  const selectedBy = new Map<Case, ConfigLayer[]>();
  for (const c of everything) selectedBy.set(c, await configLayers(root, c.dir));
  found = everything.filter((c) => selects(selectedBy.get(c)!, c.dir));
  const setAside = everything.length - found.length;

  if (!found.length) {
    const naming = [...new Set(
      [...selectedBy.values()].flat()
        .filter((l) => l.spec.cases?.length)
        .map((l) => DISPLAY(relative(root, l.path))),
    )];
    console.error(`sheet-verify: "cases" selected none of the ${everything.length} cases under ${target}`);
    console.error(`  selected by: ${naming.join(', ')}`);
    return 1;
  }

  if (args.bare) {
    let touched = 0;
    for (const c of found) {
      for (const file of [c.golden, c.actual]) {
        const label = `${c.name}/${basename(file)}`;
        try {
          const result = await makeBare(file);
          if (!result) {
            console.log(`  ${label}  already bare`);
            continue;
          }
          touched++;
          const bits = [`${result.stripped} result(s) stripped`];
          if (result.droppedCalcChain) bits.push('calcChain dropped');
          if (result.restoredFullCalc) bits.push('fullCalcOnLoad restored');
          console.log(`✓ ${label}  ${bits.join(', ')}`);
        } catch (e) {
          console.error(`✗ ${label}  ${(e as Error).message}`);
          return 1;
        }
      }
    }
    console.log(`\n${touched} file(s) rewritten. Excel regenerates the results on open.`);
    return 0;
  }

  if (args.printSpec) {
    for (const c of found) {
      const layers = await configLayers(root, c.dir);
      console.log(`// ${c.name}`);
      for (const l of layers) console.log(`//   layered with ${DISPLAY(relative(root, l.path))}`);
      console.log(`//   save edits as ${DISPLAY(relative(root, join(c.dir, 'case.json')))}`);
      console.log(JSON.stringify(await specFor(c, layers), null, 2));
    }
    return 0;
  }

  let failed = 0;
  for (const c of found) {
    const layers = await configLayers(root, c.dir);
    const spec = await specFor(c, layers);
    // `label` is the one setting that must not inherit. Every other key
    // describes how to compare, and sharing that down a tree is the point;
    // a label describes what one case is, so taken from a folder above it
    // would put the same sentence on every case beneath and tell nobody
    // anything. `reportType` is the opposite -- it names the family, so it
    // comes down the tree like the rest.
    const label = layers.find((l) => basename(l.path) === 'case.json')?.spec.label;

    // Re-blessing a pair kept in folders writes the new golden under the name
    // the new report came with, since that name carries the download it is.
    // Copying the content into the old file would leave a golden stamped with
    // a timestamp belonging to a report it no longer holds. The file it
    // replaces is removed below, so the folder never ends up with two.
    const inFolders = dirname(c.golden) !== c.dir;
    const blessTo = args.bless && inFolders && basename(c.golden) !== basename(c.actual)
      ? join(dirname(c.golden), basename(c.actual))
      : undefined;

    const options: CaseOptions = {
      ...spec,
      label,
      cellLedger: args.ledger,
      sweepCells: args.sweep,
      updateGolden: args.bless,
      names: {
        // Relative, not bare: a pair kept in golden/ and current/ folders needs
        // the folder carried through, and for a flat case this is the file name.
        golden: relative(c.dir, blessTo ?? c.golden),
        actual: relative(c.dir, c.actual),
        report: join('results', 'report.md'),
        diffJson: join('results', 'diff.json'),
        differences: join('results', 'differences.xlsx'),
        compared: join('results', 'compared.xlsx'),
      },
    };

    // A pair whose two sides were produced differently will report every
    // formula in the file as a value change. Saying so once, up front, beats
    // letting the reader work it out from hundreds of diffs.
    const skew = await calcSkew(c);

    const { head, path } = headline(c, spec.reportType, label);

    try {
      const result = await runCase(c.actual, c.dir, options);
      if (!result.blessed && !result.ok) failed++;
      if (result.blessed && blessTo) await rm(c.golden, { force: true });

      const lines: string[] = path ? [path] : [];
      // Blessing into folders writes a file that was not there under that name,
      // so runCase calls it a creation. From here it is plainly a replacement,
      // and the file that went is worth naming.
      lines.push(result.blessed && blessTo
        ? `golden replaced by ${basename(c.actual)}, and ${basename(c.golden)} removed`
        : result.summary);
      if (skew) lines.push(`! ${skew}`);
      // Layer 2 does not decide the outcome, but a case that passed while
      // something changed where nobody was looking is the one thing worth
      // interrupting a clean run to say.
      const gaps = !result.blessed && !!result.sweep && result.sweep.totalGaps > 0;
      if (gaps) lines.push(`! ${summarizeSweep(result.sweep!)}`);
      // The report is where the detail is, so its path goes last: either the
      // run found something, or it found something nobody had asked about.
      if (!result.blessed && (!result.ok || gaps)) lines.push(result.files.report);

      console.log(`${result.ok ? '✓' : '✗'} ${head}`);
      for (const line of lines) console.log(`    ${line}`);
    } catch (e) {
      failed++;
      console.log(`✗ ${head}`);
      if (path) console.log(`    ${path}`);
      console.error(`    ${(e as Error).message}`);
    }
  }

  // Folders that were meant to be cases and could not be run. Counted as
  // failures: an unverified report is not a passing one.
  if (broken.length) {
    console.error(`\n${broken.length} folder(s) meant to be cases could not be run:`);
    console.error(broken.join('\n'));
  }

  const n = found.length;
  console.log(
    `\n${n} case${n === 1 ? '' : 's'}, ${failed} failing`
    + (broken.length ? `, ${broken.length} could not be run` : '')
    + (setAside ? ` — ${setAside} not selected by "cases"` : ''),
  );
  return failed || broken.length ? 1 : 0;
}

main().then(
  (code) => { process.exitCode = code; },
  (e) => {
    console.error(`sheet-verify: ${(e as Error).message}`);
    process.exitCode = 1;
  },
);
