#!/usr/bin/env node
import { mkdir,readdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, join, relative, resolve, sep } from 'node:path';
import { runCase, type CaseOptions } from './case.js';
import { recalculatePair } from './recalc.js';
import {
  writeSummary, UNSPECIFIED_TYPE, type CaseRecord, type CaseVerdict, type SummaryFiles,
} from './summary.js';
import { detectSpec } from './detect.js';
import { summarizeSweep } from './sweep.js';
import { makeBare, cachedValueState } from './bare.js';
import { proposeMeta } from './propose.js';
import { mergeSheetSpec } from './workbook.js';
import type { LedgerScope } from './ledger.js';
import type {
  DiffResult, WorkbookDiffResult, WorkbookSheetSpec, WorkbookSpec,
} from './types.js';

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

/**
 * Where the run summary goes, at the root of whatever was run.
 *
 * The name is chosen to sort first, because this is the first thing anyone
 * should read after a run and `results` would land alphabetically in the middle
 * of the report types. `_summary` was the obvious try and is wrong: Windows
 * orders a leading underscore *after* letters, so it sorted to the bottom --
 * measured, not assumed. `!` sorts ahead of letters and digits alike in both
 * Explorer and .NET, which is why it is a Windows convention for pinning a
 * folder to the top.
 */
const SUMMARY_DIR = '!summary';

/**
 * The folder each case writes into, from `--results`. Module state because the
 * folder walker is recursive and this is a property of the run rather than of
 * any one directory -- threading it through every frame would say less.
 */
let resultDir: string = RESULT_DIR;
const SPREADSHEET = new Set(['.xlsx', '.xlsm', '.csv', '.tsv', '.txt']);

const GOLDEN = /^(golden|baseline|expected|before)\b/i;
const ACTUAL = /^(actual|new|current|after|report)\b/i;

interface Args {
  target: string;
  bless: boolean;
  printSpec: boolean;
  writeMeta: boolean;
  writeExpect: boolean;
  ledger: LedgerScope;
  /** How much per-finding detail report.md writes out. */
  detail: 'capped' | 'full';
  sweep: boolean;
  bare: boolean;
  help: boolean;
  /** Have Excel work the formulas out before comparing. See recalc.ts. */
  recalc: boolean;
  /** Folder each case writes its artefacts into. Default `results`. */
  results: string;
  /** Rebuild the run summary from results on disk, comparing nothing. */
  summaryOnly: boolean;
}

function parseArgs(argv: string[]): Args {
  const args: Args = {
    target: '', bless: false, printSpec: false, writeMeta: false, writeExpect: false,
    ledger: 'differences', detail: 'capped', sweep: true, bare: false, help: false,
    recalc: false, results: RESULT_DIR, summaryOnly: false,
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!;
    if (a === '-h' || a === '--help') args.help = true;
    else if (a === '--bless' || a === '--update') args.bless = true;
    else if (a === '--print-spec') args.printSpec = true;
    else if (a === '--write-meta') args.writeMeta = true;
    else if (a === '--write-expect') args.writeExpect = true;
    else if (a === '--no-sweep') args.sweep = false;
    else if (a === '--bare') args.bare = true;
    else if (a === '--recalc') args.recalc = true;
    else if (a === '--summary') args.summaryOnly = true;
    else if (a === '--results') args.results = argv[++i] ?? RESULT_DIR;
    else if (a.startsWith('--results=')) args.results = a.slice(10);
    else if (a === '--ledger') args.ledger = argv[++i] as LedgerScope;
    else if (a.startsWith('--ledger=')) args.ledger = a.slice(9) as LedgerScope;
    else if (a === '--detail') args.detail = argv[++i] as 'capped' | 'full';
    else if (a.startsWith('--detail=')) args.detail = a.slice(9) as 'capped' | 'full';
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
          differences.xlsx    one row per differing cell, plus a sheet naming the
                              cells that will recalculate; absent if neither exists
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
  --write-meta       write a starting meta.json for a report type, from the
                     pairs themselves, instead of typing one out
  --write-expect     record what each case verified into its case.json, so a
                     table that later stops being compared fails the run
  --ledger <scope>   all | differences | none      (default: differences)
  --detail <level>   capped | full                  (default: capped)
                     How much per-finding detail report.md writes out. Capped
                     shows the first ten rows of each finding and names the
                     file holding the rest; the counts and the per-column
                     tallies are always complete. Use full when you want every
                     row in the markdown -- one case ran to 18,661 lines and
                     1.6 MB, which is why it is no longer the default.
  --no-sweep         skip layer 2. Saves a second parse of both files
  --recalc           have Excel work the formulas out before comparing. These
                     reports arrive with formulas and no stored results, so
                     without this a formula compares as its text alone and a
                     total that moved by billions reads as identical. Copies are
                     recalculated; golden/ and current/ are never touched.
                     Windows with Excel installed, and Excel must be closed
  --results <name>   folder each case writes into. Default "results". Use it to
                     keep a plain run and a --recalc run side by side
  --summary          rebuild !summary/ from the results already on disk and
                     stop. Compares nothing, so it is instant -- and it covers
                     the whole tree however narrow the last run was
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
  /**
   * Files sitting in golden/ or current/ that no comparison read.
   *
   * A case holding results.csv plus details.zip and unused.zip was reported
   * "Identical" having compared one of the three, and said nothing about the
   * other two. Naming them is the difference between a pass and a pass nobody
   * can trust.
   */
  uncompared: string[];
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
async function roleFolder(
  dir: string,
  roles: string[],
): Promise<{ file: string; others: string[] } | undefined | Problem> {
  let found: string | undefined;
  let entries: string[] = [];
  let others: string[] = [];
  for (const role of roles) {
    try {
      const all = (await readdir(join(dir, role), { withFileTypes: true }))
        .filter((e) => e.isFile())
        .map((e) => e.name);
      if (found) continue;
      found = role;
      entries = all.filter(isSpreadsheet);
      // Everything else in there. A golden/ or current/ folder exists to hold
      // one side's output, so a file in it that nothing compares is an
      // artefact nobody checked -- and a case reporting "Identical" while
      // ignoring two of its three files is the exact failure this tool is for.
      others = all.filter((f) => !isSpreadsheet(f));
    } catch { /* no folder by that name */ }
  }
  if (!found) return undefined;

  if (!entries.length) {
    return {
      problem: others.length
        ? `${found}/ holds no .xlsx, .xlsm or .csv file, only [${others.sort().join(', ')}]`
        : `${found}/ holds no .xlsx, .xlsm or .csv file`,
    };
  }
  // Picking one would be a guess, and the whole point of the folder is that
  // the file names no longer say which is which.
  if (entries.length > 1) {
    return {
      problem: `${found}/ holds ${entries.length} spreadsheets [${entries.sort().join(', ')}]`
        + ' — it must hold exactly one',
    };
  }
  return { file: join(found, entries[0]!), others: others.map((f) => join(found, f)) };
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
      return {
        name,
        dir,
        golden: join(dir, g.file),
        actual: join(dir, a.file),
        uncompared: [...g.others, ...a.others].map(DISPLAY),
      };
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
  return { name, dir, golden: join(dir, golden), actual: join(dir, actual), uncompared: [] };
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
    // Both names: a run writing to results-recalculated/ must still not walk
    // into the results/ folder an earlier run left, and read it as a case.
    .filter(
      (e) =>
        e.isDirectory() &&
        e.name !== RESULT_DIR &&
        e.name !== resultDir &&
        // The summary folder holds a workbook with no golden beside it. Walked
        // into, it reads as a broken case and stops the next run dead.
        e.name !== SUMMARY_DIR,
    )
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
 * Folders that declare themselves a report type, at any depth.
 *
 * Used to catch one that holds no cases at all. A type whose downloads stopped
 * arriving, or whose cases were moved, leaves a folder with a `meta.json` in
 * it, contributes nothing to the run and appears in no summary -- which from
 * the outside is indistinguishable from a type where everything passed.
 */
async function reportTypeFolders(
  root: string,
  dir: string,
  cases: Case[],
  into: string[] = [],
): Promise<string[]> {
  // A case is a leaf. Nothing inside one is a report type.
  if (cases.some((c) => c.dir === dir)) return into;

  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return into;
  }

  if (dir !== root) {
    // An unreadable meta.json is reported by the run that reads it properly.
    // The folder still counts here, so a type with a broken config is not also
    // an invisible one.
    const spec = await readJson(join(dir, 'meta.json'))
      .catch(() => ({ reportType: '?' }) as WorkbookSpec);
    if (spec && typeNamedBy(spec)) into.push(dir);
  }

  for (const e of entries) {
    if (
      e.isDirectory() && e.name !== RESULT_DIR && e.name !== resultDir && e.name !== SUMMARY_DIR
    ) {
      await reportTypeFolders(root, join(dir, e.name), cases, into);
    }
  }
  return into;
}

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
  'label', 'reportType', 'analysisType', 'entityType', 'source', 'expect', '//',
]);

/**
 * Keys that only mean something inside one sheet's settings. Finding one at the
 * top level is what identifies the mistake below.
 */
const SHEET_KEYS = new Set([
  'sheet', 'headerRow', 'endRow', 'columns', 'keyColumns', 'matchRowsByPosition',
  'fillKeyDown', 'keySeparator', 'tolerance', 'relativeTolerance',
  'ignoreColumns', 'ignoreRows',
  'metadataCells', 'compareFormulas', 'formulaMode', 'requireCachedValues',
  'trimStrings', 'looseHeaders', 'strictSchema', 'invariants', 'csv', 'tables',
]);

/** A value shaped like a sheet's settings: an object naming at least one. */
function looksLikeSheetSpec(value: unknown): boolean {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  return Object.keys(value).some((k) => SHEET_KEYS.has(k));
}

/**
 * Catches a config that will quietly do nothing.
 *
 * Per-sheet settings belong under `sheets`, and writing them at the top level
 * is the easy mistake -- `{ "Occupancy": { "keyColumns": [...] } }` parses, is
 * accepted, and has no effect whatsoever. Nothing downstream would ever
 * mention it, and the sheet would go on being compared without the key.
 *
 * That mistake, and nothing else. This used to reject every key it did not
 * recognise, which is a different and wrong rule: a `case.json` is a good place
 * to describe a case to *whatever generates it* -- the datasets it runs on, the
 * units, the environment -- and this tool has no business refusing a file
 * because somebody else's fields are in it. Anything not shaped like a sheet's
 * settings is left alone.
 */
function misplacedSheetSpecs(spec: WorkbookSpec): string[] {
  return Object.entries(spec)
    // Anything starting with "//" is a note. JSON has no comments, so that is
    // the convention people reach for.
    .filter(([k]) => !CONFIG_KEYS.has(k) && !k.startsWith('//'))
    .filter(([, v]) => looksLikeSheetSpec(v))
    .map(([k]) => k);
}

/**
 * The shape each of this tool's own keys has to have.
 *
 * A key it does not recognise is somebody else's and is left alone -- see
 * `misplacedSheetSpecs`. But a key it DOES recognise is read as its own, and a
 * generator that happens to use the same word writes something of a different
 * shape into it. `metadata` is the one that bit: this tool wants a list of cell
 * addresses to read but not judge, and an e2e harness wrote an object
 * describing the analysis. The comparison then failed with
 * "(given.metadata ?? []) is not iterable", which names neither the file nor
 * the key nor what was expected.
 */
const CONFIG_SHAPES: Record<string, { is: (v: unknown) => boolean; want: string }> = {
  metadata: { is: (v) => Array.isArray(v), want: 'an array of cell addresses or column names' },
  ignoreSheets: { is: (v) => Array.isArray(v), want: 'an array of sheet names' },
  cases: { is: (v) => Array.isArray(v), want: 'an array of case names' },
  sheets: { is: isPlainObject, want: 'an object keyed by sheet name' },
  defaults: { is: isPlainObject, want: "an object of settings applied to every sheet" },
  expect: { is: isPlainObject, want: 'an object keyed by sheet name' },
  label: { is: (v) => typeof v === 'string', want: 'a string' },
  reportType: { is: (v) => typeof v === 'string', want: 'a string' },
  analysisType: { is: (v) => typeof v === 'string', want: 'a string' },
  entityType: { is: (v) => typeof v === 'string', want: 'a string' },
  strictSheets: { is: (v) => typeof v === 'boolean', want: 'true or false' },
  matchUnkeyedRowsByPosition: { is: (v) => typeof v === 'boolean', want: 'true or false' },
};

function isPlainObject(v: unknown): boolean {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

/** Names every key of this tool's that holds something it cannot read. */
function wrongShapes(spec: WorkbookSpec): string[] {
  const out: string[] = [];
  for (const [key, { is, want }] of Object.entries(CONFIG_SHAPES)) {
    const v = (spec as Record<string, unknown>)[key];
    if (v === undefined || is(v)) continue;
    const got = Array.isArray(v) ? 'an array' : v === null ? 'null'
      : v === undefined ? 'nothing' : `${/^[aeiou]/.test(typeof v) ? 'an' : 'a'} ${typeof v}`;
    out.push(`"${key}" should be ${want}, and is ${got}`);
  }
  return out;
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

  // Checked after parsing, and outside the catch, so these read as what they
  // are rather than being re-labelled syntax errors.
  const wrong = wrongShapes(spec);
  if (wrong.length) {
    throw new Error(
      [
        `${DISPLAY(path)} cannot be read as settings:`,
        ...wrong.map((w) => `  ${w}`),
        '',
        "These are this tool's own key names. A file that describes a case to whatever",
        'generates it can carry any other field it likes -- only these are read here, so',
        'rename the clashing one and the rest of the file is left alone.',
      ].join('\n'),
    );
  }

  const stray = misplacedSheetSpecs(spec);
  if (stray.length) {
    throw new Error(
      `${DISPLAY(path)} has ${stray.length === 1 ? "a sheet's settings" : "sheets' settings"}` +
      ` at the top level, where they do nothing: ` +
      `${stray.map((k) => JSON.stringify(k)).join(', ')}.\n` +
      '  Per-sheet settings go inside "sheets". If you meant a sheet named ' +
      `${JSON.stringify(stray[0])}, write:\n` +
      `    { "sheets": { ${JSON.stringify(stray[0])}: { … } } }\n` +
      `  Recognised at the top level: ${[...CONFIG_KEYS].join(', ')}.\n` +
      '  Anything else is left alone, so a case.json can also describe the case ' +
      'to whatever generates it.',
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

/**
 * Keeps the table a sheet already had when a layer above adds one to it.
 *
 * A sheet holding a single table carries its `headerRow` and key at the sheet
 * level, with no `tables` block -- that is what makes reports read "Policies"
 * rather than "Policies · Table 1". Declaring `tables` on such a sheet used to
 * *replace* it: writing one entry to check a title block also stopped the
 * sheet's real table being compared, and nothing said so. Adding a table has
 * to mean adding a table.
 *
 * Only fires when the layer above declares tables and the one below has none,
 * so a sheet-level `keyColumns` written on its own -- the common correction,
 * and the one in every example -- behaves exactly as it did.
 *
 * The existing table is filed under the sheet's own name, which is the name
 * reports already give it.
 */
function namedTables(
  sheet: string,
  under: WorkbookSheetSpec | undefined,
  override: WorkbookSheetSpec,
): WorkbookSheetSpec | undefined {
  if (!under || under.tables || !override.tables) return under;
  if (under.headerRow === undefined && !under.keyColumns?.length) return under;

  const { headerRow, endRow, columns, keyColumns, ...rest } = under;
  return {
    ...rest,
    tables: {
      [sheet]: {
        ...(headerRow === undefined ? {} : { headerRow }),
        ...(endRow === undefined ? {} : { endRow }),
        ...(columns === undefined ? {} : { columns }),
        ...(keyColumns === undefined ? {} : { keyColumns }),
      },
    },
  };
}

function mergeSpecs(base: WorkbookSpec, given: WorkbookSpec): WorkbookSpec {
  const sheets: Record<string, WorkbookSheetSpec> = { ...base.sheets };
  for (const [name, override] of Object.entries(given.sheets ?? {})) {
    const found = Object.keys(sheets).find((k) => k.toLowerCase() === name.toLowerCase());
    const under = found ? sheets[found] : undefined;
    sheets[found ?? name] = mergeSheetSpec(namedTables(found ?? name, under, override), override);
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
  if (base.expect || given.expect) {
    // Per sheet, like `sheets`: a case naming what one sheet should hold must
    // not delete the report type's expectations for every other sheet.
    merged.expect = { ...base.expect, ...given.expect };
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

/**
 * Records what a run verified into the case's own case.json, as `expect`.
 *
 * Ranges, one per table, keyed by sheet -- read off the outcome rather than
 * off the config, so it is a statement about what happened. The rest of the
 * file is left exactly as it was: this is written into a file people also
 * write by hand, and losing a label or a note to a generated key would be a
 * poor trade for the guard it buys.
 */
async function writeExpectations(dir: string, diff: WorkbookDiffResult): Promise<number> {
  const expect: Record<string, string[]> = {};
  for (const o of diff.sheets) {
    if (o.status !== 'compared' || !o.range?.base) continue;
    (expect[o.sheet] ??= []).push(o.range.base);
  }

  const path = join(dir, 'case.json');
  const existing = (await readJson(path)) ?? {};
  const text = JSON.stringify({ ...existing, expect }, null, 2);
  await writeFile(path, text.endsWith('\n') ? text : `${text}\n`, 'utf8');
  return Object.keys(expect).length;
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
 * What kind of report a case belongs to, for the summaries.
 *
 * `reportType` is written once in the type folder's `meta.json` and inherited
 * by every case below it. Without one the folder path stands in, inside a name
 * that says plainly that nobody set it -- so it reads as a gap to fill rather
 * than as a report type somebody chose, and two unnamed folders stay two
 * groups instead of merging into one heading and one file.
 */
/**
 * What a config file calls the kind of thing it holds.
 *
 * `reportType` is this tool's name for it, and a generator naturally reaches
 * for the word its own domain uses: a Conditional EP is an analysis, a Data
 * Transmittal is an entity. Six of twelve type folders in one tree said
 * `analysisType` or `entityType` and every one of them was filed under
 * "Unspecified report type" -- the summary heading that exists to mean nobody
 * set this, printed over folders where somebody plainly had.
 */
const TYPE_KEYS = ['reportType', 'analysisType', 'entityType'] as const;

const typeNamedBy = (spec: WorkbookSpec): string | undefined => {
  for (const k of TYPE_KEYS) {
    const v = (spec as Record<string, unknown>)[k];
    if (typeof v === 'string' && v.trim()) return v.trim();
  }
  return undefined;
};

function reportTypeOf(spec: WorkbookSpec, root: string, caseDir: string): string {
  const named = typeNamedBy(spec);
  if (named) return named;
  // Empty for a case sitting at the root of the tree, which has no folder
  // above it to name and needs no parenthetical saying so.
  const folder = DISPLAY(relative(root, dirname(caseDir)));
  return folder ? `${UNSPECIFIED_TYPE} (${folder})` : UNSPECIFIED_TYPE;
}

/**
 * The folder a report type's own summary belongs in: the one whose `meta.json`
 * named the type, or the folder holding the cases when nothing did.
 *
 * A `case.json` is skipped even when it sets `reportType`. It sits inside a
 * single case, and a summary of a report type does not belong inside one of
 * its cases.
 */
function reportTypeDir(layers: ConfigLayer[], caseDir: string): string {
  for (let i = layers.length - 1; i >= 0; i--) {
    const layer = layers[i]!;
    if (basename(layer.path) === 'case.json') continue;
    if (typeNamedBy(layer.spec)) return dirname(layer.path);
  }
  return dirname(caseDir);
}

/**
 * The deepest folder holding all of them.
 *
 * Two folders can declare the same `reportType`, which makes them one type with
 * one summary -- and it has to sit somewhere above both rather than inside
 * whichever happened to run first.
 */
function commonAncestor(dirs: string[]): string {
  const split = dirs.map((d) => d.split(sep));
  const first = split[0]!;
  let i = 0;
  while (i < first.length && split.every((parts) => parts[i] === first[i])) i++;
  return first.slice(0, i).join(sep);
}

/**
 * Where each report type's summary goes. Built as the run goes, from the
 * configuration each case resolved, so a type is placed by what named it.
 */
class TypeFolders {
  private readonly seen = new Map<string, Set<string>>();

  add(type: string, layers: ConfigLayer[], caseDir: string): void {
    const dirs = this.seen.get(type) ?? new Set<string>();
    dirs.add(reportTypeDir(layers, caseDir));
    this.seen.set(type, dirs);
  }

  resolve(): Map<string, string> {
    const out = new Map<string, string>();
    for (const [type, dirs] of this.seen) {
      out.set(type, join(commonAncestor([...dirs]), SUMMARY_DIR, ...(
        // A named results folder gets a subfolder of its own here too, so a
        // plain run and a --recalc run keep their summaries side by side.
        resultDir === RESULT_DIR ? [] : [resultDir]
      )));
    }
    return out;
  }
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

/**
 * The failure, broken down by table.
 *
 * The one-line summary says how much failed and the report says everything;
 * between them sat the question people actually ask next -- *which* tables, and
 * what kind of difference. Without it the answer was "open report.md", and a
 * colleague reading only the terminal concluded the tool had not caught
 * differences it had caught and listed.
 *
 * Kept small on purpose. A run of forty cases must stay readable, so only the
 * worst few tables are named and the rest are counted.
 */
const BREAKDOWN_ROWS = 5;
const BREAKDOWN_NAME = 34;

interface Measure {
  head: string;
  cell(d: DiffResult): string;
  found(d: DiffResult): number;
}

const MEASURES: Measure[] = [
  { head: 'values', cell: (d) => count(d.values.length), found: (d) => d.values.length },
  { head: 'formulas', cell: (d) => count(d.formulas.length), found: (d) => d.formulas.length },
  { head: 'types', cell: (d) => count(d.types.length), found: (d) => d.types.length },
  {
    head: 'invariants',
    cell: (d) => count(d.invariants.length),
    found: (d) => d.invariants.length,
  },
  {
    head: 'rows +/-',
    cell: (d) => `${d.rows.added.length}/${d.rows.removed.length}`,
    found: (d) => d.rows.added.length + d.rows.removed.length,
  },
  {
    head: 'cols +/-',
    cell: (d) => `${d.schema.added.length}/${d.schema.removed.length}`,
    found: (d) => d.schema.added.length + d.schema.removed.length,
  },
  // Errors count towards `defects` like any difference, so they fail a table on
  // their own -- a key column the config names and the sheet does not have, a
  // formula with no cached value. Left out of this list they were invisible
  // twice over: the table appeared in the breakdown with a zero under every
  // column and nothing saying why, and because `weight` is the sum of these
  // measures it scored 0, sorted last, and dropped off the end of the five.
  // A table nobody could compare is the one most worth naming.
  { head: 'errors', cell: (d) => count(d.errors.length), found: (d) => d.errors.length },
];

const count = (v: number): string => v.toLocaleString('en-US');

/** Longest first, so the five that are named are the five worth naming. */
const weight = (d: DiffResult): number =>
  MEASURES.reduce((total, m) => total + m.found(d), 0);

function breakdown(diff: WorkbookDiffResult): string[] {
  const failing = diff.sheets
    .filter((o) => o.status === 'compared' && o.diff && !o.diff.ok)
    .sort((a, b) => weight(b.diff!) - weight(a.diff!) || a.label.localeCompare(b.label));
  if (!failing.length) return [];

  // A column of nothing but zeroes is noise, and judged against the tables
  // actually printed rather than against all of them -- deciding from the whole
  // set puts up a `cols +/-` column reading 0/0 five times because a table
  // nobody can see has one. Sorting by weight is what makes that safe: a
  // category with real numbers in it lifts its table into these five.
  const worst = failing.slice(0, BREAKDOWN_ROWS);
  const shown = MEASURES.filter((m) => worst.some((o) => m.found(o.diff!) > 0));
  if (!shown.length) return [];

  const head = ['table', ...shown.map((m) => m.head)];
  const rows = worst.map((o) => [
    o.label.length > BREAKDOWN_NAME ? `${o.label.slice(0, BREAKDOWN_NAME - 1)}…` : o.label,
    ...shown.map((m) => m.cell(o.diff!)),
  ]);

  const width = head.map((h, i) => Math.max(h.length, ...rows.map((r) => r[i]!.length)));
  // The name reads left to right; every count is a number and lines up on the
  // right, which is the only way a column of them can be compared by eye.
  const line = (cells: string[]) =>
    cells.map((c, i) => (i ? c.padStart(width[i]!) : c.padEnd(width[i]!))).join('  ').trimEnd();

  const out = [line(head), width.map((w) => '─'.repeat(w)).join('  ')];
  for (const r of rows) out.push(line(r));
  if (failing.length > rows.length) {
    out.push(`… and ${failing.length - rows.length} more table(s)`);
  }
  return out;
}

/**
 * Where a run's summaries go: `!summary/` at the tree root, and a subfolder of
 * it for a run given `--results <name>`, so a plain run and a `--recalc` run
 * keep their summaries instead of one overwriting the other. The default run
 * owns the top of the folder, which is the common path.
 */
const summaryDir = (root: string): string =>
  resultDir === RESULT_DIR ? join(root, SUMMARY_DIR) : join(root, SUMMARY_DIR, resultDir);

/** Fits a list onto one line, saying how many did not fit rather than cutting. */
function oneLine(items: string[], width = 96): string {
  const out: string[] = [];
  let used = 0;
  for (const item of items) {
    if (out.length && used + item.length + 2 > width) {
      return `${out.join(', ')} … and ${items.length - out.length} more`;
    }
    out.push(item);
    used += item.length + 2;
  }
  return out.join(', ');
}

function printTypeSummaries(written: SummaryFiles): void {
  if (!written.types.length) return;
  const n = written.types.length;
  console.log(
    `  ${n} report type summar${n === 1 ? 'y' : 'ies'}, each in its type's own ` +
    `${SUMMARY_DIR}/ — ` +
    // `x12`, not `(12)` -- a type with no name of its own already ends in a
    // parenthetical, and two in a row read as one broken one.
    oneLine(written.types.map((t) => `${t.type} ×${t.cases}`)),
  );
}

/** Wall clock, in the units a person waits in. */
function elapsed(ms: number): string {
  const secs = Math.round(ms / 1000);
  if (secs < 60) return `${secs}s`;
  return `${Math.floor(secs / 60)}m ${String(secs % 60).padStart(2, '0')}s`;
}

async function main(): Promise<number> {
  const started = Date.now();
  const typeFolders = new TypeFolders();
  const args = parseArgs(process.argv.slice(2));
  resultDir = args.results;
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
  const everything: Case[] = [];
  const selectedBy = new Map<Case, ConfigLayer[]>();
  for (const c of found) {
    try {
      selectedBy.set(c, await configLayers(root, c.dir));
      everything.push(c);
    } catch (e) {
      // One case's settings are one case's problem. This used to throw out of
      // the whole run, so a single file another tool happened to write stopped
      // every other case from being compared -- twenty-six reports unchecked
      // over one key, and nothing to say which of them would have passed.
      broken.push(`  ${DISPLAY(relative(root, c.dir))}: ${(e as Error).message}`);
    }
  }
  found = everything.filter((c) => selects(selectedBy.get(c)!, c.dir));
  const setAside = everything.length - found.length;

  if (!everything.length && broken.length) {
    console.error(`sheet-verify: no case could be read under ${target}`);
    console.error(broken.join('\n'));
    return 1;
  }

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
    let untouched = 0;
    for (const c of found) {
      for (const file of [c.golden, c.actual]) {
        const label = `${c.name}/${basename(file)}`;
        try {
          const result = await makeBare(file);
          if (!result) {
            // Counted, not printed. Running this twice over a tree of forty
            // cases put eighty lines on the screen saying nothing happened,
            // and buried the handful that said something did.
            untouched++;
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
    console.log(
      `\n${touched} file(s) rewritten` +
      (untouched ? `, ${untouched} already bare` : '') +
      '. Excel regenerates the results on open.',
    );
    return 0;
  }

  // Rebuilt from what each case's last run left behind, comparing nothing.
  //
  // Two things make this worth having. A run narrowed to one report type
  // overwrites the tree-wide summary with its own three cases, and a run of a
  // single case writes no summary at all -- so the overview drifts out of date
  // by ordinary use. And rebuilding is instant where re-comparing the tree is
  // a minute, or half an hour with --recalc.
  //
  // It walks the whole tree whatever the target was, because a partial summary
  // is the problem rather than the fix.
  if (args.summaryOnly) {
    const all = await findCases(root, root);
    const records: CaseRecord[] = [];
    for (const c of all.cases) {
      // Config layers only -- no detectSpec, which would open both workbooks.
      // The two things wanted here, the report type and the label, are written
      // in the config and never detected, so reading the files would buy
      // nothing and cost the speed that makes this worth having.
      const layers = await configLayers(root, c.dir);
      const spec = layers.reduce((acc, l) => mergeSpecs(acc, l.spec), {} as WorkbookSpec);
      const type = reportTypeOf(spec, root, c.dir);
      typeFolders.add(type, layers, c.dir);
      const label = layers.find((l) => basename(l.path) === 'case.json')?.spec.label;
      // Read plainly, not through readJson: that one validates a *config* and
      // rejects unknown keys, so every diff.json threw and every case looked
      // like it had never been run.
      const diffPath = join(c.dir, resultDir, 'diff.json');
      const diff = await readFile(diffPath, 'utf8')
        .then((t) => JSON.parse(t) as WorkbookDiffResult)
        .catch(() => null);
      if (!diff) {
        records.push({
          reportType: type, name: c.name, label,
          verdict: 'could not run', summary: 'never run — no results on disk',
        });
        continue;
      }
      const compared = diff.sheets.filter((o) => o.status === 'compared');
      records.push({
        reportType: type,
        name: c.name,
        label,
        verdict: diff.ok ? 'passed' : 'failed',
        summary: diff.ok ? 'identical' : `${compared.filter((o) => o.diff && !o.diff.ok).length} table(s) failing`,
        tablesCompared: compared.length,
        tablesFailing: compared.filter((o) => o.diff && !o.diff.ok).length,
        tablesNotCompared: diff.sheets.length - compared.length,
        report: join(c.dir, resultDir, 'report.md'),
      });
    }
    if (!records.length) {
      console.error(`sheet-verify: no cases under ${root}`);
      return 1;
    }
    const written = await writeSummary(
      summaryDir(root), records, typeFolders.resolve(), new Date(), { rebuilt: true },
    );

    const missing = records.filter((r) => r.verdict === 'could not run').length;
    const bad = records.filter((r) => r.verdict === 'failed').length;
    // What the file says, said here too. A rebuild that reports only how many
    // cases it read makes the reader open the file to learn whether the tree
    // is green, which is the one thing a summary line should already answer.
    console.log(
      `${DISPLAY(relative(process.cwd(), written.markdown))} — ${records.length} case(s): ` +
      [
        `${records.filter((r) => r.verdict === 'passed').length} passed`,
        `${bad} failing`,
        ...(missing ? [`${missing} with no results on disk`] : []),
      ].join(', '),
    );
    printTypeSummaries(written);
    return 0;
  }

  if (args.writeMeta) {
    const path = join(target, 'meta.json');
    const existing = await readJson(path).catch(() => ({} as WorkbookSpec));
    const settings = Object.keys(existing ?? {}).filter((k) => !k.startsWith('//'));
    if (settings.length) {
      // Refusing rather than merging. A generated file and a hand-written one
      // are different kinds of thing, and quietly folding one into the other
      // is how a setting somebody meant disappears.
      console.error([
        `sheet-verify: ${DISPLAY(relative(root, path))} already has settings in it: `
        + `${settings.join(', ')}.`,
        '  Move it aside and run this again to see what would be written,',
        '  or keep the file you have and add to it by hand.',
      ].join('\n'));
      return 1;
    }

    console.log(`sheet-verify: reading ${found.length} case(s) under ${DISPLAY(relative(root, target)) || '.'}`);
    const proposal = await proposeMeta(target, found);
    await writeFile(path, proposal.json, 'utf8');

    console.log('');
    for (const line of proposal.notes) console.log(`  ${line}`);
    console.log('');
    console.log(`written to ${DISPLAY(relative(root, path))} — read it before trusting it.`);
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
  const records: CaseRecord[] = [];
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

    // Worked out once, and registered against the folder that named it, so
    // every record below files the case under the same type and that type's
    // summary lands beside its own cases.
    const reportType = reportTypeOf(spec, root, c.dir);
    typeFolders.add(reportType, layers, c.dir);

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
      detail: args.detail,
      sweepCells: args.sweep,
      updateGolden: args.bless,
      names: {
        // Relative, not bare: a pair kept in golden/ and current/ folders needs
        // the folder carried through, and for a flat case this is the file name.
        golden: relative(c.dir, blessTo ?? c.golden),
        actual: relative(c.dir, c.actual),
        report: join(args.results, 'report.md'),
        diffJson: join(args.results, 'diff.json'),
        differences: join(args.results, 'differences.xlsx'),
        compared: join(args.results, 'compared.xlsx'),
      },
    };

    // A pair whose two sides were produced differently will report every
    // formula in the file as a value change. Saying so once, up front, beats
    // letting the reader work it out from hundreds of diffs.
    const skew = await calcSkew(c);

    const { head, path } = headline(c, spec.reportType, label);

    // Excel works the formulas out first, on copies, so the comparison has
    // values rather than formula text alone. The inputs are left as they
    // arrived; `names` is repointed at the copies, which sit under the results
    // folder and are cleared with it.
    let recalculated = false;
    let recalcActual = '';
    if (args.recalc) {
      try {
        const into = join(c.dir, args.results, 'recalculated');
        const pair = await recalculatePair(c.golden, c.actual, into);
        recalcActual = relative(c.dir, pair.actual);
        options.names = {
          ...options.names!,
          golden: relative(c.dir, pair.golden),
          actual: recalcActual,
        };
        recalculated = true;
        options.recalculated = true;
      } catch (e) {
        // Never fall back to comparing the files as they arrived: the run would
        // find far less than it promised to look for, and every silence would
        // read as a pass.
        process.stdout.write(`${head}
    ${(e as Error).message}

`);
        failed++;
        records.push({
          reportType,
          name: c.name, label, verdict: 'could not run', summary: (e as Error).message,
        });
        continue;
      }
    }

    try {
      const against = recalculated ? join(c.dir, recalcActual) : c.actual;
      const result = await runCase(against, c.dir, { ...options, uncompared: c.uncompared });
      // An artefact nobody compared is a coverage hole, not a difference, and
      // it fails the case for the same reason a table with no key does: the
      // verdict has to mean the whole case was checked.
      const ok = result.ok && !c.uncompared.length;
      if (!result.blessed && !ok) failed++;
      if (result.blessed && blessTo) await rm(c.golden, { force: true });

      const lines: string[] = path ? [path] : [];
      // Blessing into folders writes a file that was not there under that name,
      // so runCase calls it a creation. From here it is plainly a replacement,
      // and the file that went is worth naming.
      lines.push(result.blessed && blessTo
        ? `golden replaced by ${basename(c.actual)}, and ${basename(c.golden)} removed`
        : result.summary);
      // Indented under the summary it expands, so the block reads as belonging
      // to the case rather than as more lines about it.
      if (result.diff && !result.blessed) {
        for (const row of breakdown(result.diff)) lines.push(`  ${row}`);
      }
      if (recalculated) lines.push('recalculated by Excel before comparing');
      if (skew) lines.push(`! ${skew}`);
      // Named in the log as well as the report. A file nobody opened is the
      // kind of thing that has to interrupt a clean run, or it is discovered
      // the day somebody asks what the zip was for.
      if (c.uncompared.length) {
        lines.push(`! ${c.uncompared.length} file(s) in golden/ or current/ that nothing compared:`);
        for (const f of c.uncompared) lines.push(`    ${f}`);
      }
      // Layer 2 does not decide the outcome, but a case that passed while
      // something changed where nobody was looking is the one thing worth
      // interrupting a clean run to say.
      const gaps = !result.blessed && !!result.sweep && result.sweep.totalGaps > 0;
      if (gaps) lines.push(`! ${summarizeSweep(result.sweep!)}`);
      // The report is where the detail is, so its path goes last: either the
      // run found something, or it found something nobody had asked about.
      // Relative, like every other path in this log -- an absolute one here
      // wrapped the line and buried the block it was meant to close.
      if (!result.blessed && (!ok || gaps)) {
        lines.push(DISPLAY(relative(process.cwd(), result.files.report)));
      }

      // Recorded from the run rather than from the config, so it says what was
      // verified and not what somebody hoped would be. Deliberately its own
      // step and not folded into --bless: blessing accepts a change to the
      // output, and accepting a change to what is *checked* is a separate
      // decision that deserves to be made on purpose.
      if (args.writeExpect && result.diff) {
        const written = await writeExpectations(c.dir, result.diff);
        lines.push(`expectations recorded for ${written} sheet(s) in case.json`);
      }

      console.log(`${ok ? '✓' : '✗'} ${head}`);
      for (const line of lines) console.log(`    ${line}`);

      const compared = result.diff?.sheets.filter((o) => o.status === 'compared') ?? [];
      const verdict: CaseVerdict = result.blessed ? 'blessed' : ok ? 'passed' : 'failed';
      records.push({
        reportType,
        name: c.name,
        label,
        verdict,
        summary: result.summary,
        tablesCompared: compared.length,
        tablesFailing: compared.filter((o) => o.diff && !o.diff.ok).length,
        tablesNotCompared: (result.diff?.sheets.length ?? 0) - compared.length,
        uncheckedDiffering: result.sweep?.totalGaps ?? 0,
        recalculated,
        report: result.files.report,
      });
    } catch (e) {
      failed++;
      console.log(`✗ ${head}`);
      if (path) console.log(`    ${path}`);
      console.error(`    ${(e as Error).message}`);
      records.push({
        reportType,
        name: c.name, label, verdict: 'could not run', summary: (e as Error).message,
      });
    }
  }

  // One view of the whole run, grouped by report type, at the root of the tree
  // rather than inside any case, because it belongs to none of them.
  //
  // Written for a single case too, though the case's own report says more. The
  // alternative is worse: this file sits at the root describing "the run", so
  // skipping it leaves the *previous* run's summary in place, claiming
  // thirty-four cases when one was just compared. A stale overview reads as
  // current, which is the failure this whole tool exists to prevent.
  if (records.length) {
    // In its own folder, not beside the tree. A workbook at the root makes the
    // root itself look like a case with no golden beside it, and the next run
    // refuses to start: "no golden file. Rename one of [run-summary.xlsx]".
    //
    // A named results folder gets a subfolder of its own, so a plain run and a
    // --recalc run keep their summaries instead of one overwriting the other.
    // The default run owns the top of _summary/, which is the common path.
    const scoped = DISPLAY(relative(root, target));
    const written = await writeSummary(
      summaryDir(root), records, typeFolders.resolve(), new Date(),
      { target: scoped && scoped !== '.' ? scoped : undefined },
    );
    console.log(`
${DISPLAY(relative(process.cwd(), written.markdown))} — and the .xlsx beside it`);
    printTypeSummaries(written);
  }

  // Folders that were meant to be cases and could not be run. Counted as
  // failures: an unverified report is not a passing one.
  if (broken.length) {
    console.error(`\n${broken.length} folder(s) meant to be cases could not be run:`);
    console.error(broken.join('\n'));
  }

  // A report type folder with nothing under it. Measured against every case
  // found rather than every case run, so a type whose cases were merely left
  // out by a "cases" list is counted there and not accused of being empty.
  //
  // Not a failure: a folder can legitimately be waiting for its first
  // download. But it must not be silent -- with no cases it appears in no
  // summary at all, and a type nobody is checking looks exactly like a type
  // where everything passed.
  const idle = (await reportTypeFolders(root, target, everything))
    .filter((d) => !everything.some((c) => c.dir === d || c.dir.startsWith(d + sep)));
  if (idle.length) {
    console.log(
      `\n${idle.length} report type folder(s) held no cases:\n  ` +
      idle.map((d) => DISPLAY(relative(root, d))).join('\n  '),
    );
  }

  const n = found.length;
  console.log(
    `\n${n} case${n === 1 ? '' : 's'}, ${failed} failing`
    + (broken.length ? `, ${broken.length} could not be run` : '')
    + (setAside ? ` — ${setAside} not selected by "cases"` : '')
    // How long it took, because a run with --recalc is half an hour and a
    // plain one is a minute, and knowing which you just did is the difference
    // between waiting and going to look at something else.
    + ` (${elapsed(Date.now() - started)})`,
  );

  // The quiet number, said last so it is the one left on the screen. A run can
  // report zero failures while cells nobody keyed moved underneath it, and
  // that total lives in each case's report where a green run is never read.
  const unchecked = records.reduce((total, r) => total + (r.uncheckedDiffering ?? 0), 0);
  if (unchecked) {
    const cases = records.filter((r) => (r.uncheckedDiffering ?? 0) > 0).length;
    console.log(
      `${count(unchecked)} differing cell(s) nobody checked, across ${cases} case(s)` +
      ' — a table with no row key was not compared.',
    );
  }

  return failed || broken.length ? 1 : 0;
}

main().then(
  (code) => { process.exitCode = code; },
  (e) => {
    console.error(`sheet-verify: ${(e as Error).message}`);
    process.exitCode = 1;
  },
);
