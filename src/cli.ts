#!/usr/bin/env node
import { readdir, readFile, stat } from 'node:fs/promises';
import { basename, dirname, extname, join, relative, resolve, sep } from 'node:path';
import { runCase, type CaseOptions } from './case.js';
import { detectSpec } from './detect.js';
import { mergeSheetSpec } from './workbook.js';
import type { LedgerScope } from './ledger.js';
import type { WorkbookSheetSpec, WorkbookSpec } from './types.js';

/**
 * Compares report folders. Put the golden output and the new report in a
 * folder, run this, read the result.
 *
 *   report-comparison/
 *     case_001/
 *       golden.xlsx     <- the output you trust
 *       actual.xlsx     <- the output under test
 *       result/         <- written by this command
 */

const DEFAULT_ROOT = 'report-comparison';
const RESULT_DIR = 'result';
const SPREADSHEET = new Set(['.xlsx', '.xlsm', '.csv', '.tsv', '.txt']);

const GOLDEN = /^(golden|baseline|expected|before)\b/i;
const ACTUAL = /^(actual|new|current|after|report)\b/i;

interface Args {
  target: string;
  bless: boolean;
  printSpec: boolean;
  ledger: LedgerScope;
  help: boolean;
}

function parseArgs(argv: string[]): Args {
  const args: Args = {
    target: '', bless: false, printSpec: false, ledger: 'differences', help: false,
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!;
    if (a === '-h' || a === '--help') args.help = true;
    else if (a === '--bless' || a === '--update') args.bless = true;
    else if (a === '--print-spec') args.printSpec = true;
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
    meta.json                          applies to every case below
    reports/
      global_standard_cat/
        meta.json                      applies to this report type
        case_001/
          golden.xlsx    the output you trust    (golden|baseline|expected|before)
          actual.xlsx    the output under test   (actual|new|current|after|report)
          case.json      optional; only this case
          result/        written by this command
            diff.txt            human-readable summary — start here
            diff.json           the same, structured
            differences.xlsx    one row per differing cell; absent if none differed
            compared.xlsx       every cell checked, a worksheet per table

  Configuration is inherited: every meta.json from the root down applies, and
  the case's own case.json wins. Settings a whole report type shares are
  written once; only a case that differs -- an extra sheet, say -- needs a file.

  Sheets, header rows and row keys are detected from the files, so no
  configuration is needed to start. CSV works the same way.

  Detection never invents a row key. If nothing identifies a row on some
  table, that table is NOT COMPARED and is reported as such -- a wrong key
  would pair rows arbitrarily and produce a confident wrong answer. Watch for
  "not compared" in the summary, and name the key in case.json:

    { "sheets": { "Summary": { "keyColumns": ["Region", "Band"] } } }

OPTIONS
  --bless            replace the golden output with the new report and pass
  --print-spec       print the detected layout as JSON and exit, so it can be
                     saved as case.json and edited
  --ledger <scope>   all | differences | none      (default: differences)
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
 * Finds the two inputs in a case folder. Named files win; failing that, if the
 * folder holds exactly two spreadsheets the older-named one is not guessed at
 * -- an explicit error beats comparing them the wrong way round.
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
 *   report-comparison/reports/global_standard_cat/case_001/
 *   report-comparison/analyses/marginal/case_003/
 */
async function findCases(
  root: string,
  dir: string = root,
  found: Case[] = [],
  problems: string[] = [],
): Promise<{ cases: Case[]; problems: string[] }> {
  const here = await readCase(dir, root);
  if (typeof here !== 'string') {
    found.push(here);
    return { cases: found, problems };
  }

  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    problems.push(`  ${DISPLAY(relative(root, dir)) || '.'}: cannot read folder`);
    return { cases: found, problems };
  }

  const subdirs = entries
    .filter((e) => e.isDirectory() && e.name !== RESULT_DIR)
    .sort((a, b) => a.name.localeCompare(b.name));

  // A leaf folder that is not a case is worth reporting; a grouping is not.
  if (!subdirs.length) problems.push(`  ${DISPLAY(relative(root, dir)) || '.'}: ${here}`);

  for (const s of subdirs) await findCases(root, join(dir, s.name), found, problems);
  return { cases: found, problems };
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

async function readJson(path: string): Promise<WorkbookSpec | null> {
  try {
    return JSON.parse(await readFile(path, 'utf8')) as WorkbookSpec;
  } catch (e) {
    // A malformed config is a mistake worth stopping for, but a missing one
    // is the normal case.
    if ((e as NodeJS.ErrnoException).code === 'ENOENT') return null;
    throw new Error(`${DISPLAY(path)} is not valid JSON: ${(e as Error).message}`);
  }
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
  return merged;
}

/** Detection, then every configuration layer over it in order. */
async function specFor(c: Case, layers: ConfigLayer[]): Promise<WorkbookSpec> {
  return layers.reduce((acc, l) => mergeSpecs(acc, l.spec), await detectSpec(c.golden));
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
  try {
    ({ cases: found, problems } = await findCases(root, target));
  } catch (e) {
    console.error(`sheet-verify: ${(e as Error).message}`);
    return 1;
  }

  if (!found.length) {
    console.error(`sheet-verify: no runnable cases under ${target}`);
    if (problems.length) console.error(problems.join('\n'));
    console.error(`\nA case is any folder holding a golden file and the report to compare:\n  ${DEFAULT_ROOT}/reports/<type>/case_001/golden.xlsx\n  ${DEFAULT_ROOT}/reports/<type>/case_001/actual.xlsx`);
    return 1;
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
    const spec = await specFor(c, await configLayers(root, c.dir));
    const options: CaseOptions = {
      ...spec,
      cellLedger: args.ledger,
      updateGolden: args.bless,
      names: {
        golden: basename(c.golden),
        actual: basename(c.actual),
        diffText: join('result', 'diff.txt'),
        diffJson: join('result', 'diff.json'),
        differences: join('result', 'differences.xlsx'),
        compared: join('result', 'compared.xlsx'),
      },
    };

    try {
      const result = await runCase(c.actual, c.dir, options);
      if (result.blessed) {
        console.log(`✓ ${c.name}  ${result.summary}`);
        continue;
      }
      if (result.ok) {
        console.log(`✓ ${c.name}  ${result.summary}`);
      } else {
        failed++;
        console.log(`✗ ${c.name}  ${result.summary}`);
        console.log(`    ${result.files.diffText}`);
      }
    } catch (e) {
      failed++;
      console.error(`✗ ${c.name}  ${(e as Error).message}`);
    }
  }

  const n = found.length;
  console.log(`\n${n} case${n === 1 ? '' : 's'}, ${failed} failing`);
  return failed ? 1 : 0;
}

main().then(
  (code) => { process.exitCode = code; },
  (e) => {
    console.error(`sheet-verify: ${(e as Error).message}`);
    process.exitCode = 1;
  },
);
