#!/usr/bin/env node
import { readdir, readFile, stat } from 'node:fs/promises';
import { basename, extname, join, resolve } from 'node:path';
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

  With no folder, looks in ./${DEFAULT_ROOT} and runs every case inside it.
  Give a single case folder to run just that one.

LAYOUT
  ${DEFAULT_ROOT}/
    case_001/
      golden.xlsx      the output you trust        (golden|baseline|expected|before)
      actual.xlsx      the output under test       (actual|new|current|after|report)
      case.json        optional; overrides what was detected
      result/          written by this command
        diff.txt            human-readable summary — start here
        diff.json           the same, structured
        differences.xlsx    one row per differing cell; absent if none differed
        compared.xlsx       every cell checked, a worksheet per table

  Sheets, header rows and row keys are detected from the files, so no
  configuration is needed to start. CSV works the same way.

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
async function readCase(dir: string): Promise<Case | string> {
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

  return { name: basename(dir), dir, golden: join(dir, golden), actual: join(dir, actual) };
}

/** Case folders under a root, or the root itself when it is one. */
async function findCases(target: string): Promise<Case[] | string> {
  const direct = await readCase(target);
  if (typeof direct !== 'string') return [direct];

  let entries;
  try {
    entries = await readdir(target, { withFileTypes: true });
  } catch {
    return `no such folder: ${target}`;
  }

  const cases: Case[] = [];
  const problems: string[] = [];
  for (const e of entries.filter((x) => x.isDirectory()).sort((a, b) => a.name.localeCompare(b.name))) {
    const c = await readCase(join(target, e.name));
    if (typeof c === 'string') problems.push(`  ${e.name}: ${c}`);
    else cases.push(c);
  }

  if (!cases.length) {
    return problems.length
      ? `no runnable cases in ${target}\n${problems.join('\n')}`
      : `no cases in ${target}. ${direct}`;
  }
  return cases;
}

/** case.json, layered over what detection found. */
async function overrides(dir: string): Promise<WorkbookSpec> {
  try {
    return JSON.parse(await readFile(join(dir, 'case.json'), 'utf8')) as WorkbookSpec;
  } catch {
    return {};
  }
}

function mergeSpecs(detected: WorkbookSpec, given: WorkbookSpec): WorkbookSpec {
  const sheets: Record<string, WorkbookSheetSpec> = { ...detected.sheets };
  for (const [name, override] of Object.entries(given.sheets ?? {})) {
    const found = Object.keys(sheets).find((k) => k.toLowerCase() === name.toLowerCase());
    sheets[found ?? name] = mergeSheetSpec(found ? sheets[found] : undefined, override);
  }
  return { ...detected, ...given, sheets };
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

  const found = await findCases(target);
  if (typeof found === 'string') {
    console.error(`sheet-verify: ${found}`);
    return 1;
  }

  if (args.printSpec) {
    for (const c of found) {
      const spec = mergeSpecs(await detectSpec(c.golden), await overrides(c.dir));
      console.log(`// ${c.name} — save as ${join(c.dir, 'case.json')}`);
      console.log(JSON.stringify(spec, null, 2));
    }
    return 0;
  }

  let failed = 0;
  for (const c of found) {
    const spec = mergeSpecs(await detectSpec(c.golden), await overrides(c.dir));
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
