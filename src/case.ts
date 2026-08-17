import { copyFile, mkdir, writeFile, access } from 'node:fs/promises';
import { createWriteStream } from 'node:fs';
import { join, basename, dirname } from 'node:path';
import type { WorkbookDiffResult, WorkbookSpec } from './types.js';
import { runWorkbook } from './workbook.js';
import { formatWorkbookReport, summarizeWorkbook, type ReportOptions } from './report.js';
import {
  ledgerCsvLines, writeLedgerWorkbook, writeComparedWorkbook, type LedgerScope,
} from './ledger.js';
import type { ComparedTable } from './workbook.js';

/**
 * A case is a folder. It holds the golden output the new report is judged
 * against, the report from the latest run, and the artefacts describing what
 * the comparison did -- so everything about one report type sits together and
 * can be reviewed, archived or attached to a ticket as a unit.
 *
 *   cases/monthly-policy-export/
 *     golden.xlsx     committed; the contract
 *     actual.xlsx     the latest run, copied in
 *     diff.txt        human-readable differences
 *     diff.json       the same, structured
 *     cells.xlsx      one row per differing cell, formatted for working through
 *     compared.xlsx   every cell compared, one worksheet per compared table
 */
export interface CaseFiles {
  golden: string;
  actual: string;
  diffText: string;
  diffJson: string;
  cells: string;
  compared: string;
}

export interface CaseOptions extends WorkbookSpec, ReportOptions {
  /** File names within the case folder. Defaults shown in `CaseFiles`. */
  names?: Partial<CaseFiles>;
  /**
   * How much of the cell-by-cell ledger to write. Default 'differences' --
   * a cell that matched needs no row. 'all' adds every matching cell too, as
   * a full audit trail, but grows with rows x columns; pair it with a `.csv`
   * ledger name so the file streams instead of being built in memory.
   */
  cellLedger?: LedgerScope;
  /**
   * Also write `compared.xlsx`: every cell compared, in a minimal column set,
   * split one worksheet per compared table. This is the file that grows with
   * the report, so turn it off for a case where it is not worth the time.
   * Default true.
   */
  comparedLedger?: boolean;
  /** Overwrite the golden output with the new report and pass. */
  updateGolden?: boolean;
  /** Create the golden output from the new report when absent. Default true. */
  createMissingGolden?: boolean;
}

export interface CaseResult {
  /** Case folder. */
  dir: string;
  /** Case name, taken from the folder. */
  name: string;
  /** Absolute paths of everything the run wrote or read. */
  files: CaseFiles;
  /** Null when the golden output was just created or re-blessed. */
  diff: WorkbookDiffResult | null;
  /** One-line verdict for logs and test titles. */
  summary: string;
  ok: boolean;
  /** True when the golden output was created or overwritten this run. */
  blessed: boolean;
}

const DEFAULTS: CaseFiles = {
  golden: 'golden.xlsx',
  actual: 'actual.xlsx',
  diffText: 'diff.txt',
  diffJson: 'diff.json',
  cells: 'cells.xlsx',
  compared: 'compared.xlsx',
};

const exists = (p: string) => access(p).then(() => true, () => false);

/** Streams CSV so a large ledger is never held in memory all at once. */
async function writeCsv(path: string, lines: Generator<string>): Promise<void> {
  const out = createWriteStream(path, { encoding: 'utf8' });
  try {
    for (const line of lines) {
      if (!out.write(line + '\n')) {
        await new Promise<void>((resolve) => out.once('drain', () => resolve()));
      }
    }
  } finally {
    await new Promise<void>((resolve, reject) => {
      out.end((err?: Error | null) => (err ? reject(err) : resolve()));
    });
  }
}

/**
 * The ledger format follows the file name: `cells.xlsx` gets a formatted
 * table, `cells.csv` gets streamed text. Naming the file is a clearer way to
 * choose than a separate option that could contradict it.
 */
async function writeLedger(
  path: string,
  tables: ComparedTable[],
  scope: LedgerScope,
): Promise<void> {
  if (scope === 'none') return;
  if (path.toLowerCase().endsWith('.csv')) {
    await writeCsv(path, ledgerCsvLines(tables, scope));
    return;
  }
  await writeLedgerWorkbook(path, tables, scope);
}

/**
 * Runs one case: copies the new report into the case folder, compares it
 * against the golden output, and writes the diff artefacts beside them.
 *
 * A missing golden output is created from the new report and the run passes,
 * the way a snapshot test behaves on first use. Commit it and review it.
 */
export async function runCase(
  actualPath: string,
  dir: string,
  options: CaseOptions = {},
): Promise<CaseResult> {
  const names = { ...DEFAULTS, ...options.names };
  const files: CaseFiles = {
    golden: join(dir, names.golden),
    actual: join(dir, names.actual),
    diffText: join(dir, names.diffText),
    diffJson: join(dir, names.diffJson),
    cells: join(dir, names.cells),
    compared: join(dir, names.compared),
  };
  const name = basename(dir);

  if (!(await exists(actualPath))) {
    throw new Error(`sheet-verify: new report not found\n  ${actualPath}`);
  }
  await mkdir(dir, { recursive: true });

  // The new report is kept in the case folder whatever the outcome, so a
  // failure can be opened next to the golden output it was judged against.
  if (join(actualPath) !== files.actual) await copyFile(actualPath, files.actual);

  const hasGolden = await exists(files.golden);
  if (!hasGolden && (options.createMissingGolden ?? true)) {
    await copyFile(files.actual, files.golden);
    return {
      dir, name, files, diff: null, blessed: true, ok: true,
      summary: `golden output created at ${files.golden} — commit it and review the contents`,
    };
  }
  if (!hasGolden) {
    throw new Error(`sheet-verify: golden output not found\n  ${files.golden}`);
  }
  if (options.updateGolden) {
    await copyFile(files.actual, files.golden);
    return {
      dir, name, files, diff: null, blessed: true, ok: true,
      summary: `golden output re-blessed from ${basename(actualPath)}`,
    };
  }

  const { diff, compared } = await runWorkbook(files.golden, files.actual, options);
  const report = formatWorkbookReport(diff, options);

  // A name may point into a subfolder -- the CLI keeps its output under
  // result/ so it cannot be mistaken for one of the two inputs.
  const parents = new Set(
    [files.diffText, files.diffJson, files.cells, files.compared].map((f) => dirname(f)),
  );
  await Promise.all([...parents].map((d) => mkdir(d, { recursive: true })));

  await Promise.all([
    writeFile(files.diffText, `${name}\n${'='.repeat(name.length)}\n\n${report}\n`, 'utf8'),
    writeFile(files.diffJson, JSON.stringify(diff, null, 2), 'utf8'),
    writeLedger(files.cells, compared, options.cellLedger ?? 'differences'),
    (options.comparedLedger ?? true)
      ? writeComparedWorkbook(files.compared, compared)
      : Promise.resolve(),
  ]);

  return {
    dir, name, files, diff, blessed: false,
    ok: diff.ok,
    summary: summarizeWorkbook(diff),
  };
}
