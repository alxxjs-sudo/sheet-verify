import { copyFile, mkdir, writeFile, access } from 'node:fs/promises';
import { createWriteStream } from 'node:fs';
import { join, basename } from 'node:path';
import type { WorkbookDiffResult, WorkbookSpec } from './types.js';
import { runWorkbook } from './workbook.js';
import { formatWorkbookReport, summarizeWorkbook, type ReportOptions } from './report.js';
import { ledgerRows, type LedgerScope } from './ledger.js';

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
 *     cells.csv       every cell compared, and its verdict
 */
export interface CaseFiles {
  golden: string;
  actual: string;
  diffText: string;
  diffJson: string;
  cells: string;
}

export interface CaseOptions extends WorkbookSpec, ReportOptions {
  /** File names within the case folder. Defaults shown in `CaseFiles`. */
  names?: Partial<CaseFiles>;
  /**
   * How much of the cell-by-cell ledger to write. 'all' records every cell
   * compared including matches, which is the full audit trail but grows with
   * rows x columns. Default 'all'.
   */
  cellLedger?: LedgerScope;
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
  cells: 'cells.csv',
};

const exists = (p: string) => access(p).then(() => true, () => false);

/** Streams the ledger so a large case does not build the whole CSV in memory. */
async function writeLedger(
  path: string,
  rows: Generator<string>,
): Promise<void> {
  const out = createWriteStream(path, { encoding: 'utf8' });
  try {
    for (const row of rows) {
      if (!out.write(row + '\n')) {
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

  await Promise.all([
    writeFile(files.diffText, `${name}\n${'='.repeat(name.length)}\n\n${report}\n`, 'utf8'),
    writeFile(files.diffJson, JSON.stringify(diff, null, 2), 'utf8'),
    writeLedger(files.cells, ledgerRows(compared, options.cellLedger ?? 'all')),
  ]);

  return {
    dir, name, files, diff, blessed: false,
    ok: diff.ok,
    summary: summarizeWorkbook(diff),
  };
}
