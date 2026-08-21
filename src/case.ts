import { copyFile, mkdir, writeFile, access, rm, stat } from 'node:fs/promises';
import { createWriteStream } from 'node:fs';
import { join, basename, dirname } from 'node:path';
import type { WorkbookDiffResult, WorkbookSpec } from './types.js';
import { runWorkbook } from './workbook.js';
import { summarizeWorkbook, type ReportOptions } from './report.js';
import { formatMarkdownReport } from './markdown.js';
import {
  ledgerCsvLines, writeLedgerWorkbook, writeComparedWorkbook, type LedgerScope,
} from './ledger.js';
import { sweep, type AffectedCell, type SweepResult } from './sweep.js';
import { DEFAULT_TOLERANCE } from './model.js';
import type { ComparedTable } from './workbook.js';

/**
 * A case is a folder. It holds the golden output the new report is judged
 * against, the report from the latest run, and the artefacts describing what
 * the comparison did -- so everything about one report type sits together and
 * can be reviewed, archived or attached to a ticket as a unit.
 *
 *   cases/monthly-policy-export/
 *     golden.xlsx        committed; the contract
 *     actual.xlsx        the latest run, copied in
 *     report.md          everything the run found -- start here
 *     diff.json          the same, structured
 *     differences.xlsx   one row per differing cell, plus the cells that will
 *                        recalculate once Excel opens the file
 *     compared.xlsx      every cell compared, one worksheet per table
 */
/**
 * Where each artefact goes. Every field is a path the run *would* use --
 * `differences` is not written at all when nothing differed, so check the file
 * exists before reading it.
 */
export interface CaseFiles {
  golden: string;
  actual: string;
  /** The whole report, in Markdown. What changed and what went unchecked. */
  report: string;
  diffJson: string;
  differences: string;
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
  /**
   * Also run layer 2: sweep every cell of both files by address and report what
   * differs outside layer 1's compared set. This is what turns "18 tables not
   * compared" into a measured statement rather than an open question, so it is
   * on by default. It costs a second parse of both files, and never changes the
   * outcome -- `ok` is layer 1's verdict alone.
   */
  sweepCells?: boolean;
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
  /** Layer 2. Null when it did not run. Never affects `ok`. */
  sweep: SweepResult | null;
  /** One-line verdict for logs and test titles. */
  summary: string;
  ok: boolean;
  /** True when the golden output was created or overwritten this run. */
  blessed: boolean;
}

const DEFAULTS: CaseFiles = {
  golden: 'golden.xlsx',
  actual: 'actual.xlsx',
  report: 'report.md',
  diffJson: 'diff.json',
  differences: 'differences.xlsx',
  compared: 'compared.xlsx',
};

const exists = (p: string) => access(p).then(() => true, () => false);

/**
 * Streams CSV so a large ledger is never held in memory all at once. Returns
 * the number of data rows, not counting the header.
 */
async function writeCsv(path: string, lines: Generator<string>): Promise<number> {
  const out = createWriteStream(path, { encoding: 'utf8' });
  let written = 0;
  try {
    for (const line of lines) {
      written++;
      if (!out.write(line + '\n')) {
        await new Promise<void>((resolve) => out.once('drain', () => resolve()));
      }
    }
  } finally {
    await new Promise<void>((resolve, reject) => {
      out.end((err?: Error | null) => (err ? reject(err) : resolve()));
    });
  }
  return Math.max(0, written - 1);
}

/**
 * The ledger format follows the file name: `differences.xlsx` gets a formatted
 * table, `differences.csv` gets streamed text. Naming the file is a clearer way
 * to choose than a separate option that could contradict it.
 *
 * A run with nothing to report leaves no file at all, and clears one an earlier
 * run left behind. An empty differences file reads as a fault rather than as
 * the answer, and a stale one is worse: it describes a comparison that no
 * longer holds.
 */
async function writeLedger(
  path: string,
  tables: ComparedTable[],
  scope: LedgerScope,
  affected: AffectedCell[],
): Promise<boolean> {
  const discard = async () => {
    await rm(path, { force: true });
    return false;
  };
  if (scope === 'none') return discard();

  const rows = path.toLowerCase().endsWith('.csv')
    ? await writeCsv(path, ledgerCsvLines(tables, scope))
    : await writeLedgerWorkbook(path, tables, scope, affected);

  return rows > 0 ? true : discard();
}

/**
 * The `*` tolerance from a workbook's defaults, whatever form it was written
 * in. A number means every column; a record names them, and `*` is its
 * fallback. Anything narrower than `*` belongs to a column, and a cell outside
 * every compared table has no column to be judged by.
 */
function blanketTolerance(tolerance: number | Record<string, number> | undefined): number {
  if (typeof tolerance === 'number') return tolerance;
  return tolerance?.['*'] ?? DEFAULT_TOLERANCE;
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
    report: join(dir, names.report),
    diffJson: join(dir, names.diffJson),
    differences: join(dir, names.differences),
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
      dir, name, files, diff: null, sweep: null, blessed: true, ok: true,
      summary: `golden output created at ${files.golden} — commit it and review the contents`,
    };
  }
  if (!hasGolden) {
    throw new Error(`sheet-verify: golden output not found\n  ${files.golden}`);
  }
  if (options.updateGolden) {
    await copyFile(files.actual, files.golden);
    return {
      dir, name, files, diff: null, sweep: null, blessed: true, ok: true,
      summary: `golden output re-blessed from ${basename(actualPath)}`,
    };
  }

  const { diff, compared } = await runWorkbook(files.golden, files.actual, options);

  // Layer 2 runs off the models layer 1 already built, so it costs a re-read of
  // the two files and nothing more. Its verdict is deliberately not folded into
  // `ok`: a positional sweep lights up whenever the layout moves, and letting
  // that fail a run would undo the point of aligning by key in the first place.
  const swept = (options.sweepCells ?? true)
    ? await sweep(files.golden, files.actual, compared, {
      ignoreSheets: options.ignoreSheets,
      metadata: options.metadata,
      // Per-column tolerances travel with the tables layer 1 compared. This is
      // the blanket one, for the cells layer 1 never reached -- without it a
      // tolerance would quiet the keyed comparison and leave the same float
      // noise in the headline count, which is where most people look first.
      tolerance: blanketTolerance(options.defaults?.tolerance),
    })
    : null;

  // Stamped as read, so a report can be told apart from a stale one describing
  // the same two paths. Best effort: a stamp that cannot be taken is left out
  // rather than failing a comparison that has already been made.
  const stampOf = async (f: string) => {
    try {
      const st = await stat(f);
      return { bytes: st.size, modified: st.mtime.toISOString().slice(0, 19).replace('T', ' ') };
    } catch {
      return undefined;
    }
  };
  const inputs = {
    golden: await stampOf(files.golden),
    actual: await stampOf(files.actual),
  };

  // A name may point into a subfolder -- the CLI keeps its output under
  // results/ so it cannot be mistaken for one of the two inputs.
  const parents = new Set(
    [files.report, files.diffJson, files.differences, files.compared].map((f) => dirname(f)),
  );
  await Promise.all([...parents].map((d) => mkdir(d, { recursive: true })));

  await Promise.all([
    // Written even when nothing differed: "identical, and here is what was
    // checked to say so" is the result, not the absence of one.
    writeFile(files.report, formatMarkdownReport(diff, swept, { name, ...options, inputs }), 'utf8'),
    writeFile(files.diffJson, JSON.stringify(diff, null, 2), 'utf8'),
    writeLedger(
      files.differences,
      compared,
      options.cellLedger ?? 'differences',
      swept?.affected ?? [],
    ),
    (options.comparedLedger ?? true)
      ? writeComparedWorkbook(files.compared, compared)
      : Promise.resolve(),
  ]);

  return {
    dir, name, files, diff, sweep: swept, blessed: false,
    ok: diff.ok,
    summary: summarizeWorkbook(diff),
  };
}
