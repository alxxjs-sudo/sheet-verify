import { copyFile, mkdir, access, readFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import { expect as baseExpect, test } from '@playwright/test';
import type { DiffResult, SheetSpec, WorkbookDiffResult, WorkbookSpec } from './types.js';
import { verifySheet } from './verify.js';
import { verifyWorkbook } from './workbook.js';
import { runCase, type CaseOptions, type CaseResult } from './case.js';
import {
  formatReport, formatWorkbookReport, summarize, summarizeWorkbook, type ReportOptions,
} from './report.js';
import { resolveSpec, KEY_SEP } from './model.js';

/** Baseline lifecycle switches, shared by both matchers. */
export interface BaselineOptions {
  /**
   * Write the actual file over the baseline and pass. Also enabled by
   * UPDATE_SHEET_BASELINE=1, so CI and local runs share one switch.
   */
  updateBaseline?: boolean;
  /**
   * Create the baseline and pass when it does not exist yet, the way
   * snapshot tests behave on first run. Default true.
   */
  createMissingBaseline?: boolean;
}

export interface MatcherOptions extends SheetSpec, ReportOptions, BaselineOptions {}

export interface WorkbookMatcherOptions
  extends WorkbookSpec, ReportOptions, BaselineOptions {}

const envUpdate = () =>
  ['1', 'true', 'yes'].includes(String(process.env.UPDATE_SHEET_BASELINE ?? '').toLowerCase());

const exists = (p: string) => access(p).then(() => true, () => false);

async function attach(name: string, body: string): Promise<void> {
  try {
    // The content type decides how reporters render the attachment, and how
    // tooling that filters attachments by type finds the machine-readable one.
    const contentType = name.endsWith('.json') ? 'application/json' : 'text/plain';
    await test.info().attach(name, { body, contentType });
  } catch {
    // Outside a running test; the message alone carries the detail.
  }
}

async function bless(baselinePath: string, actualPath: string): Promise<void> {
  await mkdir(dirname(baselinePath), { recursive: true });
  await copyFile(actualPath, baselinePath);
}

interface Verdict {
  name: string;
  pass: boolean;
  message: () => string;
}

/**
 * The baseline lifecycle both matchers share: validate the actual file, then
 * create, re-bless or fall through to comparison. Returns a verdict when the
 * run is already decided, null when the caller should compare.
 */
async function baselineGate(
  name: string,
  actualPath: string,
  baselinePath: string,
  options: BaselineOptions,
): Promise<Verdict | null> {
  if (typeof actualPath !== 'string') {
    return { name, pass: false, message: () => `${name}: expected a file path, received ${typeof actualPath}` };
  }
  if (!(await exists(actualPath))) {
    return { name, pass: false, message: () => `${name}: actual file not found\n  ${actualPath}` };
  }

  if (!(await exists(baselinePath))) {
    if (options.createMissingBaseline ?? true) {
      await bless(baselinePath, actualPath);
      await attach('sheet-baseline-created.txt', `Created baseline from actual output:\n  ${baselinePath}`);
      return {
        name, pass: true,
        message: () => `${name}: baseline created at ${baselinePath}. Commit it and review the contents.`,
      };
    }
    return { name, pass: false, message: () => `${name}: baseline not found\n  ${baselinePath}` };
  }

  if (options.updateBaseline || envUpdate()) {
    await bless(baselinePath, actualPath);
    return { name, pass: true, message: () => `${name}: baseline re-blessed from ${actualPath}` };
  }

  return null;
}

export const expect = baseExpect.extend({
  /**
   * Compares a generated sheet against its baseline, aligning columns by
   * header name and rows by business key so an inserted column is reported
   * as a schema change rather than as churn across every downstream cell.
   */
  async toMatchSheetBaseline(
    actualPath: string,
    baselinePath: string,
    options: MatcherOptions,
  ) {
    const name = 'toMatchSheetBaseline';

    const gated = await baselineGate(name, actualPath, baselinePath, options);
    if (gated) return gated;

    let diff: DiffResult;
    try {
      diff = await verifySheet(baselinePath, actualPath, options);
    } catch (e) {
      return { name, pass: false, message: () => `${name}: comparison failed\n  ${(e as Error).message}` };
    }

    const spec = resolveSpec(options);
    const report = formatReport(diff, {
      limit: options.limit,
      showCascades: options.showCascades,
      keySeparator: spec.keySeparator,
    });

    await attach('sheet-diff.txt', report);
    await attach('sheet-diff.json', JSON.stringify(diff, null, 2));

    if (diff.ok) {
      return {
        name, pass: true,
        message: () => `${name}: matched baseline (${summarize(diff)})\n\n${report}`,
      };
    }

    return {
      name, pass: false,
      message: () =>
        `${name}: output differs from baseline — ${summarize(diff)}\n\n${report}\n\n` +
        `If these changes are intended, re-bless the baseline:\n` +
        `  UPDATE_SHEET_BASELINE=1 npx playwright test\n`,
    };
  },

  /**
   * Compares every sheet of a generated workbook against its baseline in a
   * single pass. The baseline's sheet list is the contract: a sheet present
   * only in the new output is noted rather than compared, and one that has
   * disappeared is a failure.
   */
  async toMatchWorkbookBaseline(
    actualPath: string,
    baselinePath: string,
    options: WorkbookMatcherOptions,
  ) {
    const name = 'toMatchWorkbookBaseline';

    const gated = await baselineGate(name, actualPath, baselinePath, options);
    if (gated) return gated;

    let diff: WorkbookDiffResult;
    try {
      diff = await verifyWorkbook(baselinePath, actualPath, options);
    } catch (e) {
      return { name, pass: false, message: () => `${name}: comparison failed\n  ${(e as Error).message}` };
    }

    const report = formatWorkbookReport(diff, {
      limit: options.limit,
      showCascades: options.showCascades,
      // Display-only. A per-sheet override would just change how composite
      // keys are joined in the printout, so the workbook default is enough.
      keySeparator: options.defaults?.keySeparator ?? KEY_SEP,
    });

    await attach('sheet-diff.txt', report);
    await attach('sheet-diff.json', JSON.stringify(diff, null, 2));

    if (diff.ok) {
      return {
        name, pass: true,
        message: () => `${name}: matched baseline (${summarizeWorkbook(diff)})\n\n${report}`,
      };
    }

    return {
      name, pass: false,
      message: () =>
        `${name}: workbook differs from baseline — ${summarizeWorkbook(diff)}\n\n${report}\n\n` +
        `If these changes are intended, re-bless the baseline:\n` +
        `  UPDATE_SHEET_BASELINE=1 npx playwright test\n`,
    };
  },

  /**
   * Compares a generated report against the golden output in its case folder,
   * writing the new report and the diff artefacts into that folder so
   * everything about one report type stays together.
   */
  async toMatchCase(actualPath: string, dir: string, options: CaseOptions = {}) {
    const name = 'toMatchCase';

    if (typeof actualPath !== 'string') {
      return { name, pass: false, message: () => `${name}: expected a file path, received ${typeof actualPath}` };
    }

    let result: CaseResult;
    try {
      result = await runCase(actualPath, dir, {
        ...options,
        updateGolden: options.updateGolden || envUpdate(),
      });
    } catch (e) {
      return { name, pass: false, message: () => `${name}: ${(e as Error).message}` };
    }

    if (result.blessed) {
      return { name, pass: true, message: () => `${name}: ${result.summary}` };
    }

    const report = await readFile(result.files.diffText, 'utf8');
    await attach('sheet-diff.txt', report);
    await attach('sheet-diff.json', JSON.stringify(result.diff, null, 2));

    const written =
      `Case folder: ${result.dir}\n` +
      `  golden   ${result.files.golden}\n` +
      `  actual   ${result.files.actual}\n` +
      `  diff     ${result.files.diffText}\n` +
      `  json     ${result.files.diffJson}\n` +
      `  diffs    ${result.files.differences}\n`;

    if (result.ok) {
      return {
        name, pass: true,
        message: () => `${name}: ${result.name} matched (${result.summary})\n\n${written}`,
      };
    }

    return {
      name, pass: false,
      message: () =>
        `${name}: ${result.name} differs from its golden output — ${result.summary}\n\n` +
        `${report}\n\n${written}\n` +
        `If these changes are intended, re-bless the golden output:\n` +
        `  UPDATE_SHEET_BASELINE=1 npx playwright test\n`,
    };
  },
});

declare global {
  namespace PlaywrightTest {
    interface Matchers<R, T> {
      /**
       * Compares a generated Excel/CSV file against a committed baseline.
       * The receiver is the path to the generated file.
       */
      toMatchSheetBaseline(baselinePath: string, options: MatcherOptions): Promise<R>;
      /**
       * Compares every sheet of a generated workbook against a committed
       * baseline. The receiver is the path to the generated file.
       */
      toMatchWorkbookBaseline(
        baselinePath: string,
        options: WorkbookMatcherOptions,
      ): Promise<R>;
      /**
       * Compares a generated report against the golden output in its case
       * folder, writing the report and the diff artefacts into that folder.
       * The receiver is the path to the generated file.
       */
      toMatchCase(dir: string, options?: CaseOptions): Promise<R>;
    }
  }
}
