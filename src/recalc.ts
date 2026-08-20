import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { copyFile, mkdir } from 'node:fs/promises';
import { basename, dirname, extname, join, resolve } from 'node:path';

const run = promisify(execFile);

/**
 * Makes Excel work out what the formulas come to, so the comparison has values
 * to compare.
 *
 * These generators write formulas and no results. Across a tree of 33 real
 * cases that is 379,959 formula cells and not one carrying a cached value, so
 * every one of them compares as formula text alone: identical text on both
 * sides matches, and the numbers a reader sees in Excel -- which can differ by
 * billions -- never reach the comparison at all. Layer 2 does not help, because
 * an uncached formula contributes its text there too.
 *
 * `impact.ts` answers part of this by naming the cells that *will* move. It
 * cannot answer all of it: it finds what a formula reads by parsing its
 * references, and `OFFSET($A$1, MATCH(...), 0)` textually references `$A$1`
 * while actually reading wherever the arithmetic lands. 62,406 of the formulas
 * in that same tree are that shape, so a change reaching a cell through an
 * OFFSET is invisible to the graph.
 *
 * Nothing here evaluates a formula, and nothing here should: a spreadsheet
 * engine is an enormous thing to own and be wrong about. Excel already has one.
 * Opening a file and saving it stores every computed result, and the comparison
 * that follows is then an ordinary value comparison with no special cases.
 *
 * The inputs are never touched. A copy is recalculated and compared; `golden/`
 * and `current/` stay exactly as the generator wrote them, because they are the
 * record of what it produced and a run must not rewrite its own evidence.
 */

/** Where a recalculated pair was written. */
export interface Recalculated {
  golden: string;
  actual: string;
}

export class RecalcUnavailable extends Error {}

/**
 * PowerShell to open a workbook, force a full rebuild, and save it.
 *
 * Written as an argument list rather than a script file so there is nothing to
 * clean up, and with every risky default turned off: no alerts to answer, no
 * link updates to prompt for, no events, and macros forced off, since a report
 * may arrive as `.xlsm` and automation must not be the thing that runs it.
 *
 * `CalculateFullRebuild` rather than `Calculate`: a plain calculate honours
 * Excel's dependency tree, which is exactly the thing being distrusted here.
 */
const SCRIPT = [
  '$ErrorActionPreference = "Stop"',
  '$src = $args[0]; $dst = $args[1]',
  // Excel refuses to open, or opens read-only into Protected View, a file
  // carrying the mark-of-the-web. The copy is ours, so clearing it is safe.
  'try { Unblock-File -Path $src -ErrorAction SilentlyContinue } catch {}',
  '$xl = New-Object -ComObject Excel.Application',
  'try {',
  '  $xl.Visible = $false',
  '  $xl.DisplayAlerts = $false',
  '  $xl.AskToUpdateLinks = $false',
  '  $xl.EnableEvents = $false',
  '  $xl.AutomationSecurity = 3',
  '  $wb = $xl.Workbooks.Open($src, 0, $false)',
  '  $xl.CalculateFullRebuild()',
  '  $wb.SaveAs($dst, 51)',
  '  $wb.Close($false)',
  '} finally {',
  '  $xl.Quit()',
  '  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null',
  '}',
].join('\n');

/**
 * Whether Excel is already running.
 *
 * `New-Object -ComObject Excel.Application` can attach to a session a person
 * has open rather than starting its own, and this script ends by quitting
 * whatever it got hold of with alerts suppressed -- which would close their
 * workbooks and discard unsaved work without asking. Refusing is the only
 * responsible behaviour: the cost is a message telling someone to close Excel,
 * and the alternative cost is their morning.
 */
async function excelIsRunning(): Promise<boolean> {
  try {
    const { stdout } = await run('powershell', [
      '-NoProfile', '-NonInteractive', '-Command',
      '@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count',
    ]);
    return Number(stdout.trim()) > 0;
  } catch {
    return false; // no PowerShell, or no way to tell -- the open below will say
  }
}

/**
 * Recalculates one workbook into `dst`. The source is copied first, so the
 * original is neither opened nor written by Excel.
 */
async function recalcOne(src: string, dst: string): Promise<void> {
  await copyFile(src, dst);
  try {
    await run(
      'powershell',
      ['-NoProfile', '-NonInteractive', '-Command', SCRIPT, '-args', resolve(dst), resolve(dst)],
      { windowsHide: true, maxBuffer: 1 << 24 },
    );
  } catch (e) {
    const detail = String((e as { stderr?: string }).stderr || (e as Error).message)
      .split('\n')
      .map((l) => l.trim())
      .filter(Boolean)
      .slice(0, 3)
      .join(' ');
    throw new RecalcUnavailable(
      `sheet-verify: could not recalculate ${basename(src)} with Excel. ${detail}`,
    );
  }
}

/**
 * Recalculates both sides of a pair into `into`, and returns the two paths.
 *
 * Refuses rather than degrades. A run asked to recalculate and unable to do so
 * must not quietly compare the files as they arrived: it would report far fewer
 * differences than the run promised to look for, and every one of those
 * silences would look like a pass.
 */
export async function recalculatePair(
  golden: string,
  actual: string,
  into: string,
): Promise<Recalculated> {
  if (process.platform !== 'win32') {
    throw new RecalcUnavailable(
      'sheet-verify: --recalc drives Excel, which needs Windows. Run without it, and read ' +
        'the "Will recalculate" sheet in differences.xlsx for what would have moved.',
    );
  }
  if (await excelIsRunning()) {
    throw new RecalcUnavailable(
      'sheet-verify: Excel is open. Automation would attach to that session and close it, ' +
        'discarding anything unsaved. Close Excel and run again.',
    );
  }

  await mkdir(into, { recursive: true });
  const out = {
    golden: join(into, `golden${extname(golden)}`),
    actual: join(into, `actual${extname(actual)}`),
  };
  // One at a time. Two Excel instances on the same machine contend for the same
  // automation server, and the failure when they do is a timeout rather than an
  // error worth reading.
  await recalcOne(golden, out.golden);
  await recalcOne(actual, out.actual);
  return out;
}

/** Where a case's recalculated copies belong: beside its results, not its inputs. */
export const recalcDir = (reportPath: string): string =>
  join(dirname(reportPath), 'recalculated');
