import ExcelJS from 'exceljs';
import { writeFile } from 'node:fs/promises';

/**
 * A run-level view: which cases passed, which failed, grouped by report type.
 *
 * Every artefact until now describes one case. That is the right shape for
 * fixing something and the wrong shape for the question asked first, which is
 * "how did the run go?" -- and the answer to it lived only in terminal
 * scrollback, where it is gone as soon as anyone scrolls, and cannot be sent to
 * somebody who was not watching.
 *
 * Grouped by report type because that is the unit people work in: a release
 * breaks a *kind* of report, not case_002 specifically, and eight failures on
 * one type with the rest clean is a different morning from one failure on each
 * of eight types.
 *
 * Written in both forms on purpose. The markdown is for reading and pasting
 * into a message; the workbook is for sorting and filtering, and for the
 * readers here who live in Excel and will not open a `.md` file.
 */

export type CaseVerdict = 'passed' | 'failed' | 'blessed' | 'could not run';

export interface CaseRecord {
  reportType: string;
  name: string;
  label?: string;
  verdict: CaseVerdict;
  /** One-line result, as the run log prints it. */
  summary: string;
  /**
   * Counted per *table*, not per sheet: a sheet holding four tables
   * contributes four. Named accordingly, because the one-line summary once
   * said "33 sheets failing" of a workbook with 22 sheets in it.
   */
  tablesCompared?: number;
  tablesFailing?: number;
  tablesNotCompared?: number;
  /** Differing cells layer 1 never looked at. The number worth watching. */
  uncheckedDiffering?: number;
  recalculated?: boolean;
  /** Path to the case's own report, which is where the detail is. */
  report?: string;
}

const ORDER: Record<CaseVerdict, number> = {
  failed: 0, 'could not run': 1, blessed: 2, passed: 3,
};

/** Report types in run order, each with its cases, failures first. */
function grouped(cases: CaseRecord[]): [string, CaseRecord[]][] {
  const byType = new Map<string, CaseRecord[]>();
  for (const c of cases) {
    const list = byType.get(c.reportType);
    if (list) list.push(c);
    else byType.set(c.reportType, [c]);
  }
  for (const list of byType.values()) {
    // Failures first within a type: the reason anyone opened this is to find
    // them, and a reader should not have to scroll past twelve ticks to.
    list.sort((a, b) => ORDER[a.verdict] - ORDER[b.verdict] || a.name.localeCompare(b.name));
  }
  return [...byType.entries()];
}

const tally = (list: CaseRecord[], v: CaseVerdict) => list.filter((c) => c.verdict === v).length;

const MARK: Record<CaseVerdict, string> = {
  passed: '✓', failed: '✗', blessed: '•', 'could not run': '!',
};

export interface SummaryScope {
  /**
   * What the run covered, relative to the tree root -- empty for the whole
   * tree. A run narrowed to one report type overwrites this file, and a reader
   * who does not know that reads "3 case(s)" as the size of the tree.
   */
  target?: string;
  /** Rebuilt from results already on disk rather than by a fresh comparison. */
  rebuilt?: boolean;
}

export function summaryMarkdown(
  cases: CaseRecord[],
  when = new Date(),
  scope: SummaryScope = {},
): string {
  const groups = grouped(cases);
  const failed = tally(cases, 'failed') + tally(cases, 'could not run');

  const out: string[] = [
    '# Run summary',
    '',
    `_${when.toISOString().slice(0, 16).replace('T', ' ')} · ${cases.length} case(s) ` +
      `across ${groups.length} report type(s)_`,
    '',
    failed
      ? `**${failed} of ${cases.length} case(s) need attention.**`
      : `**All ${cases.length} case(s) passed.**`,
    '',
    '| report type | cases | passed | failed | could not run |',
    '| --- | ---: | ---: | ---: | ---: |',
  ];

  // This file is written at the tree root whatever was run, so a narrowed run
  // replaces the last full one. Unsaid, "3 case(s)" reads as the size of the
  // tree rather than the size of the run.
  const note = scope.target
    ? `> **Scoped run — \`${scope.target}\` only.** Cases elsewhere in the tree were ` +
      'not compared and are not represented below. Run `npm run summary` to rebuild ' +
      'the whole tree from the results already on disk.'
    : scope.rebuilt
      ? '> Rebuilt from the results already on disk rather than from a fresh ' +
        'comparison. Each case is as its own last run left it.'
      : '';
  if (note) out.splice(5, 0, '', note);

  for (const [type, list] of groups) {
    const bad = tally(list, 'failed');
    out.push(
      `| ${type} | ${list.length} | ${tally(list, 'passed')} | ` +
        `${bad ? `**${bad}**` : '0'} | ${tally(list, 'could not run')} |`,
    );
  }

  for (const [type, list] of groups) {
    out.push('', `## ${type}`, '');
    out.push('| | case folder | what it is | result |', '| --- | --- | --- | --- |');
    for (const c of list) {
      // The label says what the case is for, and is the difference between a
      // reader recognising the failure and looking up what case_002 was.
      out.push(`| ${MARK[c.verdict]} | \`${c.name}\` | ${c.label ?? '—'} | ${c.summary} |`);
    }
    const watch = list.filter((c) => (c.uncheckedDiffering ?? 0) > 0);
    if (watch.length) {
      out.push(
        '',
        // A case can pass while something changed in a table layer 1 never
        // keyed. That is the one thing a green run can be hiding, so it is
        // called out here rather than left in the per-case report.
        `> **${watch.length} case(s) here have differing cells nobody checked.** ` +
          watch.map((c) => `\`${c.name}\` (${c.uncheckedDiffering})`).join(', ') +
          ' — a table with no row key was not compared. See the case report.',
      );
    }
  }

  out.push('', '---', '', '_Per-case detail is in each case\'s `results/report.md`._');
  return out.join('\n') + '\n';
}

const HEAD_FILL = 'FF1F3864';
const COLUMNS = [
  { header: '', key: 'mark', width: 4 },
  { header: 'Case folder', key: 'name', width: 34 },
  { header: 'What it is', key: 'label', width: 52 },
  { header: 'Result', key: 'summary', width: 46 },
  { header: 'Tables compared', key: 'tablesCompared', width: 16 },
  { header: 'Tables failing', key: 'tablesFailing', width: 14 },
  { header: 'Tables not compared', key: 'tablesNotCompared', width: 19 },
  { header: 'Unchecked differing', key: 'uncheckedDiffering', width: 19 },
  { header: 'Recalculated', key: 'recalculated', width: 13 },
  { header: 'Report', key: 'report', width: 64 },
];

function styleHead(ws: ExcelJS.Worksheet, width: number, row = 1): void {
  const r = ws.getRow(row);
  r.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  r.alignment = { vertical: 'middle', wrapText: true };
  r.height = 28;
  for (let c = 1; c <= width; c++) {
    r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD_FILL } };
  }
}

const VERDICT_STYLE: Record<CaseVerdict, { argb: string; bold: boolean }> = {
  failed: { argb: 'FF9F2F26', bold: true },
  'could not run': { argb: 'FF9F2F26', bold: true },
  blessed: { argb: 'FF6B7280', bold: false },
  passed: { argb: 'FF166534', bold: false },
};

export async function writeSummaryWorkbook(
  path: string,
  cases: CaseRecord[],
  when = new Date(),
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'sheet-verify';
  const groups = grouped(cases);

  const overview = wb.addWorksheet('Overview', { views: [{ state: 'frozen', ySplit: 3 }] });
  overview.getCell('A1').value = `Run summary — ${when.toISOString().slice(0, 16).replace('T', ' ')}`;
  overview.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF1F3864' } };
  overview.getRow(1).height = 24;

  const head = ['Report type', 'Cases', 'Passed', 'Failed', 'Could not run', 'Unchecked differing'];
  overview.addRow([]);
  overview.addRow(head);
  head.forEach((_, i) => { overview.getColumn(i + 1).width = i ? 17 : 34; });
  styleHead(overview, head.length, 3);

  for (const [type, list] of groups) {
    const row = overview.addRow([
      type,
      list.length,
      tally(list, 'passed'),
      tally(list, 'failed'),
      tally(list, 'could not run'),
      list.reduce((n, c) => n + (c.uncheckedDiffering ?? 0), 0),
    ]);
    if (tally(list, 'failed') + tally(list, 'could not run')) {
      row.getCell(1).font = { bold: true };
      row.getCell(4).font = { bold: true, color: { argb: 'FF9F2F26' } };
    }
  }
  const totals = overview.addRow([
    'All report types',
    cases.length,
    tally(cases, 'passed'),
    tally(cases, 'failed'),
    tally(cases, 'could not run'),
    cases.reduce((n, c) => n + (c.uncheckedDiffering ?? 0), 0),
  ]);
  totals.font = { bold: true };
  totals.eachCell((cell) => {
    cell.border = { top: { style: 'thin', color: { argb: 'FF1F3864' } } };
  });

  for (const [type, list] of groups) {
    // Excel forbids these in a sheet name, and truncates past 31 characters.
    const name = type.replace(/[\\/*?:[\]]/g, '-').slice(0, 31) || 'Report type';
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    COLUMNS.forEach((c, i) => { ws.getColumn(i + 1).width = c.width; });
    ws.addRow(COLUMNS.map((c) => c.header));
    styleHead(ws, COLUMNS.length);

    for (const c of list) {
      const row = ws.addRow([
        MARK[c.verdict],
        c.name,
        c.label ?? '',
        c.summary,
        c.tablesCompared ?? '',
        c.tablesFailing ?? '',
        c.tablesNotCompared ?? '',
        c.uncheckedDiffering ?? '',
        c.recalculated ? 'yes' : '',
        c.report ?? '',
      ]);
      const style = VERDICT_STYLE[c.verdict];
      row.getCell(1).font = { bold: true, color: { argb: style.argb } };
      row.getCell(2).font = { bold: style.bold, color: { argb: style.argb } };
      // A non-zero count here is the quiet one: the case may well have passed.
      if ((c.uncheckedDiffering ?? 0) > 0) {
        row.getCell(8).font = { bold: true, color: { argb: 'FF9F2F26' } };
      }
    }
    ws.autoFilter = { from: 'A1', to: { row: 1, column: COLUMNS.length } };
  }

  await wb.xlsx.writeFile(path);
}

export async function writeSummary(
  base: string,
  cases: CaseRecord[],
  when = new Date(),
  scope: SummaryScope = {},
): Promise<{ markdown: string; workbook: string }> {
  const markdown = `${base}.md`;
  const workbook = `${base}.xlsx`;
  await writeFile(markdown, summaryMarkdown(cases, when, scope), 'utf8');
  await writeSummaryWorkbook(workbook, cases, when);
  return { markdown, workbook };
}
