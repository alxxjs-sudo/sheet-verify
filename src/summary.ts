import ExcelJS from 'exceljs';
import { mkdir, readdir, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

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
 *
 * Each report type also gets a file of its own, in its own folder. The
 * whole-run view is the right shape for "how did the run go" and the wrong
 * shape for the message that actually gets sent, which is about one kind of
 * report and goes to the person who owns it. Sending them the tree-wide
 * summary means sending them ten types they do not work on.
 *
 * Those started out beside the run summary and were moved: eleven types is
 * twenty-two files in one folder, which is more to scan than it is worth, and
 * a type's summary is easiest to find in the folder already being worked in.
 */

export type CaseVerdict = 'passed' | 'failed' | 'blessed' | 'could not run';

/**
 * What a report type is called when nothing named it.
 *
 * `reportType` is written in the type folder's `meta.json` and inherited by
 * every case beneath it. A folder that has no `meta.json`, or one that never
 * set the key, still has to be called something -- and the folder path is
 * appended so two unnamed types stay two types rather than merging into one
 * heading and one file.
 */
export const UNSPECIFIED_TYPE = 'Unspecified report type';

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

const n = (v: number): string => v.toLocaleString('en-US');
const stamp = (when: Date): string => when.toISOString().slice(0, 16).replace('T', ' ');
const sum = (list: CaseRecord[], pick: (c: CaseRecord) => number | undefined): number =>
  list.reduce((total, c) => total + (pick(c) ?? 0), 0);

/** The rows both the run summary and a single type's summary are built from. */
function caseTable(list: CaseRecord[]): string[] {
  const out = ['| | case folder | what it is | result |', '| --- | --- | --- | --- |'];
  for (const c of list) {
    // The label says what the case is for, and is the difference between a
    // reader recognising the failure and looking up what case_002 was.
    out.push(`| ${MARK[c.verdict]} | \`${c.name}\` | ${c.label ?? '—'} | ${c.summary} |`);
  }
  return out;
}

/**
 * A case can pass while something changed in a table layer 1 never keyed. That
 * is the one thing a green run can be hiding, so it is called out beside the
 * cases rather than left in the per-case report.
 */
function watchNote(list: CaseRecord[]): string | undefined {
  const watch = list.filter((c) => (c.uncheckedDiffering ?? 0) > 0);
  if (!watch.length) return undefined;
  return `> **${watch.length} case(s) here have differing cells outside the keyed comparison.** ` +
    watch.map((c) => `\`${c.name}\` (${n(c.uncheckedDiffering!)})`).join(', ') +
    ' — layer 2 found them; layer 1 had no row key for their table. See the case report.';
}

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
    `_${stamp(when)} · ${cases.length} case(s) across ${groups.length} report type(s)_`,
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
        'comparison. Each case is as its own last run left it. **Differing cells ' +
        "outside the keyed comparison are not counted here** — that total belongs " +
        "to layer 2, which is not in `diff.json`. Each case report has it."
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
    out.push(...caseTable(list));
    const watch = watchNote(list);
    if (watch) out.push('', watch);
  }

  out.push(
    '', '---', '',
    '_Per-case detail is in each case\'s `results/report.md`. Each report type also',
    'has a summary of its own, in that type\'s own folder under `!summary/`._',
  );
  return out.join('\n') + '\n';
}

/**
 * One report type, on its own page.
 *
 * The same cases as that type's section of the run summary, with the totals
 * spelled out rather than left to be added up -- this is the file that gets
 * sent to whoever owns the report, and they should not have to read past ten
 * other types to reach theirs.
 */
export function typeSummaryMarkdown(
  type: string,
  list: CaseRecord[],
  when = new Date(),
  scope: SummaryScope = {},
): string {
  const sorted = [...list].sort(
    (a, b) => ORDER[a.verdict] - ORDER[b.verdict] || a.name.localeCompare(b.name),
  );
  const need = tally(sorted, 'failed') + tally(sorted, 'could not run');

  const out: string[] = [
    `# ${type}`,
    '',
    `_${stamp(when)} · ${sorted.length} case(s)_`,
    '',
    need
      ? `**${need} of ${sorted.length} case(s) need attention.**`
      : `**All ${sorted.length} case(s) passed.**`,
    '',
  ];

  // Same reasoning as the run summary: this file is overwritten by whatever
  // ran last, so a run narrowed to one case would otherwise leave a page
  // claiming the type has one case in it.
  const note = scope.target
    ? `> **Scoped run — \`${scope.target}\` only.** Cases of this type outside that ` +
      'path were not compared and are not listed below. Run `npm run summary` to ' +
      'rebuild every type from the results already on disk.'
    : scope.rebuilt
      ? '> Rebuilt from the results already on disk rather than from a fresh ' +
        'comparison. Each case is as its own last run left it. **Differing cells ' +
        "outside the keyed comparison are not counted here** — that total belongs " +
        "to layer 2, which is not in `diff.json`. Each case report has it."
      : '';
  if (note) out.push(note, '');

  out.push(...caseTable(sorted));

  const compared = sum(sorted, (c) => c.tablesCompared);
  const unchecked = sum(sorted, (c) => c.uncheckedDiffering);
  if (compared || unchecked) {
    out.push(
      '',
      `**Totals** — ${n(compared)} table(s) compared, ` +
        `${n(sum(sorted, (c) => c.tablesFailing))} failing, ` +
        `${n(sum(sorted, (c) => c.tablesNotCompared))} not compared` +
        // A rebuild never saw layer 2, so a "0" here would be an answer it does
        // not have rather than the answer.
        (scope.rebuilt ? '.' : `, ${n(unchecked)} differing cell(s) outside the keyed comparison.`),
    );
  }

  const watch = watchNote(sorted);
  if (watch) out.push('', watch);

  out.push(
    '', '---', '',
    '_Per-case detail is in each case\'s `results/report.md`. Every report type at',
    'once is in `run-summary.md`, in the `!summary/` folder at the top of the tree._',
  );
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

  const used = new Set<string>();
  for (const [type, list] of groups) {
    const ws = wb.addWorksheet(sheetName(type, used), { views: [{ state: 'frozen', ySplit: 1 }] });
    caseSheet(ws, list);
  }

  await wb.xlsx.writeFile(path);
}

/**
 * A worksheet name Excel will accept and that no other sheet has taken.
 *
 * Truncation is what makes a clash likely rather than theoretical: two unnamed
 * types both start "Unspecified report type (rep", and ExcelJS throws on the
 * second rather than renaming it, which would take down the whole summary.
 */
function sheetName(type: string, used: Set<string>): string {
  // Excel forbids these in a sheet name, and truncates past 31 characters.
  const base = type.replace(/[\\/*?:[\]]/g, '-').slice(0, 31) || 'Report type';
  let name = base;
  for (let i = 2; used.has(name.toLowerCase()); i++) {
    const tag = ` (${i})`;
    name = base.slice(0, 31 - tag.length) + tag;
  }
  used.add(name.toLowerCase());
  return name;
}

/**
 * The case rows, appended wherever the sheet has got to -- straight onto an
 * empty worksheet in the run summary, or under the title block in a type's own
 * workbook. Read off the sheet rather than passed in, so the styling and the
 * filter cannot end up describing a different row from the header.
 */
function caseSheet(ws: ExcelJS.Worksheet, list: CaseRecord[]): void {
  const headRow = ws.rowCount + 1;
  COLUMNS.forEach((c, i) => { ws.getColumn(i + 1).width = c.width; });
  ws.addRow(COLUMNS.map((c) => c.header));
  styleHead(ws, COLUMNS.length, headRow);

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
  ws.autoFilter = {
    from: { row: headRow, column: 1 },
    to: { row: headRow, column: COLUMNS.length },
  };
}

/**
 * One report type as a workbook of its own: a title, the totals, and the same
 * case rows the run summary gives the type.
 */
export async function writeTypeSummaryWorkbook(
  path: string,
  type: string,
  list: CaseRecord[],
  when = new Date(),
  scope: SummaryScope = {},
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'sheet-verify';
  const sorted = [...list].sort(
    (a, b) => ORDER[a.verdict] - ORDER[b.verdict] || a.name.localeCompare(b.name),
  );

  const ws = wb.addWorksheet(sheetName(type, new Set()), {
    views: [{ state: 'frozen', ySplit: 3 }],
  });
  const need = tally(sorted, 'failed') + tally(sorted, 'could not run');

  ws.getCell('A1').value = `${type} — ${stamp(when)}`;
  ws.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF1F3864' } };
  ws.getRow(1).height = 24;
  ws.getCell('A2').value =
    `${sorted.length} case(s), ${need} needing attention · ` +
    `${n(sum(sorted, (c) => c.tablesCompared))} table(s) compared, ` +
    `${n(sum(sorted, (c) => c.tablesFailing))} failing, ` +
    `${n(sum(sorted, (c) => c.tablesNotCompared))} not compared` +
    // See the markdown: a rebuild never ran layer 2 and cannot answer this.
    (scope.rebuilt
      ? ''
      : `, ${n(sum(sorted, (c) => c.uncheckedDiffering))} differing cell(s) outside the keyed comparison`);
  ws.getCell('A2').font = { color: { argb: 'FF6B7280' } };

  caseSheet(ws, sorted);
  await wb.xlsx.writeFile(path);
}

export interface TypeSummary {
  type: string;
  /** The folder it was written into, which is the report type's own. */
  dir: string;
  markdown: string;
  workbook: string;
  cases: number;
}

export interface SummaryFiles {
  /** The whole run, every report type in one file. */
  markdown: string;
  workbook: string;
  /** One pair per report type, in run order. */
  types: TypeSummary[];
}

/**
 * Every summary file for a run.
 *
 * The run-wide pair goes in `dir`. Each report type's own pair goes beside its
 * cases, in the folder `typeDirs` names for it -- twenty-two files in one
 * folder was more than anyone wanted to scan, and a type's summary is easier
 * to find in the folder already being worked in.
 */
export async function writeSummary(
  dir: string,
  cases: CaseRecord[],
  typeDirs: Map<string, string>,
  when = new Date(),
  scope: SummaryScope = {},
): Promise<SummaryFiles> {
  const markdown = join(dir, `${RUN_SUMMARY}.md`);
  const workbook = join(dir, `${RUN_SUMMARY}.xlsx`);

  // Names are unique within a folder, not across the run, since that is where
  // one file can overwrite another. `run-summary` is reserved wherever the
  // run-wide pair lands, in case a report type is declared at the tree root.
  const taken = new Map<string, Set<string>>([[dir, new Set([RUN_SUMMARY])]]);
  const claim = (into: string, type: string): string => {
    const names = taken.get(into) ?? new Set<string>();
    taken.set(into, names);
    return fileStem(type, names);
  };

  const types: TypeSummary[] = [];
  for (const [type, list] of grouped(cases)) {
    const into = typeDirs.get(type) ?? dir;
    const stem = claim(into, type);
    types.push({
      type,
      dir: into,
      markdown: join(into, `${stem}.md`),
      workbook: join(into, `${stem}.xlsx`),
      cases: list.length,
    });
  }

  // Per folder, because each holds a different set of files this run owns.
  const keep = new Map<string, string[]>([[dir, [markdown, workbook]]]);
  for (const t of types) {
    keep.set(t.dir, [...(keep.get(t.dir) ?? []), t.markdown, t.workbook]);
  }
  for (const [into, files] of keep) {
    await mkdir(into, { recursive: true });
    await clearStale(into, files, scope);
  }

  await writeFile(markdown, summaryMarkdown(cases, when, scope), 'utf8');
  await writeSummaryWorkbook(workbook, cases, when);
  for (const t of types) {
    const list = cases.filter((c) => c.reportType === t.type);
    await writeFile(t.markdown, typeSummaryMarkdown(t.type, list, when, scope), 'utf8');
    await writeTypeSummaryWorkbook(t.workbook, t.type, list, when, scope);
  }

  return { markdown, workbook, types };
}

const RUN_SUMMARY = 'run-summary';

/** Characters Windows will not put in a file name, plus the path separators. */
const FORBIDDEN = /[\\/:*?"<>|]/g;

/**
 * A file name for a report type, unique among the ones already claimed.
 *
 * Case-insensitively unique, because Windows will not hold two files whose
 * names differ only in case: the second write would replace the first, and one
 * report type's summary would silently become another's.
 */
function fileStem(type: string, taken: Set<string>): string {
  const base =
    type.replace(FORBIDDEN, '-').replace(/\s+/g, ' ').trim()
      // Windows drops a trailing dot or space from a file name without saying so.
      .replace(/[. ]+$/, '')
      .slice(0, 60) || 'report-type';
  let name = base;
  for (let i = 2; taken.has(name.toLowerCase()); i++) name = `${base} (${i})`;
  taken.add(name.toLowerCase());
  return name;
}

/**
 * Removes summaries this run is not about to rewrite.
 *
 * A report type that is renamed, or a folder that stops holding cases, leaves
 * a file behind that reads exactly as current as the ones beside it -- the
 * failure this whole tool exists to prevent, in the folder people look at
 * first.
 *
 * Only on a run that covered the whole tree. A scoped run knows which types it
 * touched and nothing about the rest, so clearing there would delete summaries
 * that are still true.
 */
async function clearStale(dir: string, keep: string[], scope: SummaryScope): Promise<void> {
  if (scope.target) return;
  const mine = new Set(keep.map((p) => p.toLowerCase()));
  let entries: string[];
  try {
    entries = (await readdir(dir, { withFileTypes: true }))
      .filter((e) => e.isFile() && /\.(md|xlsx)$/i.test(e.name))
      .map((e) => e.name);
  } catch {
    return; // First run: the folder is not there yet.
  }
  for (const name of entries) {
    const path = join(dir, name);
    if (!mine.has(path.toLowerCase())) await rm(path, { force: true });
  }
}
