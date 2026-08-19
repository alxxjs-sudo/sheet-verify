import { readFile } from 'node:fs/promises';
import { basename, extname } from 'node:path';
import type ExcelJS from 'exceljs';
import { openWorkbook } from './open-xlsx.js';
import { detectWorkbook, type DetectedSheet } from './detect.js';
import { cachedValueState } from './bare.js';

/**
 * Writes the `meta.json` a report type would otherwise be typed out by hand.
 *
 * Everything here is *evidence*, gathered from the pairs themselves, or a
 * comment. Nothing is guessed. The one thing deliberately left out is
 * per-sheet configuration -- header rows, end rows, keys. Those are what a
 * report type genuinely needs a human for, they are the entries that go stale
 * when a report changes shape, and a generated file full of them is
 * indistinguishable from one somebody meant. The unkeyed tables are listed as
 * a note instead, so the work is visible without being pre-decided.
 */

export interface ProposalCase {
  name: string;
  golden: string;
  actual: string;
}

export interface Proposal {
  /** The file to write, comments and all. */
  json: string;
  /** Lines for the console, saying what was found and what is left to do. */
  notes: string[];
}

/**
 * Labels whose value differs between any two runs *by construction*. This list
 * is deliberately short, and deliberately excludes everything that merely
 * sounds like metadata: a creator name is stable when the same account
 * generates every report, so a change there is a finding, not noise. Anything
 * the figures depend on -- view of risk, currency, model version, the as-at
 * date -- is excluded for the opposite reason: if it moved, the numbers under
 * it should have moved too.
 */
const RUN_IDENTITY = [
  'report name', 'report id', 'report number',
  'program name', 'program id',
  'project name', 'project id',
  'creation date', 'created', 'created on', 'created at', 'date created',
  'generated', 'generated on', 'generated at', 'date generated',
  'run id', 'run date', 'run at',
  'last modified date', 'last modified',
  'elapsed processing time',
];

/** Lowercase, collapse whitespace, drop a trailing colon and any quotes. */
const norm = (s: string): string =>
  s.replace(/["']/g, '').replace(/\s+/g, ' ').replace(/:\s*$/, '').trim().toLowerCase();

/** Flattens ExcelJS's several object-shaped cell values to something printable. */
function textOf(cell: ExcelJS.Cell): string {
  const v = cell.value as unknown;
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? '' : v.toISOString();
  if (typeof v === 'object') {
    const o = v as Record<string, unknown>;
    if ('richText' in o) return (o.richText as { text: string }[]).map((t) => t.text).join('');
    if ('text' in o) return String(o.text);
    if ('result' in o) return String(o.result ?? '');
    if ('formula' in o) return `=${String(o.formula)}`;
    return '';
  }
  return String(v);
}

/**
 * A cell's label, and what it labels.
 *
 * Two shapes cover every report here. A pair -- "Report ID" in column A, 4542
 * in column B -- and a fused string, `="Report ID: " & id`, where the label and
 * the value are one cell. The fused form is why the label is looked for inside
 * the text as well as beside it.
 */
interface Found {
  /** The canonical phrase matched, e.g. "report name". */
  phrase: string;
  /** The metadata entry that would match it, `*`-prefixed where it has to be. */
  entry: string;
  /** The label as this report spells it. */
  label: string;
  /** What it read on this side, for the note. */
  value: string;
  sheet: string;
  address: string;
}

/**
 * The run-identity phrase a label carries, if any, and the entry that would
 * match it.
 *
 * A phrase at the *end* of a longer label counts, because every report type
 * spells its own name differently -- Facility Report Name, Pro-Forma Report
 * Name, RiskPlay Report Name. Those get a `*` entry, which is exactly what the
 * wildcard is for: listing each spelling by hand is how a config quietly stops
 * covering the newest one.
 */
function identityIn(label: string): { phrase: string; entry: string } | null {
  const text = norm(label);
  for (const phrase of RUN_IDENTITY) {
    if (text === phrase) return { phrase, entry: label.replace(/:\s*$/, '').trim() };
    if (!text.endsWith(` ${phrase}`)) continue;
    // Keep the report's own casing for the part that matched.
    const words = label.replace(/:\s*$/, '').trim().split(/\s+/);
    return { phrase, entry: `*${words.slice(-phrase.split(' ').length).join(' ')}` };
  }
  return null;
}

function findLabels(ws: ExcelJS.Worksheet): Found[] {
  const out: Found[] = [];
  const width = Math.max(ws.columnCount, 1);

  for (let r = 1; r <= ws.rowCount; r++) {
    // A key-value row is a label and its value, and little else. A header row
    // is wide. Without this, "Program ID" as the 28th column heading of a
    // table reads as run identity, and putting it in `metadata` would drop
    // that whole column out of the comparison.
    let filled = 0;
    for (let c = 1; c <= width; c++) {
      const cell = ws.getRow(r).getCell(c);
      if (cell.isMerged && cell.master !== cell) continue;
      if (textOf(cell).trim()) filled++;
    }
    const keyValueRow = filled <= 4;

    for (let c = 1; c <= width; c++) {
      const cell = ws.getRow(r).getCell(c);
      if (cell.isMerged && cell.master !== cell) continue;
      const text = textOf(cell).trim();
      if (!text) continue;

      // Fused: the label and its value in one cell, so the value is whatever
      // follows the colon.
      const fused = /^(.{2,40}?)\s*:\s*(.+)$/s.exec(text);
      const asFused = fused && identityIn(fused[1]!);
      if (asFused) {
        out.push({
          phrase: asFused.phrase, entry: asFused.entry, label: fused![1]!.trim(),
          value: fused![2]!.trim(), sheet: ws.name, address: cell.address,
        });
        continue;
      }

      // A fused string carries its own value and is unambiguous wherever it
      // sits. A bare label is only a label if its row is a key-value row.
      if (!keyValueRow) continue;

      const hit = identityIn(text);
      if (!hit) continue;

      // A pair: the value is the cell to the right. Only the next column -- a
      // label with blank cells after it and a number further along is a table,
      // not a pair.
      const beside = c + 1 <= width ? textOf(ws.getRow(r).getCell(c + 1)).trim() : '';
      out.push({
        phrase: hit.phrase, entry: hit.entry, label: text.replace(/:\s*$/, '').trim(),
        value: beside, sheet: ws.name, address: cell.address,
      });
    }
  }
  return out;
}

/** Two spellings of the same phrase are one candidate. */
const labelKey = (f: Found) => f.phrase;

interface Candidate {
  /** As the report spells it, first spelling seen. */
  label: string;
  /** The entry to write. A `*` form when the label carries a report's own name. */
  entry: string;
  /** Every spelling seen, listed in the note when they disagree. */
  spellings: Set<string>;
  /**
   * Every distinct value seen, across both sides of every case scanned.
   *
   * More than one is the evidence. The test for run identity is that the value
   * differs between two runs *by construction*, and a report type's cases are
   * exactly that -- several runs of the same generator. Comparing only the two
   * sides of one pair is the wrong question and usually finds nothing: a
   * golden and the report judged against it are frequently the same download.
   */
  values: Set<string>;
  /** Stronger still: the two sides of one pair disagreed. */
  differsWithinPair: boolean;
  where: string;
}

/**
 * Whether two columns of a sheet look like a group heading written once and
 * left blank beneath -- the shape `fillKeyDown` exists for. Reported per sheet
 * so the proposal can say which sheet the evidence came from rather than
 * asserting the setting out of nowhere.
 */
function groupedColumn(sheets: DetectedSheet[]): string | null {
  for (const s of sheets) {
    for (const t of s.tables) {
      if (t.rows < 8) continue; // too short for the pattern to mean anything
      for (const [i, name] of t.headers.entries()) {
        if (!name) continue;
        const filled = t.filledPerColumn[i];
        if (filled === undefined) continue;
        const ratio = filled / t.rows;
        // Written on a few rows out of many, and on the first row of the
        // table: a heading, not a column that merely has gaps in it.
        if (ratio > 0.01 && ratio < 0.5 && t.firstRowFilled[i]) {
          return `${s.sheet} · ${t.name}, column "${name}" (${filled} of ${t.rows} rows)`;
        }
      }
    }
  }
  return null;
}

/**
 * A column name fit for one line of a note. Report headers carry hard line
 * breaks for layout and run to whole sentences -- one of them here is a
 * 230-character A.M. Best instruction -- and pasted raw they turn a list into
 * a wall.
 */
const oneLine = (s: string, max = 34): string => {
  const flat = s.replace(/\s+/g, ' ').trim();
  return flat.length <= max ? flat : `${flat.slice(0, max - 1)}…`;
};

/** Folder name to something readable. The spelling is the author's to fix. */
function readableType(folder: string): string {
  return folder
    .split(/[-_\s]+/)
    .filter(Boolean)
    .map((w) => w[0]!.toUpperCase() + w.slice(1))
    .join(' ');
}

const isSpreadsheet = (p: string) => ['.xlsx', '.xlsm'].includes(extname(p).toLowerCase());

export async function proposeMeta(
  typeFolder: string,
  cases: ProposalCase[],
): Promise<Proposal> {
  const candidates = new Map<string, Candidate>();
  const notes: string[] = [];

  let bareFiles = 0;
  let readFiles = 0;
  let grouped: string | null = null;
  const unkeyed: { label: string; rows: number; columns: string[] }[] = [];

  for (const c of cases) {
    if (!isSpreadsheet(c.golden) || !isSpreadsheet(c.actual)) continue;

    const [goldenWb, actualWb] = await Promise.all([
      openWorkbook(c.golden), openWorkbook(c.actual),
    ]);
    readFiles += 2;

    // Bare files are the norm for generated reports and the reason
    // requireCachedValues has to be turned off. Asked of the files rather than
    // assumed, so the proposal only carries it when it is true.
    for (const p of [c.golden, c.actual]) {
      const state = await cachedValueState(await readFile(p));
      if (state.bare > 0 && state.cached === 0) bareFiles++;
    }

    const onSide = (wb: ExcelJS.Workbook) => {
      const found = new Map<string, Found>();
      for (const ws of wb.worksheets) {
        for (const f of findLabels(ws)) if (!found.has(labelKey(f))) found.set(labelKey(f), f);
      }
      return found;
    };

    const before = onSide(goldenWb);
    const after = onSide(actualWb);

    for (const [key, f] of [...before, ...after]) {
      const existing = candidates.get(key) ?? {
        label: f.label,
        entry: f.entry,
        spellings: new Set<string>(),
        values: new Set<string>(),
        differsWithinPair: false,
        where: `${f.sheet}!${f.address}`,
      };
      existing.spellings.add(f.label);
      if (f.value) existing.values.add(f.value);

      const g = before.get(key);
      const a = after.get(key);
      if (g && a && g.value !== a.value) existing.differsWithinPair = true;
      candidates.set(key, existing);
    }

    if (!grouped || !unkeyed.length) {
      const detected = await detectWorkbook(c.golden);
      grouped ??= groupedColumn(detected);
      if (!unkeyed.length) {
        for (const s of detected) {
          for (const t of s.tables) {
            if (t.keyColumns || t.rows < 5) continue;
            unkeyed.push({
              label: s.tables.length === 1 ? s.sheet : `${s.sheet} · ${t.name}`,
              rows: t.rows,
              columns: t.headers.filter(Boolean).slice(0, 4),
            });
          }
        }
      }
    }
  }

  const confirmed = [...candidates.values()].filter((c) => c.values.size > 1);
  const seenOnly = [...candidates.values()].filter((c) => c.values.size <= 1);

  const entryFor = (c: Candidate): string => c.entry;

  const meta: Record<string, unknown> = {};
  meta['//'] = [
    `Written by sheet-verify --write-meta from ${cases.length} case(s).`,
    'Everything in it came from the files themselves. Read it, fix the report',
    'type spelling, and keep it under version control.',
    '',
    'Deliberately absent: per-sheet header rows, end rows and keys. Those are',
    'detected fresh from every file on every run, which is what lets a report',
    'change shape without breaking the config. Add one only where a run tells',
    'you detection got it wrong -- see docs/detection-tuning.md.',
  ];

  meta.reportType = readableType(basename(typeFolder));
  meta['//reportType'] = 'Taken from the folder name. Rewrite it however you say it out loud.';

  if (confirmed.length) {
    meta['//metadata'] = [
      'Cells that identify the run rather than describe it. Comparing them',
      'reports a difference on every run, so no run is ever clean and the',
      'first section stops being read. These are read, listed in report.md',
      'with both values, and left out of the verdict.',
      '',
      'Every entry below was OBSERVED to take more than one value across the',
      'files scanned, which is the test: a value that differs between two runs',
      'by construction. The two values shown are the first two seen.',
      ...confirmed.map((c) => {
        const [a, b] = [...c.values].slice(0, 2).map((v) => oneLine(v, 46));
        const pair = b ? `: ${a} / ${b}` : `: ${a}`;
        const same = c.differsWithinPair ? '' : ' (same on both sides of every pair)';
        return `  ${entryFor(c)} — ${c.where}${pair}${same}`;
      }),
      '',
      'Do NOT add a creator name here: the same account generates every report,',
      'so a change means the wrong account ran it. Nor anything the figures',
      'depend on -- view of risk, currency, model version, the as-at date. If',
      'one of those moved, the numbers under it should have moved too.',
    ];
    meta.metadata = confirmed.map(entryFor);
  }

  if (seenOnly.length) {
    meta['//metadata-candidates'] = [
      'Found in the files, but holding the same value in every file scanned,',
      'so there is no evidence they vary. Left out. Move one into "metadata"',
      'above if you know it is minted fresh on every run and these cases just',
      'happen not to show it:',
      ...seenOnly.map((c) => `  ${entryFor(c)} — ${c.where}`),
    ];
  }

  const defaults: Record<string, unknown> = {};
  const defaultNotes: string[] = [];

  if (bareFiles > 0) {
    defaults.requireCachedValues = false;
    defaultNotes.push(
      `requireCachedValues is off because ${bareFiles} of ${readFiles} file(s) scanned`,
      'carry formulas with no stored result -- Excel works them out on open.',
      'Requiring them would fail every case on the way the files are produced.',
    );
  }
  if (grouped) {
    defaults.fillKeyDown = true;
    defaultNotes.push(
      ...(defaultNotes.length ? [''] : []),
      'fillKeyDown carries a key value down the blank cells beneath it, which is',
      'how these reports write a grouped breakdown: the group is named once and',
      'the rows under it leave the column empty. Evidence:',
      `  ${grouped}`,
      'Only key building is affected. No cell value is invented, and nothing',
      'filled this way is ever compared.',
    );
  }
  if (Object.keys(defaults).length) {
    meta['//defaults'] = defaultNotes;
    meta.defaults = defaults;
  }

  if (unkeyed.length) {
    const worst = unkeyed.sort((a, b) => b.rows - a.rows).slice(0, 10);
    meta['//keys'] = [
      `${unkeyed.length} table(s) have no column that identifies a row, so their`,
      'rows are paired by order. That is exact while both sides hold the same',
      'rows; an inserted row shifts the rest, so one change reads as many.',
      '',
      'Nothing is written for them here, because guessing a key is worse than',
      'having none. Name the ones worth pinning down, largest first:',
      ...worst.map((t) => `  ${t.label} — ${t.rows} rows — ${t.columns.map((c) => oneLine(c)).join(', ')}`),
      ...(unkeyed.length > worst.length ? [`  … and ${unkeyed.length - worst.length} more`] : []),
      '',
      'Then write, per sheet:',
      '  "sheets": { "<sheet>": { "tables": { "<table>": { "keyColumns": [...] } } } }',
    ];
  }

  notes.push(`read ${readFiles} file(s) across ${cases.length} case(s)`);
  notes.push(
    confirmed.length
      ? `metadata: ${confirmed.length} label(s) observed varying — ${confirmed.map(entryFor).join(', ')}`
      : 'metadata: nothing observed varying across these files; the list is left out',
  );
  if (seenOnly.length) notes.push(`  ${seenOnly.length} more found but unvarying, listed as candidates`);
  if (Object.keys(defaults).length) notes.push(`defaults: ${Object.keys(defaults).join(', ')}`);
  if (unkeyed.length) notes.push(`${unkeyed.length} table(s) have no row key — listed in the file, none guessed at`);

  return { json: `${JSON.stringify(meta, null, 2)}\n`, notes };
}
