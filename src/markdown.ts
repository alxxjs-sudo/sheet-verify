import type { SheetOutcome, WorkbookDiffResult, CellValue, ValueDiff } from './types.js';
import type { SweepResult } from './sweep.js';

/**
 * One report per case, in Markdown.
 *
 * There used to be two files: what changed, and what nothing looked at. They
 * are two halves of one answer -- a change and the totals it moves are the
 * same story, and so is a passing table and the fact that nothing checked it
 * -- so reading them meant holding two documents side by side. This is both,
 * in the order someone reads them: the verdict, then what changed, then what
 * that will do, then what was not covered.
 *
 * Nothing is truncated. A report exists to be complete, and "and 34 more"
 * hides exactly the row someone was looking for. A terminal has a reason to
 * elide; a file does not.
 */

export interface MarkdownOptions {
  /** Case name, taken from the folder. Titles the report when nothing better. */
  name: string;
  /**
   * What this case is for. Titles the report when set, with the folder name
   * kept underneath so the file can still be found from what it says.
   */
  label?: string;
  /** The kind of report being compared, shown beside the case name. */
  reportType?: string;
  /** Separator used in composite keys, replaced with " / " for display. */
  keySeparator?: string;
}

/**
 * A number as text, at the precision it is stored.
 *
 * This used to trim to 12 significant digits, which read better -- until the
 * comparison started reporting gaps below that. Trimming then printed the two
 * sides of a difference as the same string, which is worse than ugly: it makes
 * the report look wrong. JavaScript's default is the shortest text that reads
 * back as the identical number, so nothing is invented and nothing is lost.
 */
const num = (v: number): string => String(v);

/** Counts read better grouped once they run to five figures. */
const n = (v: number): string => v.toLocaleString('en-US');

const BLANK = '_(blank)_';

/**
 * What differs, as its own table.
 *
 * One number cannot carry this. "894" reads as a disaster when 870 of those
 * moved by less than a thousandth; "24" alone hides that the file is full of
 * recalculation drift. Three columns and a share, with the tolerance that drew
 * the line named on the column that used it -- so the figure everyone acts on
 * can be read without scrolling and without arithmetic.
 */
function cellCounts(s: SweepResult): string[] {
  const above = s.totalDifferences;
  const total = above + s.totalTolerated;

  const share = total === 0 ? 0 : (above / total) * 100;
  const pct = above === 0 ? '0%'
    : share >= 99.95 ? '100%'
      : share < 0.1 ? '<0.1%'
        : `${share.toFixed(1)}%`;

  // The tolerance is resolved per column, so a run can apply several. One value
  // is named outright; a spread is given as a range, since claiming a single
  // number would be false.
  const applied = [...new Set(s.tolerated.map((c) => c.tolerance ?? 0))].sort((a, b) => a - b);
  const band = applied.length === 0 ? ''
    : applied.length === 1 ? ` (±${num(applied[0]!)})`
      : ` (±${num(applied[0]!)}–${num(applied[applied.length - 1]!)})`;

  return [
    '',
    '**Cells that differ**',
    '',
    `| total | within tolerance${band} | above tolerance |`,
    '| ---: | ---: | ---: |',
    `| ${n(total)} | ${n(s.totalTolerated)} | **${n(above)} (${pct})** |`,
  ];
}

/**
 * The same three counts for one table, above its findings.
 *
 * "Value changes (3)" reads very differently depending on whether the other
 * four hundred cells held still or drifted a hair each. The counts are the
 * table's own: what layer 1 reported here, and what the tolerance absorbed
 * here -- so they agree with the lists directly underneath.
 *
 * Emitted under every changed table. Where the tolerance forgave nothing the
 * row reads "3 of 3, 100% above", which is the point: a reader comparing two
 * tables should not have to work out whether a missing block means no drift or
 * no data.
 */
function tableCounts(o: SheetOutcome, s: SweepResult): string[] {
  const within = s.tolerated.filter((c) => c.table === o.label);

  const d = o.diff!;
  // One exception to "under every table": a comparison that never ran. A table
  // whose key column was not found has no rows compared, and a row of zeros
  // there would say its cells were checked and matched -- the opposite of what
  // the integrity error underneath is about to say.
  if (d.rows.compared === 0 && within.length === 0) return [];

  const above = d.values.length + d.formulas.length + d.types.length;
  const total = above + within.length;
  const share = total === 0 ? 0 : (above / total) * 100;
  const pct = above === 0 ? '0%'
    : share >= 99.95 ? '100%'
      : share < 0.1 ? '<0.1%'
        : `${share.toFixed(1)}%`;

  // The tolerance is named from the cells it actually forgave. With none to
  // read it from, the column says what it counts and claims no number --
  // guessing one from config would name a tolerance that never applied here.
  const applied = [...new Set(within.map((c) => c.tolerance ?? 0))].sort((a, b) => a - b);
  const band = applied.length === 0 ? ''
    : applied.length === 1
      ? ` (±${num(applied[0]!)})`
      : ` (±${num(applied[0]!)}–${num(applied[applied.length - 1]!)})`;

  return [
    '',
    `| total | within tolerance${band} | above tolerance |`,
    '| ---: | ---: | ---: |',
    `| ${n(total)} | ${n(within.length)} | **${n(above)} (${pct})** |`,
  ];
}

/**
 * The size of a tolerated gap, for the reader to judge the tolerance by. Both
 * sides parsed as numbers to get here, so this only has to render one.
 */
function gapBetween(base: string, next: string): string {
  const a = Number(base);
  const b = Number(next);
  return Number.isFinite(a) && Number.isFinite(b) ? num(Math.abs(a - b)) : '';
}

const show = (v: CellValue): string => {
  if (v === null || v === undefined) return BLANK;
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? BLANK : v.toISOString().slice(0, 10);
  if (typeof v === 'number') return num(v);
  if (typeof v === 'string') return v === '' ? BLANK : v;
  return String(v);
};

/** A pipe inside a cell would end the column early, and a newline the row. */
const cell = (s: string): string => s.replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');

/** Addresses and formulas read better fixed-width, and never wrap oddly. */
const code = (s: string | null | undefined): string =>
  s === null || s === undefined || s === '' ? '' : `\`${cell(s).replace(/`/g, "'")}\``;

type Align = 'left' | 'right';

function table(headers: string[], rows: string[][], align?: Align[]): string[] {
  if (!rows.length) return [];
  const rule = (i: number) => (align?.[i] === 'right' ? '---:' : '---');
  return [
    `| ${headers.join(' | ')} |`,
    `| ${headers.map((_, i) => rule(i)).join(' | ')} |`,
    ...rows.map((r) => `| ${r.join(' | ')} |`),
  ];
}

/**
 * Where a table is and how it was matched, in one line.
 *
 * The range is the rectangle layer 1 read, not the one the spec asked for: a
 * spec with no `endRow` runs to the bottom of the sheet and answers nothing.
 * When the two files disagree about it, both are shown -- a table that grew,
 * moved or lost its bottom rows between runs is worth seeing as such, and it
 * explains a wall of differences that would otherwise look like edits.
 */
function coverage(s: SheetOutcome): string {
  if (!s.range?.base) return '';
  const where = s.range.next && s.range.next !== s.range.base
    ? `${code(s.range.base)} in the golden, ${code(s.range.next)} in the report`
    : code(s.range.base);
  const plural = (count: number, noun: string) =>
    `${n(count)} ${noun}${count === 1 ? '' : 's'}`;
  const shape = s.diff
    ? ` — ${plural(s.diff.schema.compared.length, 'column')}` +
      ` × ${plural(s.diff.rows.compared, 'row')}, ` +
      `rows matched by ${s.reason ? 'position' : 'key'}`
    : '';
  return `_${where}${shape}_`;
}

/**
 * How many differences a table can list flat before it needs a summary above
 * them. Twenty fits on a screen; nine hundred does not.
 */
const GROUP_ABOVE = 12;

/**
 * The same differences, one row per column instead of one row per cell.
 *
 * A recalculated report drifts in the last digit of every total it recomputes,
 * so a single sheet can carry hundreds of differences of which two matter. Flat,
 * that is a wall nobody reads to the end of; by column it is a handful of rows,
 * and the largest delta in each says immediately which pile is worth opening.
 * One real case: 894 differences, 27 columns, and the 4,700,000 hiding among
 * them was in a group whose other 43 entries were all below 1e-10.
 */
function byColumn(list: ValueDiff[]): string[] {
  interface Group { cells: number; min?: number; max?: number; at?: string }
  const groups = new Map<string, Group>();

  for (const v of list) {
    const g = groups.get(v.column) ?? { cells: 0 };
    g.cells++;
    if (typeof v.delta === 'number') {
      if (g.min === undefined || v.delta < g.min) g.min = v.delta;
      if (g.max === undefined || v.delta > g.max) { g.max = v.delta; g.at = v.address; }
    }
    groups.set(v.column, g);
  }

  const rows = [...groups].sort((a, b) => (b[1].max ?? 0) - (a[1].max ?? 0) || b[1].cells - a[1].cells);
  return [
    '',
    `By column — ${rows.length} column(s), largest difference first.`,
    '',
    ...table(
      ['Column', 'Cells', 'Smallest Δ', 'Largest Δ', 'Largest at'],
      rows.map(([column, g]) => [
        cell(column), n(g.cells),
        g.min === undefined ? '' : num(g.min),
        g.max === undefined ? '' : num(g.max),
        code(g.at),
      ]),
    ),
  ];
}

/** One compared table's findings. */
function findings(o: SheetOutcome, sep: string): string[] {
  const d = o.diff!;
  const out: string[] = [];
  const key = (k: string) => cell(k.split(sep).join(' / '));

  if (d.errors.length) {
    out.push('', '**Comparison integrity** — these make the result untrustworthy', '');
    for (const e of d.errors) out.push(`- ${cell(e)}`);
  }

  if (d.formulas.length) {
    out.push('', `**Formula changes (${d.formulas.length})** — the calculation itself differs`, '');
    out.push(...table(
      ['Row', 'Column', 'Cell', 'Golden', 'Actual'],
      d.formulas.map((f) => [
        key(f.key), cell(f.column), code(f.address), code(f.baseA1), code(f.nextA1),
      ]),
    ));
  }

  const roots = d.values.filter((v) => v.rootCause);
  const cascades = d.values.filter((v) => !v.rootCause);
  const groups: [string, typeof roots][] = [
    [`Value changes (${roots.length})`, roots],
    [`Cascaded value changes (${cascades.length}) — downstream of the above, in this table`, cascades],
  ];
  for (const [label, list] of groups) {
    if (!list.length) continue;
    out.push('', `**${label}**`, '');

    const big = list.length > GROUP_ABOVE;
    if (big) out.push(...byColumn(list));

    const rows = table(
      ['Row', 'Column', 'Cell', 'Golden', 'Actual', 'Delta'],
      list.map((v) => [
        key(v.key), cell(v.column), code(v.address), cell(show(v.base)), cell(show(v.next)),
        v.delta === undefined ? '' : num(v.delta),
      ]),
    );

    // Every cell is still here, one per row, exactly as before -- folded away
    // rather than dropped. A summary that replaced the detail would just be a
    // slower way of truncating.
    if (big) out.push('', `<details><summary>All ${n(list.length)} cells</summary>`, '', ...rows, '', '</details>');
    else out.push(...rows);
  }

  if (d.types.length) {
    out.push('', `**Type changes (${d.types.length})** — same rendering, different type`, '');
    out.push(...table(
      ['Row', 'Column', 'Cell', 'Golden', 'Actual', 'Value'],
      d.types.map((t) => [
        key(t.key), cell(t.column), code(t.address), t.baseKind, t.nextKind, cell(show(t.value)),
      ]),
    ));
  }

  if (d.invariants.length) {
    out.push('', `**Invariant failures (${d.invariants.length})** — wrong regardless of the golden`, '');
    out.push(...table(
      ['Invariant', 'Row', 'Column', 'Cell', 'Detail'],
      d.invariants.map((i) => [
        cell(i.invariant), i.key ? key(i.key) : '', cell(i.column ?? ''),
        code(i.address), cell(i.detail),
      ]),
    ));
  }

  const { added, removed, moved } = d.schema;
  if (added.length || removed.length || moved.length) {
    out.push('', '**Column changes** — review, then re-bless the golden if intended', '');
    for (const c of added) out.push(`- added: ${code(c)}`);
    for (const c of removed) out.push(`- removed: ${code(c)}`);
    for (const m of moved) out.push(`- moved: ${code(m.column)} from ${m.from} to ${m.to}`);
  }

  if (d.rows.added.length || d.rows.removed.length) {
    out.push('', '**Row population**', '');
    if (d.rows.added.length) {
      out.push(`- ${d.rows.added.length} added: ${d.rows.added.map((k) => code(key(k))).join(', ')}`);
    }
    if (d.rows.removed.length) {
      out.push(`- ${d.rows.removed.length} removed: ${d.rows.removed.map((k) => code(key(k))).join(', ')}`);
    }
  }

  const repeated = [...new Set([...d.rows.duplicateKeysBase, ...d.rows.duplicateKeysNext])];
  if (repeated.length) {
    out.push('', `**Repeated row keys (${repeated.length})** — matched in order of appearance`, '');
    out.push('These rows carry no key of their own, a per-group total line for instance.');
    out.push('The first in the golden is compared with the first here, which holds while');
    out.push('the groups stay in the same order.');
    out.push('');
    out.push(repeated.map((k) => code(key(k) || '(blank)')).join(', '));
  }

  return out;
}

/**
 * What was actually done, in two sentences, where the old report had a bare
 * "cells checked 33,202 of 33,425".
 *
 * That line was read as "and 223 went unlooked-at", which was never true --
 * they were compared by the second layer instead of the first. A coverage
 * fraction only means something once the reader knows there are two layers
 * with different reach, so this says that first and quantifies second.
 */
function assurance(s: SweepResult): string[] {
  const oneSided = s.sheets.filter((x) => x.status !== 'swept');
  const oneSidedCells = oneSided.reduce((t, x) => t + x.cells, 0);
  const judged = Math.max(s.cellsSwept - s.metadata.length, 0);
  const byAddress = Math.max(judged - s.cellsCompared, 0);

  const out = ['', '**Two-layer verification — both layers ran over every shared sheet.**', ''];
  out.push(
    `- **Layer 1, by name and key** — ${n(s.cellsCompared)} cells, across tables whose`,
    '  columns were paired by header name and rows by the values that identify them,',
    '  so a column that moved or a row that shifted is still held against its own',
    '  counterpart. Formula text, stored value, value type and invariants.',
    `- **Layer 2, by address** — ${n(judged)} cells, every one in both files compared`,
    '  A1 against A1. A literal counts as its value, a formula as its text together',
    '  with any stored result, so a changed calculation is caught even where no',
    '  result was ever saved. It needs no keys and no configuration, which is what',
    `  lets it reach title blocks, notes, and every table layer 1 could not key:`,
    `  ${n(byAddress)} cells rest on this layer alone.`,
  );

  out.push('', 'Between them, every cell on a sheet both files share was compared.');
  if (oneSidedCells > 0) {
    const many = oneSided.length !== 1;
    out.push(
      `A further ${n(oneSidedCells)} cells sit on ${n(oneSided.length)}`,
      `${many ? 'sheets that only one file has' : 'sheet that only one file has'}, so there is`,
      `nothing to hold them against — ${many ? 'those sheets' : 'that sheet'} appearing or`,
      'disappearing is itself the finding, and is listed below.',
    );
  }
  const meta = s.metadata.length ? `${n(s.metadata.length)} cells of report metadata, and ` : '';
  out.push(
    '',
    `Read but not judged: ${meta}any sheet excluded on purpose. Both are listed at`,
    'the end of this report with their values, so nothing is set aside out of sight.',
  );
  if (s.totalTolerated > 0) {
    out.push(
      '',
      `Both layers apply the tolerances configured for a column: ${n(s.totalTolerated)} cells`,
      'differ by less than the one set for theirs, and are listed on their own rather',
      'than counted as differences.',
    );
  }
  return out;
}
/**
 * The disclaimer, and the evidence behind it.
 *
 * Two things are deliberately left out of the verdict, and a reader is owed
 * both the reason and the contents. Skipping a comparison silently is how a
 * tool loses the right to say "identical" -- so everything skipped is listed
 * here with its values, and a change nobody expected is still there to be
 * seen. It just does not fail the run.
 */
function notVerified(diff: WorkbookDiffResult, swept: SweepResult | null): string[] {
  const ignored = diff.sheets.filter((s) => s.status === 'ignored');
  const meta = swept?.metadata ?? [];
  if (!ignored.length && !meta.length) return [];

  const out = ['', '## Not verified, on purpose', ''];
  out.push('Everything below was read. None of it counts towards the verdict.');

  if (meta.length) {
    const differing = meta.filter((m) => m.base !== m.next);
    const same = meta.filter((m) => m.base === m.next);

    out.push('', `### Report metadata (${n(meta.length)} cells)`, '');
    out.push('A report name, a report id, whoever generated it and when: these');
    out.push('differ between any two runs by construction. Comparing them reports a');
    out.push('difference every single time, which makes a clean run impossible and');
    out.push('trains a reader to skip the first section.');
    out.push('');
    out.push('Cells matching the configured `metadata` patterns are therefore read');
    out.push('and set aside, and nothing that reads them is chased downstream.');
    out.push('');
    out.push('This is a judgement about identity, not about content. Anything the');
    out.push('figures depend on — view of risk, currency, model version, the as-at');
    out.push('date of the data — is compared normally: if one of those moved, the');
    out.push('numbers under it should have moved too.');

    if (differing.length) {
      out.push('', `**Differing (${n(differing.length)})** — expected, and shown so you can still spot one that is not.`, '');
      out.push(...table(
        ['Sheet', 'Cell', 'Matched', 'Golden', 'Actual'],
        differing.map((m) => [
          cell(m.sheet), code(m.address), code(m.rule ?? ''),
          cell(m.base || BLANK), cell(m.next || BLANK),
        ]),
      ));
    }

    if (same.length) {
      out.push('', `**Identical (${n(same.length)})**`, '');
      out.push('<details><summary>Show</summary>', '');
      out.push(...table(
        ['Sheet', 'Cell', 'Matched', 'Value'],
        same.map((m) => [
          cell(m.sheet), code(m.address), code(m.rule ?? ''), cell(m.base || BLANK),
        ]),
      ));
      out.push('', '</details>');
    }
  }

  if (ignored.length) {
    out.push('', `### Excluded sheets (${ignored.length})`, '');
    out.push('Named in `ignoreSheets`, so nothing on them is judged. Their cells are');
    out.push('still swept, and the count below says whether anything moved — a');
    out.push('non-zero number on a sheet you thought was boilerplate is worth a look.');
    out.push('');
    const by = new Map((swept?.sheets ?? []).map((s) => [s.sheet.trim().toLowerCase(), s]));
    out.push(...table(
      ['Sheet', 'Cells', 'Differing'],
      ignored.map((s) => {
        const sw = by.get(s.sheet.trim().toLowerCase());
        return [cell(s.label), sw ? n(sw.cells) : '—', sw ? n(sw.differing) : '—'];
      }),
    ));
  }

  return out;
}

export function formatMarkdownReport(
  diff: WorkbookDiffResult,
  swept: SweepResult | null,
  options: MarkdownOptions,
): string {
  const sep = options.keySeparator ?? '␟';
  const out: string[] = [];

  const compared = diff.sheets.filter((s) => s.status === 'compared');
  const failed = compared.filter((s) => !s.diff!.ok);
  const positional = compared.filter((s) => s.reason);
  const skipped = diff.sheets.filter((s) => s.status === 'skipped');

  const verdict = !diff.ok
    ? '**Differences found.**'
    : diff.reviewOnly
      ? '**No defects.** Something changed — review below.'
      : '**Identical.**';

  // The label titles the report when there is one, since "a peril column added
  // between two others" tells a reader what they are looking at and "case_003"
  // does not. The folder name still has to appear -- it is how the case is
  // found on disk and named in the log -- so it moves to the line under it,
  // with the report type it belongs to.
  // A label repeating the folder name is treated as no label at all, so the
  // subtitle does not print the same words back under the title.
  const label = options.label && options.label !== options.name ? options.label : '';
  out.push(`# ${cell(label || options.name)}`, '');
  const identity = [options.reportType, label ? options.name : '']
    .filter(Boolean)
    .join(' · ');
  if (identity) out.push(`_${cell(identity)}_`, '');
  out.push(verdict, '');

  out.push(...table(['', ''], [
    ['golden', `${code(diff.base.source)} — ${diff.base.sheets.length} sheet(s)`],
    ['report', `${code(diff.next.source)} — ${diff.next.sheets.length} sheet(s)`],
    ['tables compared', `${compared.length}${positional.length ? `, ${positional.length} by row position` : ''}`],
    ...(skipped.length ? [['tables not compared', String(skipped.length)]] : []),
  ]));

  // The count of what differs is the reason anyone opened the file, so it gets
  // a table of its own rather than a row in the middle of the file paths.
  if (swept) out.push(...cellCounts(swept));

  if (swept) out.push(...assurance(swept));

  if (diff.errors.length) {
    out.push('', '## Workbook integrity', '', 'These make the result untrustworthy.', '');
    for (const e of diff.errors) out.push(`- ${cell(e)}`);
  }

  if (diff.sheetSchema.removed.length) {
    out.push('', `## Sheets removed (${diff.sheetSchema.removed.length})`, '',
      'Output that is no longer produced.', '');
    for (const s of diff.sheetSchema.removed) out.push(`- ${cell(s)}`);
  }

  if (failed.length) {
    out.push('', '## What changed', '');
    for (const s of failed) {
      out.push('', `### ${cell(s.label)}`);
      // Where the table is, before what is wrong with it. A finding at B25 is
      // read differently depending on whether the table starts at row 2 or 24.
      const where = coverage(s);
      if (where) out.push('', where);
      // Every changed table gets one, whether or not the tolerance forgave
      // anything here: the same three columns in the same place under every
      // heading is what makes them comparable at a glance.
      if (swept) out.push(...tableCounts(s, swept));
      out.push(...findings(s, sep));
    }
  }

  if (swept && swept.totalAffected > 0) {
    out.push('', `## Will recalculate differently (${swept.totalAffected})`, '');
    out.push('These hold formulas whose text has not changed, so nothing above reports');
    out.push('them. They read a cell that did change, so Excel produces a different');
    out.push('number here the moment the file is opened. No value is shown because');
    out.push('neither file stores one.');
    out.push('');

    // Grouped by the column they sit in, not just the sheet. A list of ninety
    // addresses says where to look and nothing about what moved; "Geocoded
    // Sums Insured, 8 cells" says which figure is about to change, which is
    // the question someone reading this actually has.
    const NO_COLUMN = '—';
    const groups = new Map<string, { sheet: string; column: string; cells: string[] }>();
    for (const a of swept.affected) {
      const column = a.column ?? NO_COLUMN;
      const k = `${a.sheet}\u0000${column}`;
      const g = groups.get(k) ?? { sheet: a.sheet, column, cells: [] };
      g.cells.push(a.address);
      groups.set(k, g);
    }

    const rowOf = (a: string) => Number(a.replace(/^[A-Z]+/, ''));
    const groupRows = [...groups.values()].sort(
      (a, b) => a.sheet.localeCompare(b.sheet)
        || b.cells.length - a.cells.length
        || a.column.localeCompare(b.column),
    );
    for (const g of groupRows) g.cells.sort((x, y) => rowOf(x) - rowOf(y));

    // Two hundred addresses on one line is not a grouping, it is the same wall
    // turned sideways. A span says as much and stays readable; the full list
    // is below, folded.
    const SPAN_ABOVE = 8;
    const where = (cells: string[]) =>
      cells.length <= SPAN_ABOVE
        ? cells.map((c) => code(c)).join(' ')
        : `${code(cells[0])} … ${code(cells[cells.length - 1])}`;

    // One table per sheet, rather than one table with the sheet repeated down
    // it. Markdown has no per-row border, so whitespace is the only separator
    // available -- and once the sheet heads its own block, the column that
    // carried it is redundant and the table gets narrower for free.
    const sheets: string[] = [];
    for (const g of groupRows) if (!sheets.includes(g.sheet)) sheets.push(g.sheet);

    for (const sheet of sheets) {
      const mine = groupRows.filter((g) => g.sheet === sheet);
      const cells = mine.reduce((t, g) => t + g.cells.length, 0);
      out.push('', `**${cell(sheet)}** — ${n(cells)} cell(s)`, '');
      out.push(...table(
        ['Column', 'Cells', 'Where'],
        mine.map((g) => [cell(g.column), n(g.cells.length), where(g.cells)]),
      ));
    }
    if (groupRows.some((g) => g.column === NO_COLUMN)) {
      out.push('', `${NO_COLUMN} — outside any table that was compared by name, so no header names it.`);
    }

    if (groupRows.some((g) => g.cells.length > SPAN_ABOVE)) {
      out.push('', `<details><summary>All ${n(swept.affected.length)} cells</summary>`, '');
      for (const g of groupRows) {
        out.push('', `**${cell(g.sheet)} · ${cell(g.column)}** (${n(g.cells.length)})`, '');
        out.push(g.cells.map((c) => code(c)).join(' '));
      }
      out.push('', '</details>');
    }
  }

  if (swept && swept.totalGaps > 0) {
    out.push('', `## Differing, and nothing checked them (${swept.totalGaps})`, '');
    out.push('Found by sweeping every cell by address. The keyed comparison did not');
    out.push('reach these, so they would otherwise have gone unreported.');
    out.push('');

    // Grouped by sheet, as "Will recalculate differently" is: one block each,
    // separated by whitespace, and no column spent repeating the sheet name
    // down every row of it.
    const gaps = swept.differences.filter((d) => d.status === 'gap');
    const sheetsWithGaps: string[] = [];
    for (const d of gaps) if (!sheetsWithGaps.includes(d.sheet)) sheetsWithGaps.push(d.sheet);

    for (const sheet of sheetsWithGaps) {
      const mine = gaps.filter((d) => d.sheet === sheet);
      out.push('', `**${cell(sheet)}** — ${n(mine.length)} cell(s)`, '');
      out.push(...table(
        ['Cell', 'Golden', 'Actual', 'Why'],
        mine.map((d) => [
          code(d.address), cell(d.base || BLANK), cell(d.next || BLANK), d.reason ?? '',
        ]),
      ));
    }
  }

  if (swept && swept.totalTolerated > 0) {
    out.push('', `## Inside the tolerance you set (${n(swept.totalTolerated)})`, '');
    out.push('These cells hold different numbers, by less than the tolerance configured');
    out.push('for their column. They are not counted as differences and do not affect');
    out.push('the verdict — the tolerance is the statement that a gap this size does not');
    out.push('matter. They are listed anyway, so the rule can be seen doing its work and');
    out.push('a tolerance set too wide is visible rather than silent.');
    out.push('');

    const sheetsWithTolerated: string[] = [];
    for (const d of swept.tolerated) {
      if (!sheetsWithTolerated.includes(d.sheet)) sheetsWithTolerated.push(d.sheet);
    }

    out.push('<details><summary>Show the cells</summary>', '');
    for (const sheet of sheetsWithTolerated) {
      const mine = swept.tolerated.filter((d) => d.sheet === sheet);
      out.push('', `**${cell(sheet)}** — ${n(mine.length)} cell(s)`, '');
      out.push(...table(
        ['Cell', 'Golden', 'Actual', 'Gap'],
        mine.map((d) => [
          code(d.address),
          cell(d.base || BLANK),
          cell(d.next || BLANK),
          gapBetween(d.base, d.next),
        ]),
      ));
    }
    out.push('', '</details>');
  }

  if (compared.length) {
    out.push('', `## What was verified (${compared.length})`, '');
    out.push('Every table layer 1 read, and the rectangle it read. This is the answer to');
    out.push('"what does this run actually cover" -- a table missing from here was never');
    out.push('compared by name and key, whatever else the report says about it.');
    out.push('');
    out.push('<details><summary>Show the tables</summary>', '');
    out.push(...table(
      ['table', 'golden', 'report', 'columns', 'rows', 'rows matched by'],
      compared.map((s) => [
        cell(s.label),
        code(s.range?.base ?? ''),
        s.range && s.range.next !== s.range.base ? code(s.range.next) : 'same',
        n(s.diff!.schema.compared.length),
        n(s.diff!.rows.compared),
        s.reason ? 'position' : 'key',
      ]),
      ['left', 'left', 'left', 'right', 'right', 'left'],
    ));
    out.push('', '</details>');
  }

  if (positional.length) {
    out.push('', `## Matched by row position (${positional.length})`, '');
    out.push('These tables have no column that identifies a row, so rows were paired by');
    out.push('their order. That is exact while both sides hold the same rows; an inserted');
    out.push('row shifts the rest, so one change would read as many.');
    out.push('');
    out.push('To pin one down, name the columns that identify a row. The key has to come');
    out.push('from that table\'s own columns, and has to name the table, since a sheet');
    out.push('holds more than one.');
    out.push('');

    // The warning above is the part that has to be read; the tables under it
    // are reference, and one of them can run to sixty column names. Folded,
    // the caveat stays in sight and the list is a click away.
    out.push('<details><summary>Show the tables</summary>', '');
    for (const s of positional) {
      const cols = s.diff!.schema.compared;
      out.push(`- **${cell(s.label)}** — ${n(s.diff!.rows.compared)} rows`);
      out.push(`  - columns: ${cols.map((c) => code(c)).join(', ')}`);
      const path = s.table !== s.sheet
        ? `{ "sheets": { ${JSON.stringify(s.sheet)}: { "tables": { ${JSON.stringify(s.table)}: { "keyColumns": [...] } } } } }`
        : `{ "sheets": { ${JSON.stringify(s.sheet)}: { "keyColumns": [...] } } }`;
      out.push(`  - ${code(path)}`);
    }
    out.push('', '</details>');
  }

  const review = diff.sheets.filter((s) => s.status === 'added' || s.status === 'skipped');
  if (review.length || diff.sheetSchema.moved.length) {
    out.push('', '## Sheets to review', '');
    for (const s of review) {
      out.push(s.status === 'added'
        ? `- added: **${cell(s.label)}** — not in the golden, so nothing to compare against`
        : `- not compared: **${cell(s.label)}** — ${cell(s.reason ?? '')}`);
    }
    for (const m of diff.sheetSchema.moved) {
      out.push(`- moved: **${cell(m.sheet)}** from position ${m.from} to ${m.to}`);
    }
  }

  if (swept) {
    // An excluded sheet has no coverage by definition, and is listed under
    // "Not verified" with its reason. Naming it here as well would read as two
    // problems where there is one decision.
    const onPurpose = new Set(
      diff.sheets.filter((s) => s.status === 'ignored').map((s) => s.sheet.trim().toLowerCase()),
    );
    const blind = swept.sheets.filter(
      (s) => s.status === 'swept' && s.compared === 0 && s.cells > 0
        && !onPurpose.has(s.sheet.trim().toLowerCase()),
    );
    if (blind.length) {
      out.push('', `## No coverage (${blind.length})`, '');
      out.push('Nothing on these sheets was compared by name and key. They match today;');
      out.push('nothing would report it if they stopped.');
      out.push('');
      out.push(...table(
        ['Sheet', 'Cells', 'Differing'],
        blind.map((s) => [cell(s.sheet), n(s.cells), n(s.differing)]),
      ));
    }
  }

  out.push(...notVerified(diff, swept));

  return `${out.join('\n')}\n`;
}
