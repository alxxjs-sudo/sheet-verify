import type { CellValue, DiffResult, WorkbookDiffResult } from './types.js';

export interface ReportOptions {
  /** Max rows listed per section before truncating. Default 20. */
  limit?: number;
  /** Separator used in composite keys, replaced with " / " for display. */
  keySeparator?: string;
  /** Show cascaded (non-root-cause) value differences. Default false. */
  showCascades?: boolean;
  /**
   * Both files were recalculated by Excel before comparison. Reported, never
   * acted on -- it changes what the comparison could see, so it belongs at the
   * top of the report rather than in the command someone ran last week.
   */
  recalculated?: boolean;
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
const num = (n: number): string => String(n);

const show = (v: CellValue): string => {
  if (v === null || v === undefined) return '∅';
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? '∅' : v.toISOString().slice(0, 10);
  if (typeof v === 'string') return JSON.stringify(v);
  if (typeof v === 'number') return num(v);
  return String(v);
};

/**
 * Human-readable diff. Ordered so the most actionable thing is first:
 * structural errors, then defects, then review items.
 */
export function formatReport(d: DiffResult, opts: ReportOptions = {}): string {
  const limit = opts.limit ?? 20;
  const sep = opts.keySeparator ?? '␟';
  const key = (k: string) => k.split(sep).join(' / ');
  const out: string[] = [];
  const h = (t: string) => out.push('', t, '─'.repeat(Math.min(t.length, 72)));

  const more = (n: number) => (n > limit ? [`  … and ${n - limit} more`] : []);

  out.push(
    `baseline  ${d.base.source}${d.base.sheet ? ` [${d.base.sheet}]` : ''}  ${d.base.rows} rows × ${d.base.columns} cols`,
    `actual    ${d.next.source}${d.next.sheet ? ` [${d.next.sheet}]` : ''}  ${d.next.rows} rows × ${d.next.columns} cols`,
    `compared  ${d.rows.compared} rows × ${d.schema.compared.length} columns`,
  );

  if (d.errors.length) {
    h('COMPARISON INTEGRITY — these make the result untrustworthy');
    for (const e of d.errors) out.push(`  ✗ ${e}`);
  }

  if (d.formulas.length) {
    h(`FORMULA CHANGES (${d.formulas.length}) — calculation logic differs`);
    for (const f of d.formulas.slice(0, limit)) {
      out.push(`  ${key(f.key)} · ${f.column} @${f.address}`);
      out.push(`      baseline  ${f.baseA1 ?? '∅'}`);
      out.push(`      actual    ${f.nextA1 ?? '∅'}`);
      if (f.base !== f.baseA1 || f.next !== f.nextA1) {
        out.push(`      normalised  ${f.base ?? '∅'}  →  ${f.next ?? '∅'}`);
      }
    }
    out.push(...more(d.formulas.length));
  }

  const roots = d.values.filter((v) => v.rootCause);
  const cascades = d.values.filter((v) => !v.rootCause);
  if (roots.length) {
    h(`VALUE CHANGES (${roots.length} root cause${roots.length === 1 ? '' : 's'}${cascades.length ? `, ${cascades.length} cascaded` : ''})`);
    for (const v of roots.slice(0, limit)) {
      const delta = v.delta !== undefined ? `  (Δ ${num(v.delta)})` : '';
      const flag = v.formulaChanged ? '  [formula also changed]' : '';
      out.push(`  ${key(v.key)} · ${v.column} @${v.address}: ${show(v.base)} → ${show(v.next)}${delta}${flag}`);
    }
    out.push(...more(roots.length));
  }
  if (cascades.length && opts.showCascades) {
    h(`CASCADED VALUE CHANGES (${cascades.length}) — downstream of the above`);
    for (const v of cascades.slice(0, limit)) {
      out.push(`  ${key(v.key)} · ${v.column} @${v.address}: ${show(v.base)} → ${show(v.next)}`);
    }
    out.push(...more(cascades.length));
  }

  if (d.types.length) {
    h(`TYPE CHANGES (${d.types.length}) — same rendering, different type`);
    for (const t of d.types.slice(0, limit)) {
      out.push(`  ${key(t.key)} · ${t.column} @${t.address}: ${t.baseKind} → ${t.nextKind}  (${show(t.value)})`);
    }
    out.push(...more(d.types.length));
  }

  if (d.invariants.length) {
    h(`INVARIANT FAILURES (${d.invariants.length}) — wrong regardless of baseline`);
    for (const i of d.invariants.slice(0, limit)) {
      const where = [i.key && key(i.key), i.column, i.address].filter(Boolean).join(' · ');
      out.push(`  ${i.invariant}${where ? ` · ${where}` : ''}: ${i.detail}`);
    }
    out.push(...more(d.invariants.length));
  }

  const { added, removed, moved } = d.schema;
  if (added.length || removed.length || moved.length) {
    h('SCHEMA CHANGES — review, then re-bless the baseline if intended');
    for (const c of added) out.push(`  + column "${c}"`);
    for (const c of removed) out.push(`  − column "${c}"`);
    for (const m of moved) out.push(`  ~ column "${m.column}" moved ${m.from} → ${m.to}`);
  }

  const repeated = new Set([...d.rows.duplicateKeysBase, ...d.rows.duplicateKeysNext]);
  if (repeated.size) {
    h(`REPEATED ROW KEYS (${repeated.size}) — matched in order of appearance`);
    out.push('  These rows do not have a key of their own — a per-group "Total" line, say.');
    out.push('  The first in the baseline is compared with the first here, and so on, which');
    out.push('  holds while the groups stay in the same order.');
    for (const k of [...repeated].slice(0, limit)) out.push(`  ~ ${key(k) || '(blank)'}`);
    out.push(...more(repeated.size));
  }

  if (d.rows.added.length || d.rows.removed.length) {
    h('ROW POPULATION — review');
    if (d.rows.added.length) {
      out.push(`  + ${d.rows.added.length} row(s): ${d.rows.added.slice(0, limit).map(key).join(', ')}${d.rows.added.length > limit ? ' …' : ''}`);
    }
    if (d.rows.removed.length) {
      out.push(`  − ${d.rows.removed.length} row(s): ${d.rows.removed.slice(0, limit).map(key).join(', ')}${d.rows.removed.length > limit ? ' …' : ''}`);
    }
  }

  if (d.ok) out.push('', d.reviewOnly ? '✓ No defects. Schema changed — review above.' : '✓ Identical.');
  return out.join('\n');
}

/**
 * Workbook diff, ordered the same way as a sheet diff: the sheets that failed
 * come first with their full detail, and everything merely worth reviewing is
 * collapsed to one line each at the end.
 */
export function formatWorkbookReport(w: WorkbookDiffResult, opts: ReportOptions = {}): string {
  const out: string[] = [];
  const rule = (t: string) => out.push('', t, '═'.repeat(Math.min(t.length, 72)));

  const compared = w.sheets.filter((s) => s.status === 'compared');
  const failed = compared.filter((s) => !s.diff!.ok);
  const positional = compared.filter((s) => s.reason);
  const counts = [
    // Tables, not sheets: a sheet holding an info block and a data table
    // contributes two.
    `${compared.length} table${compared.length === 1 ? '' : 's'} compared`,
    positional.length ? `${positional.length} by position` : '',
    w.sheetSchema.added.length ? `${w.sheetSchema.added.length} added` : '',
    w.sheetSchema.removed.length ? `${w.sheetSchema.removed.length} removed` : '',
    w.sheets.filter((s) => s.status === 'skipped').length
      ? `${w.sheets.filter((s) => s.status === 'skipped').length} not compared`
      : '',
    w.sheets.filter((s) => s.status === 'ignored').length
      ? `${w.sheets.filter((s) => s.status === 'ignored').length} ignored`
      : '',
  ].filter(Boolean);

  out.push(
    `baseline  ${w.base.source}  ${w.base.sheets.length} sheet(s)`,
    `actual    ${w.next.source}  ${w.next.sheets.length} sheet(s)`,
    `          ${counts.join(' · ')}`,
  );

  if (w.errors.length) {
    rule('WORKBOOK INTEGRITY — these make the result untrustworthy');
    for (const e of w.errors) out.push(`  ✗ ${e}`);
  }

  if (w.sheetSchema.removed.length) {
    rule(`SHEETS REMOVED (${w.sheetSchema.removed.length}) — output that is no longer produced`);
    for (const s of w.sheetSchema.removed) out.push(`  ✗ − sheet "${s}"`);
  }

  for (const s of failed) {
    rule(`SHEET "${s.label}" — ${summarize(s.diff!)}`);
    out.push(
      formatReport(s.diff!, opts)
        .split('\n')
        .map((l) => (l ? `  ${l}` : l))
        .join('\n'),
    );
  }

  if (positional.length) {
    rule(`MATCHED BY POSITION (${positional.length}) — no row key on these tables`);
    out.push('  Rows were paired by their order in the table, which is exact while both');
    out.push('  sides hold the same rows. An inserted row shifts the rest, so one change');
    out.push('  there reads as many.');
    out.push('');
    out.push('  To pin one down, pick the columns that identify a row and name them in');
    out.push('  case.json. Each table\'s own columns are listed below -- the key has to');
    out.push('  come from those, and it has to name the table, since a sheet holds more');
    out.push('  than one.');

    for (const s of positional) {
      const n = s.diff!.rows.compared;
      const cols = s.diff!.schema.compared;
      out.push('');
      out.push(`  ~ ${s.label} — ${n} row${n === 1 ? '' : 's'}`);
      // The real column names, so what gets copied out of here actually
      // resolves. An invented example is worse than none: it looks right,
      // fails with "key column not found", and costs an hour.
      const listed = cols.slice(0, 8).map((c) => JSON.stringify(c)).join(', ');
      out.push(`      columns: ${listed}${cols.length > 8 ? `, … ${cols.length - 8} more` : ''}`);
      if (s.table !== s.sheet) {
        out.push(`      { "sheets": { ${JSON.stringify(s.sheet)}: { "tables": { ${JSON.stringify(s.table)}: { "keyColumns": [ … ] } } } } }`);
      } else {
        out.push(`      { "sheets": { ${JSON.stringify(s.sheet)}: { "keyColumns": [ … ] } } }`);
      }
    }
  }

  const review = w.sheets.filter(
    (s) => s.status === 'added' || s.status === 'skipped' || s.diff?.reviewOnly,
  );
  if (review.length || w.sheetSchema.moved.length) {
    rule('SHEETS TO REVIEW');
    for (const s of review) {
      if (s.status === 'added') out.push(`  + sheet "${s.label}" — new, not compared`);
      else if (s.status === 'skipped') out.push(`  ? table "${s.label}" — ${s.reason}`);
      else out.push(`  ~ table "${s.label}" — ${summarize(s.diff!)}`);
    }
    for (const m of w.sheetSchema.moved) {
      out.push(`  ~ sheet "${m.sheet}" moved ${m.from} → ${m.to}`);
    }
  }

  if (w.ok) {
    out.push('', w.reviewOnly ? '✓ No defects. Sheets changed — review above.' : '✓ Identical.');
  }
  return out.join('\n');
}

/** One-line summary of a workbook diff, for test titles and CI logs. */
export function summarizeWorkbook(w: WorkbookDiffResult): string {
  // Two units, and mixing them is how "33 sheets failing" appeared on a report
  // holding 22 sheets. `sheetSchema` counts sheets -- a sheet added, removed or
  // moved is a sheet. `w.sheets` holds one entry per *table*, so a workbook of
  // 22 sheets with several tables on some of them has 49 of them, and counting
  // those as sheets overstates the damage and cannot be reconciled with the
  // file by anyone reading it.
  const sheets = (n: number) => `${n} sheet${n === 1 ? '' : 's'}`;
  const tables = (n: number) => `${n} table${n === 1 ? '' : 's'}`;
  if (w.ok && !w.reviewOnly) return 'identical';
  const bits: string[] = [];
  if (w.errors.length) bits.push(`${w.errors.length} integrity`);
  if (w.sheetSchema.removed.length) bits.push(`${sheets(w.sheetSchema.removed.length)} removed`);
  const failed = w.sheets.filter((s) => s.status === 'compared' && !s.diff!.ok);
  if (failed.length) bits.push(`${tables(failed.length)} failing`);
  if (w.sheetSchema.added.length) bits.push(`${sheets(w.sheetSchema.added.length)} added`);
  const skipped = w.sheets.filter((s) => s.status === 'skipped').length;
  if (skipped) bits.push(`${tables(skipped)} not compared`);
  if (w.sheetSchema.moved.length) bits.push(`${sheets(w.sheetSchema.moved.length)} moved`);

  // Tables that changed without a defect -- an inserted column, say. Without
  // this a schema-only release reads as "no differences", which is wrong: it
  // passed, but something did change and the summary is what most people read.
  const review = w.sheets.filter((s) => s.diff?.reviewOnly).length;
  if (review) bits.push(`${tables(review)} to review`);

  return bits.join(', ') || 'no differences';
}

/** One-line summary for test titles and CI logs. */
export function summarize(d: DiffResult): string {
  if (d.ok && !d.reviewOnly) return 'identical';
  const bits: string[] = [];
  if (d.errors.length) bits.push(`${d.errors.length} integrity`);
  if (d.formulas.length) bits.push(`${d.formulas.length} formula`);
  const roots = d.values.filter((v) => v.rootCause).length;
  if (roots) bits.push(`${roots} value`);
  if (d.types.length) bits.push(`${d.types.length} type`);
  if (d.invariants.length) bits.push(`${d.invariants.length} invariant`);
  const schema = d.schema.added.length + d.schema.removed.length + d.schema.moved.length;
  if (schema) bits.push(`${schema} schema`);
  const rows = d.rows.added.length + d.rows.removed.length;
  if (rows) bits.push(`${rows} row`);
  return bits.join(', ') || 'no differences';
}
