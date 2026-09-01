/**
 * What a case has to say, for a terminal and for a file.
 *
 * The terminal lines are what somebody reads while the run is going; the
 * markdown is what they open afterwards, and it lists every finding rather than
 * the first handful. A log that truncates is a log that gets distrusted.
 */

const n = (v) => v.toLocaleString('en-US');

const show = (v) => {
  if (v === undefined) return '(no such column)';
  if (v === null || v === '') return '(empty)';
  return typeof v === 'number' ? String(v) : `"${v}"`;
};

export function report(kind, name, outcome) {
  const lines = [];
  const md = [];

  md.push(`# ${kind} · ${name}`, '');
  md.push(`Template: \`${outcome.file}\``, '');

  const total = outcome.results.reduce((t, r) => t + (r.findings?.length ?? 0), 0)
    + outcome.fills.findings.length;

  lines.push(`${outcome.ok ? '✓' : '✗'} ${kind} · ${name}`);

  // A source that agrees is worth a line of its own. "Nothing to report" and
  // "nobody checked" read identically otherwise, and only one of them is fine.
  const summary = [];
  for (const r of outcome.results) {
    if (r.skipped) {
      summary.push(`${r.name} skipped (${r.skipped})`);
      md.push(`## ${r.name}`, '', `Skipped: ${r.skipped}`, '');
      continue;
    }
    summary.push(
      r.findings.length
        ? `${r.name} ${n(r.findings.length)} problem(s)`
        : `${r.name} ${n(r.rows)}×${r.columns} clean`,
    );

    md.push(`## ${r.name}`, '');
    md.push(`${n(r.rows)} row(s) × ${r.columns} column(s) compared.`, '');
    if (!r.findings.length) {
      md.push('Every value the template holds is the value this source gave.', '');
    } else {
      md.push('| Row | Column | Template | Source |', '| --- | --- | --- | --- |');
      for (const f of r.findings) {
        md.push(
          f.problem
            ? `| \`${f.key}\` | — | — | ${f.problem} |`
            : `| \`${f.key}\` | ${f.column} | ${show(f.template)} | ${show(f.source)} |`,
        );
      }
      md.push('');
    }
  }
  lines.push(`    ${summary.join(', ')}`);

  md.push('## Editable columns', '');
  if (outcome.fills.ok) {
    md.push(
      'The columns marked editable are the ones that should be, and no others '
      + 'carry the fill.',
      '',
    );
  } else {
    lines.push(`    fills ${n(outcome.fills.findings.length)} problem(s)`);
    md.push('| Group | Column | Problem | Expected | Actual |', '| --- | --- | --- | --- | --- |');
    for (const f of outcome.fills.findings) {
      md.push(
        `| ${f.group} | ${f.column} | ${f.problem} | ${f.expected ?? '—'} | ${f.actual ?? '—'} |`,
      );
    }
    md.push('');
  }

  if (total) lines.push(`    ${n(total)} finding(s) in total`);

  return { lines, markdown: md.join('\n') };
}
