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
  const { coverage, derived, markers, fills, blocks, painted, validations, baseline } = outcome;

  md.push(`# ${kind} · ${name}`, '');
  md.push(`Template: \`${outcome.file}\`, sheet \`${outcome.sheet}\``, '');

  const total = outcome.results.reduce((t, r) => t + (r.findings?.length ?? 0), 0)
    + fills.findings.length + (markers?.findings.length ?? 0) + (painted?.findings.length ?? 0)
    + (derived?.findings.length ?? 0) + (blocks?.findings.length ?? 0)
    + (validations?.findings.length ?? 0)
    + (coverage?.unchecked.length ?? 0) + (coverage?.shadowed.length ?? 0);

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
    // A source carrying no columns is not an empty comparison -- it is a
    // membership check, and saying "4x0 clean" makes real work look like none.
    const membership = r.columns === 0;
    summary.push(
      r.findings.length
        ? `${r.name} ${n(r.findings.length)} problem(s)`
        : membership
          ? `${r.name} ${n(r.rows)} row(s), the right ones`
          : `${r.name} ${n(r.rows)}×${r.columns} clean`,
    );

    md.push(`## ${r.name}`, '');
    md.push(
      membership
        ? `${n(r.rows)} row(s) asked for, checked for presence only -- this source carries `
          + 'no values to compare.'
        : `${n(r.rows)} row(s) × ${r.columns} column(s) compared.`,
      '',
    );
    if (!r.findings.length) {
      md.push(
        membership
          ? 'The template holds exactly the rows that were asked for, and no others.'
          : 'Every value the template holds is the value this source gave.',
        '',
      );
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

  if (baseline) {
    md.push('## Against the golden', '');
    if (baseline.missing) {
      lines.push('    no golden to compare against');
      md.push(
        'There is no `golden/` for this case yet, so nothing was compared. Run '
        + '`npm run bless:templates` once this download has been read and is worth '
        + 'keeping as the contract.',
        '',
      );
    } else {
      md.push(
        `${n(baseline.compared)} value(s) over ${n(baseline.rows)} row(s) × `
        + `${n(baseline.columns)} column(s), plus the header fills and the rules attached to them.`,
        '',
      );
      if (baseline.parts) {
        md.push(
          `The two files are also compared as archives: ${n(baseline.parts.compared)} of `
          + `${n(baseline.parts.total)} XML part(s), byte for byte. That needs no spreadsheet `
          + 'model, so drawings, merged cells, column widths, defined names, protection and '
          + 'print setup cannot drift unnoticed even though no check here understands them.',
          '',
        );
        if (baseline.parts.producer) {
          md.push(`Written by \`${baseline.parts.producer}\`.`, '');
        }
      }
      // Said plainly, because the two claims are easy to conflate and only one
      // of them is about correctness.
      md.push(
        '*Unchanged is not the same as correct.* This reaches every column, '
        + 'including the ones no capture can speak for — but a figure that was '
        + 'wrong when the golden was blessed is still wrong and this will not say so.',
        '',
      );
      if (baseline.ok) {
        lines.push(`    ${n(baseline.columns)} column(s) unchanged since the golden`);
        md.push('Identical to the golden in every cell compared.', '');
      } else {
        lines.push(`    golden ${n(baseline.findings.length)} difference(s)`);
        md.push('| What | Where | Golden | Current |', '| --- | --- | --- | --- |');
        for (const f of baseline.findings) {
          md.push(
            `| ${f.kind} | ${f.what}${f.row ? ` \`${f.row}\`` : ''} `
            + `| ${f.problem ?? show(f.golden)} | ${f.problem ? '—' : show(f.current)} |`,
          );
        }
        md.push('');
      }
    }
  }

  // Coverage before the findings, because it is the figure that says how much
  // the findings are worth. A page of "clean" over a third of the sheet is not
  // a clean template; it is a third of one.
  if (coverage) {
    md.push('## Coverage', '');
    md.push(
      `${n(coverage.checked)} of ${n(coverage.distinct)} distinct column name(s) checked against a `
      + `source or computed from the sheet`
      + (coverage.total !== coverage.distinct
        ? `, from ${n(coverage.total)} header cell(s).`
        : '.'),
      '',
    );
    lines.push(
      `    ${n(coverage.checked)}/${n(coverage.distinct)} columns covered`
      + (coverage.declared.length ? `, ${n(coverage.declared.length)} declared unverifiable` : '')
      + (coverage.unchecked.length ? `, ${n(coverage.unchecked.length)} UNCHECKED` : ''),
    );

    if (coverage.duplicates.length) {
      const bad = coverage.shadowed.length;
      if (bad) lines.push(`    ${n(bad)} SHADOWED duplicate header(s)`);
      md.push(
        'Header names that appear more than once. Lookup is by name, so the first '
        + 'wins and the rest are unreachable -- nothing can check a shadowed column.',
        '',
      );
      md.push('| Column | Cells | |', '| --- | --- | --- |');
      for (const d of coverage.duplicates) {
        const known = !coverage.shadowed.some((x) => x.column === d.column);
        md.push(`| ${d.column} | ${d.cells.join(', ')} | ${known ? 'declared' : '**undeclared**'} |`);
      }
      md.push('');
    }

    if (coverage.declared.length) {
      const byReason = new Map();
      for (const d of coverage.declared) {
        if (!byReason.has(d.reason)) byReason.set(d.reason, []);
        byReason.get(d.reason).push(d.column);
      }
      md.push('Left unchecked on purpose, with the reason each was excused:', '');
      md.push('| Reason | Columns |', '| --- | --- |');
      for (const [reason, cols] of byReason) {
        md.push(`| ${reason} | ${n(cols.length)}: ${cols.join(', ')} |`);
      }
      md.push('');
    }

    if (coverage.unchecked.length) {
      md.push(
        `**${n(coverage.unchecked.length)} column(s) nobody checked and nobody excused.** `
        + 'Either a source can cover them, or the descriptor should say why not.',
        '',
      );
      for (const u of coverage.unchecked) md.push(`- ${u.column}`);
      md.push('');
    }
  }

  if (blocks && (blocks.present.length || blocks.findings.length)) {
    md.push('## Optional blocks', '');
    for (const p of blocks.present) {
      md.push(`- **${p.name}**: included, ${n(p.columns)} column(s).`);
    }
    if (blocks.ok) {
      md.push('', 'Each block present is present in full.', '');
    } else {
      lines.push(`    blocks ${n(blocks.findings.length)} problem(s)`);
      md.push('');
      for (const f of blocks.findings) {
        md.push(`**${f.block}**: ${f.problem}`, '');
        md.push('Missing:', '');
        for (const c of f.missing) md.push(`- ${c}`);
        md.push('');
      }
    }
  }

  if (derived?.checked || derived?.findings.length) {
    md.push('## Columns the sheet computes', '');
    if (derived.ok) {
      md.push(`${n(derived.checked)} value(s) checked; each is what the row's own columns imply.`, '');
      // Anchored and merely-consistent are both real checks and are not the
      // same claim. A rule tying four unverified figures together says they
      // agree, not that any of them is right.
      if (derived.consistency) {
        md.push(
          `${n(derived.anchored)} column(s) computed from values verified against a source, so `
          + `the result is verified too; ${n(derived.consistency)} tie unverified figures to each `
          + 'other, which catches one moving without the others but does not pin any of them down.',
          '',
        );
      }
    } else {
      lines.push(`    derived ${n(derived.findings.length)} problem(s)`);
      md.push('| Row | Column | Template | Should be |', '| --- | --- | --- | --- |');
      for (const f of derived.findings) {
        md.push(
          f.problem
            ? `| — | ${f.column} | — | ${f.problem} |`
            : `| ${f.row} | ${f.column} | ${show(f.template)} | ${show(f.expected)} |`,
        );
      }
      md.push('');
    }
  }

  if (markers?.checked || markers?.findings.length) {
    md.push('## Editable markers', '');
    if (markers.ok) {
      md.push(
        `The ${n(markers.checked)} column(s) marked editable in the header are the ones that `
        + 'should be, and no others carry the marker.',
        '',
      );
    } else {
      lines.push(`    markers ${n(markers.findings.length)} problem(s)`);
      md.push('| Column | Problem |', '| --- | --- |');
      for (const f of markers.findings) md.push(`| ${f.column} | ${f.problem} |`);
      md.push('');
    }
  }

  if (validations && (validations.checked || validations.custom || validations.findings.length)) {
    md.push('## The rules the sheet carries about itself', '');
    md.push(
      `${n(validations.checked)} value(s) checked against the data validation beside them`
      + (validations.custom ? `; ${n(validations.custom)} custom rule(s) left unevaluated.` : '.'),
      '',
    );
    if (validations.ok) {
      md.push('Every value the template wrote satisfies the rule it wrote beside it.', '');
    } else {
      lines.push(`    validations ${n(validations.findings.length)} problem(s)`);
      md.push('| Cell | Column | Value | Problem |', '| --- | --- | --- | --- |');
      for (const f of validations.findings) {
        md.push(`| ${f.address} | ${f.column} | ${f.value === undefined ? '—' : show(f.value)} | ${f.problem} |`);
      }
      md.push('');
    }
  }

  if (painted?.checked || painted?.findings.length) {
    md.push('## Colour applied by rule', '');
    if (painted.ok) {
      md.push(
        `${n(painted.checked)} column(s) carry the rule that paints them, so the marker in the `
        + 'header is what a reader actually sees.',
        '',
      );
    } else {
      lines.push(`    conditional fills ${n(painted.findings.length)} problem(s)`);
      md.push('| Group | Column | Problem |', '| --- | --- | --- |');
      for (const f of painted.findings) md.push(`| ${f.group} | ${f.column} | ${f.problem} |`);
      md.push('');
    }
  }

  md.push('## Editable columns', '');
  if (fills.ok) {
    md.push(
      'The columns marked editable are the ones that should be, and no others '
      + 'carry the fill.',
      '',
    );
  } else {
    lines.push(`    fills ${n(fills.findings.length)} problem(s)`);
    md.push('| Group | Column | Problem | Expected | Actual |', '| --- | --- | --- | --- | --- |');
    for (const f of fills.findings) {
      md.push(
        `| ${f.group} | ${f.column} | ${f.problem} | ${f.expected ?? '—'} | ${f.actual ?? '—'} |`,
      );
    }
    md.push('');
  }

  if (total) lines.push(`    ${n(total)} finding(s) in total`);

  return { lines, markdown: md.join('\n') };
}
