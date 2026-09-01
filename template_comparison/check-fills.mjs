/**
 * Checks that a template still marks the right columns the right colour.
 *
 * This is not about looks. In these templates the fill IS the contract: yellow
 * says "the user may edit this", and the sheet says so in as many words at A1 --
 * "*Only fields highlighted yellowish are editable in the UI". A release that
 * drops the fill on a column takes the column away from users; one that adds it
 * hands them a field nobody meant them to have.
 *
 * Values live in the cells and are compared by sheet-verify. Colour does not
 * live in the cells -- sheet-verify's model carries address, value and formula
 * and nothing else -- so this check is separate by necessity as well as by
 * design, and stays out of that tool entirely.
 *
 * Both directions are checked. Missing is the obvious failure; EXTRA is the
 * dangerous one, because a column that quietly became editable looks exactly
 * like a column that was always meant to be.
 */

const argbOf = (cell) => {
  const f = cell?.fill;
  if (!f || !f.fgColor) return null;
  return f.fgColor.argb ?? (f.fgColor.theme !== undefined ? `theme${f.fgColor.theme}` : null);
};

/**
 * @param ws        the worksheet
 * @param fills     descriptor's `fills` block, keyed by meaning
 * @param headerRow 1-based row holding the column names
 * @returns { ok, findings[] }
 */
export function checkFills(ws, fills, headerRow) {
  const findings = [];

  // Column name -> column number, from the header row.
  const byName = new Map();
  ws.getRow(headerRow).eachCell({ includeEmpty: false }, (c, n) => {
    const name = String(c.value ?? '').trim();
    if (name) byName.set(name, n);
  });

  for (const [group, spec] of Object.entries(fills)) {
    const { argb, columns, row = 'data' } = spec;

    // Which row carries the colour. The editable yellow is on the data rows;
    // the divider's red and the header band are on the header row itself. Which
    // one it is has to be stated, because reading the wrong row reports a fill
    // as missing when it is simply somewhere else -- as this check did to its
    // own descriptor the first time it ran.
    const at = row === 'header' ? headerRow : headerRow + 1;

    // No `columns` means the whole band: every column must carry it.
    if (!columns) {
      for (const [name, n] of byName) {
        const got = argbOf(ws.getRow(at).getCell(n));
        if (got !== argb) {
          findings.push({ group, column: name, expected: argb, actual: got, problem: 'fill differs' });
        }
      }
      continue;
    }

    const want = new Set(columns);

    // 1. Every column that should carry the fill does.
    for (const name of columns) {
      const n = byName.get(name);
      if (n === undefined) {
        findings.push({ group, column: name, problem: 'column not found in the template' });
        continue;
      }
      const got = argbOf(ws.getRow(at).getCell(n));
      if (got !== argb) {
        findings.push({
          group, column: name, expected: argb, actual: got,
          problem: got === null ? 'fill is gone' : 'fill changed',
        });
      }
    }

    // 2. No column that should not carry it does. This is the direction that
    //    catches a field becoming editable when nobody asked for it.
    for (const [name, n] of byName) {
      if (want.has(name)) continue;
      if (argbOf(ws.getRow(at).getCell(n)) === argb) {
        findings.push({
          group, column: name, expected: null, actual: argb,
          problem: `carries the ${group} fill but is not listed`,
        });
      }
    }
  }

  return { ok: findings.length === 0, findings };
}

/** One line per finding, for a run log. */
export function formatFindings(findings) {
  return findings.map(
    (f) => `  ${f.group}: ${f.column} -- ${f.problem}`
      + (f.expected !== undefined ? ` (expected ${f.expected ?? 'none'}, got ${f.actual ?? 'none'})` : ''),
  );
}
