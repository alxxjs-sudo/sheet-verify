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
    const { argb, columns, row = 'data', optional = false } = spec;

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
        // Some columns come and go with the data. The ROLePlay divider is only
        // written when there is ROLePlay data to divide, so its absence is not
        // a defect -- but its absence WITH the block present would be, which is
        // why the block is checked separately rather than this being waved off.
        if (!optional) {
          findings.push({ group, column: name, problem: 'column not found in the template' });
        }
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

/**
 * Checks the other way a template can say "you may edit this".
 *
 * The program selection sheet paints the editable columns yellow. The overrides
 * sheet paints nothing and marks them in the header text instead -- "Edison
 * Treaty Premium *" is editable, "MetaRisk Treaty Premium" is not. Different
 * notation, identical contract, and worth checking in both directions for the
 * same reason: a column that quietly gained the marker hands users a field
 * nobody meant them to have.
 *
 * @param ws        the worksheet
 * @param markers   { suffix, columns } -- columns named WITHOUT the suffix
 * @param headerRow 1-based row holding the column names
 */
export function checkMarkers(ws, markers, headerRow) {
  const { suffix, columns } = markers;
  const findings = [];

  const names = [];
  ws.getRow(headerRow).eachCell({ includeEmpty: false }, (c) => {
    const name = String(c.value ?? '').trim();
    if (name) names.push(name);
  });

  const marked = new Set(names.filter((n) => n.endsWith(suffix)).map((n) => n.slice(0, -suffix.length).trim()));
  const want = new Set(columns);

  for (const name of want) {
    if (!marked.has(name)) {
      findings.push({
        column: name,
        problem: names.includes(name)
          ? `lost its "${suffix.trim()}" marker, so it no longer reads as editable`
          : 'column not found in the template',
      });
    }
  }
  for (const name of marked) {
    if (!want.has(name)) {
      findings.push({ column: name, problem: `carries the "${suffix.trim()}" marker but is not listed as editable` });
    }
  }

  return { ok: findings.length === 0, checked: want.size, findings };
}

/** "B4" -> { col: 2, row: 4 } */
function addr(a) {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
}

/** Whether a conditional-format ref ("B4", "B4:D9", or several space-separated) covers a cell. */
function covers(ref, col, row) {
  for (const part of String(ref ?? '').split(/\s+/)) {
    const [a, b] = part.split(':');
    const from = addr(a);
    if (!from) continue;
    const to = b ? addr(b) : from;
    if (!to) continue;
    if (col >= Math.min(from.col, to.col) && col <= Math.max(from.col, to.col)
      && row >= Math.min(from.row, to.row) && row <= Math.max(from.row, to.row)) return true;
  }
  return false;
}

/**
 * Checks the colour a template paints by rule rather than by hand.
 *
 * The overrides sheet has no yellow cells at all. It has a conditional format
 * on every header cell -- `endsWith "*"` -> fill FFFFFF00 -- so the yellow is
 * COMPUTED from the header text. That makes the marker and the colour one
 * contract rather than two, which is why the " *" check covers what a reader
 * sees.
 *
 * It leaves one gap, and this closes it: drop the rules and every " *" is still
 * in place while nothing is yellow any more. The marker check would pass and
 * the sheet would silently stop telling anyone which fields they may edit.
 *
 * @param ws        the worksheet
 * @param spec      { argb, type, text, row } -- the rule every column must carry
 * @param headerRow 1-based row holding the column names
 */
export function checkConditionalFills(ws, groups, headerRow) {
  const findings = [];
  let checked = 0;

  const formats = ws.conditionalFormattings ?? [];

  for (const [group, spec] of Object.entries(groups)) {
    const { argb, type, text, row = 'header' } = spec;
    const at = row === 'header' ? headerRow : headerRow + 1;

    const columns = [];
    ws.getRow(headerRow).eachCell({ includeEmpty: false }, (c, n) => {
      if (String(c.value ?? '').trim()) columns.push({ name: String(c.value).trim(), n });
    });

    for (const { name, n } of columns) {
      checked++;
      const rules = formats
        .filter((f) => covers(f.ref, n, at))
        .flatMap((f) => f.rules ?? []);

      const match = rules.find((r) => {
        if (type && r.type !== type) return false;
        // The rule's `text` attribute is dropped on read, so the text is looked
        // for in the formula it keeps: endsWith "*" arrives as
        // RIGHT(B4,LEN("*"))="*", which states the same thing twice.
        if (text !== undefined) {
          const quoted = `"${text}"`;
          const said = r.text === text || (r.formulae ?? []).some((f) => String(f).includes(quoted));
          if (!said) return false;
        }
        const fill = r.style?.fill;
        const got = fill?.bgColor?.argb ?? fill?.fgColor?.argb ?? null;
        return got === argb;
      });

      if (!match) {
        findings.push({
          group,
          column: name,
          problem: rules.length
            ? `carries ${rules.length} rule(s), none painting ${argb} on ${type ?? 'any'} "${text ?? ''}"`
            : 'carries no conditional format, so it can never be painted',
        });
      }
    }
  }

  return { ok: findings.length === 0, checked, findings };
}

/** One line per finding, for a run log. */
export function formatFindings(findings) {
  return findings.map(
    (f) => `  ${f.group}: ${f.column} -- ${f.problem}`
      + (f.expected !== undefined ? ` (expected ${f.expected ?? 'none'}, got ${f.actual ?? 'none'})` : ''),
  );
}
