/**
 * The current download against the one that was blessed.
 *
 * Every other check asks "is this right", and answers only where a source can
 * reach. This one asks "has this changed", and reaches everything -- including
 * the 136 ROLePlay columns nothing else can speak for, because the figures are
 * computed server-side and appear in no capture.
 *
 * The two questions are not interchangeable and the report keeps them apart. A
 * column that matches the golden is unchanged, not correct: if it was wrong
 * when the golden was blessed, it is still wrong and this will say nothing. It
 * catches drift, which is the failure a release actually introduces.
 *
 * Rows are paired by the same business key the sources use, never by position,
 * for the same reason: the template does not write its rows in a stable order.
 *
 * Structure is compared as well as values. A column that appeared or vanished,
 * a header that lost its colour, a data validation that changed shape -- all of
 * it drifts silently otherwise, and a descriptor only knows to check the rules
 * somebody thought to write down.
 */

const nameOf = (c) => (typeof c === 'string' ? c : c.name);

const read = (v) => {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.result !== undefined) return v.result;
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if (v.text !== undefined) return v.text;
    if (v instanceof Date) return v.toISOString();
  }
  return v;
};

/** Same rule the sources use: proportional, because Excel keeps 15 digits. */
function same(a, b) {
  const norm = (v) => (v == null || v === '' ? '' : (typeof v === 'number' ? v : String(v).trim()));
  const x = norm(a);
  const y = norm(b);
  if (x === '' && y === '') return true;
  const nx = Number(x);
  const ny = Number(y);
  if (x !== '' && y !== '' && !Number.isNaN(nx) && !Number.isNaN(ny)) {
    return Math.abs(nx - ny) <= 1e-12 * Math.max(1, Math.abs(nx), Math.abs(ny));
  }
  return String(x) === String(y);
}

const fillOf = (cell) => {
  const f = cell?.fill;
  if (!f?.fgColor) return null;
  return f.fgColor.argb ?? (f.fgColor.theme !== undefined ? `theme${f.fgColor.theme}` : null);
};

/** A validation reduced to what a comparison should care about. */
const ruleOf = (cell) => {
  const dv = cell?.dataValidation;
  if (!dv) return null;
  return `${dv.type}${dv.operator ? ` ${dv.operator}` : ''} ${JSON.stringify(dv.formulae ?? [])}`;
};

function shape(ws, headerRow, marker) {
  const columns = new Map();
  ws.getRow(headerRow).eachCell({ includeEmpty: false }, (c, n) => {
    const name = String(c.value ?? '').trim();
    if (name && !columns.has(name)) columns.set(name, n);
  });
  const value = (row, name) => (columns.has(name) ? read(ws.getRow(row).getCell(columns.get(name)).value) : undefined);
  const rows = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const v = value(r, marker);
    if (v !== undefined && v !== null && v !== '') rows.push(r);
  }
  return { columns, value, rows };
}

/**
 * @param current the sheet just downloaded
 * @param golden  the sheet it is judged against
 * @param spec    the resolved descriptor variant
 * @param key     (cell) => string, reading a row's identity from its own cells
 */
export function checkBaseline(current, golden, spec, key) {
  const findings = [];
  const { headerRow, rowMarker } = spec;

  const a = shape(golden, headerRow, rowMarker);
  const b = shape(current, headerRow, rowMarker);

  if (golden.name !== current.name) {
    findings.push({ kind: 'sheet', what: 'the sheet name', golden: golden.name, current: current.name });
  }

  // Columns first: a value comparison over a changed column set reports the
  // same drift once per row, which buries the one fact that explains it.
  for (const name of a.columns.keys()) {
    if (!b.columns.has(name)) findings.push({ kind: 'column', what: name, problem: 'gone since the golden' });
  }
  for (const name of b.columns.keys()) {
    if (!a.columns.has(name)) findings.push({ kind: 'column', what: name, problem: 'new since the golden' });
  }

  const shared = [...a.columns.keys()].filter((n) => b.columns.has(n));

  // Header presentation and the rules attached to the first data row. Neither
  // is a value, and both are contracts a descriptor might not know to check.
  for (const name of shared) {
    const g = golden.getRow(headerRow).getCell(a.columns.get(name));
    const c = current.getRow(headerRow).getCell(b.columns.get(name));
    if (fillOf(g) !== fillOf(c)) {
      findings.push({ kind: 'fill', what: name, golden: fillOf(g), current: fillOf(c) });
    }
    if (a.rows.length && b.rows.length) {
      const gr = ruleOf(golden.getRow(a.rows[0]).getCell(a.columns.get(name)));
      const cr = ruleOf(current.getRow(b.rows[0]).getCell(b.columns.get(name)));
      if (gr !== cr) findings.push({ kind: 'rule', what: name, golden: gr, current: cr });
    }
  }

  const keyed = (s, rows) => {
    const out = new Map();
    for (const r of rows) out.set(key((name) => s.value(r, name)), r);
    return out;
  };
  const ga = keyed(a, a.rows);
  const gb = keyed(b, b.rows);

  for (const k of ga.keys()) if (!gb.has(k)) findings.push({ kind: 'row', what: k, problem: 'gone since the golden' });
  for (const k of gb.keys()) if (!ga.has(k)) findings.push({ kind: 'row', what: k, problem: 'new since the golden' });

  let compared = 0;
  for (const [k, br] of gb) {
    const ar = ga.get(k);
    if (ar === undefined) continue;
    for (const name of shared) {
      compared++;
      const g = a.value(ar, name);
      const c = b.value(br, name);
      if (!same(g, c)) findings.push({ kind: 'value', what: name, row: k, golden: g, current: c });
    }
  }

  return {
    ok: findings.length === 0,
    columns: shared.length,
    rows: gb.size,
    compared,
    findings,
  };
}

/** Which columns the baseline actually looked at, for the coverage figure. */
export const baselineColumns = (outcome) => outcome?.columns ?? 0;

export { nameOf };
