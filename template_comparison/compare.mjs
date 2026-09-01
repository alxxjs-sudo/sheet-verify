/**
 * Compares one downloaded template against the sources that produced it.
 *
 * Rows are paired by business key, never by position. The template does not
 * write its rows in the order the request lists them -- in one capture 36 of 68
 * lined up and the rest had shuffled -- so a positional comparison would report
 * every row after the first shuffle as wrong.
 */
import { readdir } from 'node:fs/promises';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { checkFills, checkMarkers, checkConditionalFills } from './check-fills.mjs';
import { collapse } from './values.mjs';

/** A column is a name, or a name with its own idea of what agreement means. */
const nameOf = (c) => (typeof c === 'string' ? c : c.name);

/** Opens the one .xlsx in a folder. */
export async function openTemplate(dir) {
  const names = (await readdir(dir)).filter((n) => n.toLowerCase().endsWith('.xlsx'));
  if (names.length !== 1) {
    throw new Error(`expected exactly one .xlsx in ${dir}, found ${names.length}`);
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(join(dir, names[0]));
  return { wb, file: names[0] };
}

/**
 * Which shape of template this is.
 *
 * One kind of download can arrive as more than one sheet: the overrides
 * template is a MarketPlace Layer sheet or a MetaRisk Treaty sheet, with
 * different columns and a different key, chosen by what was being overridden. A
 * descriptor names its variants and the workbook says which one it is.
 */
export function resolveVariant(wb, descriptor) {
  if (!descriptor.variants) {
    const ws = wb.getWorksheet(descriptor.sheet);
    if (!ws) throw new Error(`no "${descriptor.sheet}" sheet in the workbook`);
    return { spec: descriptor, ws };
  }
  for (const variant of descriptor.variants) {
    const ws = wb.getWorksheet(variant.sheet);
    if (ws) return { spec: { ...descriptor, ...variant }, ws };
  }
  throw new Error(
    `no sheet matching any variant (${descriptor.variants.map((v) => v.sheet).join(', ')}); `
    + `the workbook has ${wb.worksheets.map((w) => w.name).join(', ')}`,
  );
}

function reader(ws, headerRow, marker) {
  const at = new Map();

  // A name can appear twice -- the ROLePlay block repeats the program identity
  // columns it is keyed on. Lookup is by name, so the first wins and the second
  // is unreachable; that is workable but it must not be silent, because a
  // shadowed column is one nothing can ever check.
  const duplicates = new Map();

  ws.getRow(headerRow).eachCell({ includeEmpty: false }, (c, n) => {
    const name = String(c.value ?? '').trim();
    if (!name) return;
    if (at.has(name)) {
      if (!duplicates.has(name)) duplicates.set(name, [ws.getRow(headerRow).getCell(at.get(name)).address]);
      duplicates.get(name).push(c.address);
      return;
    }
    at.set(name, n);
  });

  const value = (row, name) => {
    const n = at.get(name);
    if (!n) return undefined;
    let v = ws.getRow(row).getCell(n).value;
    if (v && typeof v === 'object') {
      if (v.result !== undefined) v = v.result;
      else if (v.richText) v = v.richText.map((t) => t.text).join('');
      else if (v.text !== undefined) v = v.text;
    }
    return v;
  };

  // Which rows hold data. Decided by one column that every real row fills --
  // named by the descriptor, because it is different in every template and
  // guessing it wrong silently shortens the comparison.
  if (!at.has(marker)) {
    throw new Error(
      `the row marker "${marker}" is not a column in this sheet; `
      + `row ${headerRow} holds ${at.size} column(s)`,
    );
  }
  const rows = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const first = value(r, marker);
    if (first !== undefined && first !== null && first !== '') rows.push(r);
  }
  return { value, rows, columns: at, duplicates };
}

/**
 * Excel keeps 15 significant digits, so a figure rebuilt in a different order
 * lands a few of those away with nothing having changed. Judged in proportion
 * rather than in units: one absolute figure cannot serve a report holding both
 * 0.0002 and 1.3e11.
 */
function same(a, b) {
  const norm = (v) => {
    if (v == null || v === '') return '';
    if (v instanceof Date) return v.toISOString();
    return typeof v === 'number' ? v : String(v).trim();
  };
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

/** One source against the template. Returns findings, empty when it agrees. */
export function compareSource(ws, spec, source, data) {
  const { value, rows, columns: templateColumns } = reader(ws, spec.headerRow, spec.rowMarker);

  // Both may be functions of the capture and of the sheet's own column names.
  // The overrides capture states its own header map, so listing those columns
  // again in the descriptor would be transcribing what the file already says --
  // and a field the capture stops carrying is caught by coverage instead, which
  // is where it belongs.
  const names = new Set(templateColumns.keys());
  const columns = typeof source.columns === 'function' ? source.columns(data, names) : source.columns;
  const expected = source.project(data, names);

  // A rendered page collapses runs of whitespace before it paints, so a source
  // read off one cannot see the difference between "A  B" and "A B". A request
  // body can, and is compared exactly.
  const seen = source.collapseWhitespace ? collapse : (v) => v;

  const byKey = new Map();
  for (const r of rows) byKey.set(source.key((name) => value(r, name)), r);

  const findings = [];

  for (const key of expected.keys()) {
    if (!byKey.has(key)) findings.push({ key, problem: 'row missing from the template' });
  }
  for (const key of byKey.keys()) {
    if (!expected.has(key)) {
      findings.push({ key, problem: `row in the template that ${source.label} does not have` });
    }
  }

  for (const [key, want] of expected) {
    const r = byKey.get(key);
    if (!r) continue;
    const cell = (name) => value(r, name);
    for (const column of columns) {
      const name = nameOf(column);
      const got = value(r, name);
      const mine = want[name];
      const agrees = column.compare
        ? column.compare(got, mine, cell)
        : same(seen(got), seen(mine));
      if (!agrees) findings.push({ key, column: name, template: got, source: mine });
    }
  }

  return { rows: expected.size, columns: columns.length, names: columns.map(nameOf), findings };
}

/**
 * Columns the template works out for itself.
 *
 * "Edison Client Level GeoScope differs from Property COE GeoScope" is a Yes/No
 * the sheet computes from two of its own columns. No source carries it, so
 * nothing else can check it -- and a template that gets it wrong tells the user
 * a field was left alone when it was changed.
 */
export function checkDerived(ws, spec) {
  if (!spec.derived?.length) return { ok: true, checked: 0, findings: [] };
  const { value, rows, columns } = reader(ws, spec.headerRow, spec.rowMarker);
  const findings = [];
  let checked = 0;

  for (const rule of spec.derived) {
    if (!columns.has(rule.column)) {
      findings.push({ column: rule.column, problem: 'column not found in the template' });
      continue;
    }
    const absent = rule.from.filter((n) => !columns.has(n));
    if (absent.length) {
      findings.push({ column: rule.column, problem: `needs ${absent.join(' and ')}, not in this sheet` });
      continue;
    }
    for (const r of rows) {
      checked++;
      const want = rule.value(...rule.from.map((n) => value(r, n)));
      const got = value(r, rule.column);
      if (String(got ?? '').trim() !== String(want ?? '').trim()) {
        findings.push({ row: r, column: rule.column, template: got, expected: want });
      }
    }
  }
  return { ok: findings.length === 0, checked, findings };
}

/**
 * Blocks of columns that are present together or not at all.
 *
 * Downloading a program selection template asks "Would you like to include
 * ROLePlay data in the template for advanced filtration?", and the answer adds
 * or omits a divider and the 144 columns behind it. So the block is a choice
 * rather than a property of the data, and that makes it checkable: all of it,
 * or none of it. Half a block is a template that lost columns on the way out,
 * and it would otherwise pass -- the columns it kept are all correct.
 */
export function checkBlocks(ws, spec) {
  const blocks = Object.entries(spec.blocks ?? {});
  if (!blocks.length) return { ok: true, present: [], findings: [] };
  const { columns } = reader(ws, spec.headerRow, spec.rowMarker);
  const findings = [];
  const present = [];

  // A name can belong to the base sheet AND to the block -- which is exactly
  // why it appears twice when both are there. Those names say nothing about
  // whether the block was included, so the lead column is what decides.
  const shared = new Set(spec.duplicateHeaders ?? []);

  for (const [name, block] of blocks) {
    const included = columns.has(block.lead);

    if (!included) {
      // Nothing of the block should have come through on its own. Names it
      // shares with the base sheet are skipped, being indistinguishable.
      const strays = block.columns.filter((c) => !shared.has(c) && columns.has(c));
      if (strays.length) {
        findings.push({
          block: name,
          problem: `not included -- "${block.lead}" is absent -- yet ${strays.length} of its `
            + 'column(s) are here anyway',
          missing: strays,
        });
      }
      continue;
    }

    present.push({ name, columns: block.columns.length + 1 });
    const absent = block.columns.filter((c) => !columns.has(c));
    if (absent.length) {
      findings.push({
        block: name,
        problem: `included -- "${block.lead}" is here -- but ${absent.length} of its `
          + `${block.columns.length} column(s) are not; a block arrives whole or not at all`,
        missing: absent,
      });
    }
  }

  return { ok: findings.length === 0, present, findings };
}

/**
 * How much of the template anybody actually looked at.
 *
 * Without this a case prints "31 columns clean" over a sheet holding 176, and
 * the 145 nobody checked read exactly like 145 that passed. Columns a source
 * cannot see are declared in the descriptor with a reason; anything else left
 * unchecked is a finding, so a column added to a future release is noticed the
 * first time it appears rather than the first time it is wrong.
 */
export function checkCoverage(ws, spec, compared) {
  const { columns, duplicates } = reader(ws, spec.headerRow, spec.rowMarker);
  const checked = new Set();
  for (const names of compared) for (const name of names) checked.add(name);
  for (const d of spec.derived ?? []) checked.add(d.column);

  const excused = new Map();
  for (const [reason, names] of Object.entries(spec.unverifiable ?? {})) {
    for (const n of names) excused.set(n, reason);
  }

  const unchecked = [];
  const declared = [];
  for (const name of columns.keys()) {
    if (checked.has(name)) continue;
    (excused.has(name) ? declared : unchecked).push({ column: name, reason: excused.get(name) });
  }

  // A duplicate the descriptor knows about is a note; one it does not is a
  // finding, because the second column of that name is unreachable and nothing
  // else in this tool would ever say so.
  const known = new Set(spec.duplicateHeaders ?? []);
  const shadowed = [...duplicates]
    .filter(([name]) => !known.has(name))
    .map(([name, cells]) => ({ column: name, cells }));

  return {
    ok: unchecked.length === 0 && shadowed.length === 0,
    total: columns.size + [...duplicates.values()].reduce((t, c) => t + c.length - 1, 0),
    distinct: columns.size,
    checked: [...columns.keys()].filter((n) => checked.has(n)).length,
    duplicates: [...duplicates].map(([name, cells]) => ({ column: name, cells })),
    shadowed,
    declared,
    unchecked,
  };
}

/** Everything one case has to say. */
export async function compareCase(caseDir, descriptor) {
  const { readFile } = await import('node:fs/promises');
  const { existsSync } = await import('node:fs');

  const { wb, file } = await openTemplate(join(caseDir, 'template'));
  const { spec, ws } = resolveVariant(wb, descriptor);
  const results = [];
  const used = [];

  for (const source of spec.sources) {
    const path = join(caseDir, 'data', source.file);
    if (!existsSync(path)) {
      results.push({ name: source.name, skipped: `no ${source.file}` });
      continue;
    }
    const data = JSON.parse(await readFile(path, 'utf8'));

    // A capture can say what it is a capture of. When it does, and it disagrees
    // with the sheet in front of us, the two were not taken from the same
    // download -- and comparing them would produce a page of differences that
    // all have one cause.
    if (source.declares) {
      const said = source.declares(data);
      if (said && spec.declared && said !== spec.declared) {
        results.push({
          name: source.name,
          skipped: `${source.file} was captured for "${said}", this template is "${spec.declared}"`,
        });
        continue;
      }
    }

    // A capture of a page that was only showing part of its rows is a partial
    // source. The row comparison would still fire -- as "rows the template has
    // that the screen does not", which reads as the template being wrong when
    // the capture was simply short. Named here instead, and the columns it
    // would have covered are then reported as uncovered, which they are.
    if (source.complete && !source.complete(data)) {
      results.push({
        name: source.name,
        skipped: `${source.file} caught the page mid-list (${data.page.shown} of ${data.page.total} rows)`,
      });
      continue;
    }

    const outcome = compareSource(ws, spec, source, data);
    used.push(outcome.names);
    results.push({ name: source.name, ...outcome });
  }

  const fills = spec.fills ? checkFills(ws, spec.fills, spec.headerRow) : { ok: true, findings: [] };
  const markers = spec.markers
    ? checkMarkers(ws, spec.markers, spec.headerRow)
    : { ok: true, checked: 0, findings: [] };
  const painted = spec.conditionalFills
    ? checkConditionalFills(ws, spec.conditionalFills, spec.headerRow)
    : { ok: true, checked: 0, findings: [] };
  const derived = checkDerived(ws, spec);
  const blocks = checkBlocks(ws, spec);
  const coverage = checkCoverage(ws, spec, used);

  const failed = results.reduce((n, r) => n + (r.findings?.length ?? 0), 0)
    + fills.findings.length + markers.findings.length + painted.findings.length
    + derived.findings.length + blocks.findings.length
    + coverage.unchecked.length + coverage.shadowed.length;

  return {
    file, sheet: ws.name, results, fills, markers, painted, derived, blocks, coverage,
    ok: failed === 0,
  };
}
