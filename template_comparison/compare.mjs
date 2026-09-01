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
import { checkFills } from './check-fills.mjs';

/** Opens the one .xlsx in a folder. */
export async function openTemplate(dir, sheet) {
  const names = (await readdir(dir)).filter((n) => n.toLowerCase().endsWith('.xlsx'));
  if (names.length !== 1) {
    throw new Error(`expected exactly one .xlsx in ${dir}, found ${names.length}`);
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(join(dir, names[0]));
  const ws = wb.getWorksheet(sheet);
  if (!ws) throw new Error(`no "${sheet}" sheet in ${names[0]}`);
  return { ws, file: names[0] };
}

function reader(ws, headerRow) {
  const at = new Map();
  ws.getRow(headerRow).eachCell({ includeEmpty: false }, (c, n) => {
    const name = String(c.value ?? '').trim();
    if (name && !at.has(name)) at.set(name, n);
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

  const rows = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const first = value(r, 'Treaty Name');
    if (first !== undefined && first !== null && first !== '') rows.push(r);
  }
  return { value, rows };
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
export function compareSource(ws, headerRow, source, data) {
  const { value, rows } = reader(ws, headerRow);
  const expected = source.project(data);

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
    for (const column of source.columns) {
      const got = value(r, column);
      if (!same(got, want[column])) {
        findings.push({ key, column, template: got, source: want[column] });
      }
    }
  }

  return { rows: expected.size, columns: source.columns.length, findings };
}

/** Everything one case has to say. */
export async function compareCase(caseDir, descriptor) {
  const { readFile } = await import('node:fs/promises');
  const { existsSync } = await import('node:fs');

  const { ws, file } = await openTemplate(join(caseDir, 'template'), descriptor.sheet);
  const results = [];

  for (const source of descriptor.sources) {
    const path = join(caseDir, 'data', source.file);
    if (!existsSync(path)) {
      results.push({ name: source.name, skipped: `no ${source.file}` });
      continue;
    }
    const data = JSON.parse(await readFile(path, 'utf8'));
    results.push({ name: source.name, ...compareSource(ws, descriptor.headerRow, source, data) });
  }

  const fills = descriptor.fills
    ? checkFills(ws, descriptor.fills, descriptor.headerRow)
    : { ok: true, findings: [] };

  const failed = results.reduce((n, r) => n + (r.findings?.length ?? 0), 0) + fills.findings.length;
  return { file, results, fills, ok: failed === 0 };
}
