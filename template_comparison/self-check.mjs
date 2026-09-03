/**
 * Checks the checks, by breaking a template on purpose.
 *
 *   npm run check:templates              against ./edison_output_comparison/templates
 *   npm run check:templates -- <folder>  against somewhere else
 *
 * A comparison that reports nothing is either a clean template or a broken
 * check, and the two read identically. This plants one defect at a time in a
 * throwaway copy of a real case and confirms the run notices -- so "0 failing"
 * means something was actually looked at.
 *
 * Nothing here knows a column name. Every fault is built from the descriptor
 * and from whatever the sheet turns out to hold, so it works on a template
 * added tomorrow without being taught anything. Faults that do not apply to a
 * case -- no fills, no derived columns, no optional block -- are skipped rather
 * than counted as passes.
 *
 * The copies live in the system temp folder and are removed afterwards. The
 * case itself is opened read-only and never written to.
 */
import { readdir, mkdtemp, cp, rm } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { join, basename, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import { compareCase, resolveVariant, openTemplate, folders } from './compare.mjs';

const args = process.argv.slice(2);
const ROOT = args.find((a) => !a.startsWith('-')) ?? 'edison_output_comparison/templates';
const HERE = fileURLToPath(new URL('.', import.meta.url));

const dirs = async (p) =>
  (await readdir(p, { withFileTypes: true })).filter((e) => e.isDirectory()).map((e) => e.name);

const text = (v) => {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if (v.result !== undefined) return String(v.result);
    if (v.text !== undefined) return String(v.text);
  }
  return String(v);
};

/** Column name -> column number, from the header row. */
function headers(ws, headerRow) {
  const at = new Map();
  ws.getRow(headerRow).eachCell({ includeEmpty: false }, (c, n) => {
    const name = text(c.value).trim();
    if (name && !at.has(name)) at.set(name, n);
  });
  return at;
}

/**
 * The faults, each described by what it breaks and which check should say so.
 *
 * `plant` returns a label when it applies and null when it does not -- a
 * template with no derived columns cannot lose one, and pretending otherwise
 * would turn an inapplicable case into a passing one.
 *
 * `count` measures the area the fault belongs to, and a fault counts as noticed
 * when it makes that number go up. Judged against the case's own baseline
 * rather than against zero, so a case that already has a real finding -- as the
 * program selection templates do, their own data validations rejecting their
 * own values -- can still be used to check that the other checks work.
 */
const FAULTS = [
  {
    name: 'a value is changed',
    count: (o) => o.results.reduce((t, r) => t + (r.findings ?? []).filter((f) => f.column).length, 0),
    plant(ws, spec, checked) {
      const first = checked.find((c) => at(ws, spec).has(c));
      if (!first) return null;
      const row = spec.headerRow + 1;
      const cell = ws.getRow(row).getCell(at(ws, spec).get(first));
      cell.value = `sheet-verify planted ${Date.now()}`;
      return `"${first}" overwritten`;
    },
  },
  {
    name: 'a row is dropped',
    count: (o) => o.results.reduce(
      (t, r) => t + (r.findings ?? []).filter((f) => /row missing/.test(f.problem ?? '')).length, 0),
    plant(ws, spec) {
      ws.spliceRows(spec.headerRow + 1, 1);
      return 'first data row removed';
    },
  },
  {
    name: 'an editable column loses its fill',
    count: (o) => o.fills.findings.length,
    plant(ws, spec) {
      const group = Object.values(spec.fills ?? {}).find((g) => g.columns?.length);
      if (!group) return null;
      const found = group.columns.find((c) => at(ws, spec).has(c));
      if (!found) return null;
      const row = group.row === 'header' ? spec.headerRow : spec.headerRow + 1;
      ws.getRow(row).getCell(at(ws, spec).get(found)).fill = {
        type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDEADBE' },
      };
      return `"${found}" repainted`;
    },
  },
  {
    name: 'an editable column loses its marker',
    count: (o) => o.markers.findings.length,
    plant(ws, spec) {
      if (!spec.markers) return null;
      const { suffix, columns } = spec.markers;
      const found = columns.map((c) => `${c}${suffix}`).find((c) => at(ws, spec).has(c));
      if (!found) return null;
      ws.getRow(spec.headerRow).getCell(at(ws, spec).get(found)).value = found.slice(0, -suffix.length);
      return `"${found}" unmarked`;
    },
  },
  {
    name: 'a column the sheet computes is wrong',
    count: (o) => o.derived.findings.length,
    plant(ws, spec) {
      const rule = (spec.derived ?? []).find((d) => at(ws, spec).has(d.column));
      if (!rule) return null;
      const cell = ws.getRow(spec.headerRow + 1).getCell(at(ws, spec).get(rule.column));
      cell.value = `${text(cell.value)}!`;
      return `"${rule.column}" altered`;
    },
  },
  {
    name: 'an optional block arrives half-written',
    count: (o) => o.blocks.findings.length,
    plant(ws, spec) {
      for (const block of Object.values(spec.blocks ?? {})) {
        if (!at(ws, spec).has(block.lead)) continue;
        const found = block.columns.find((c) => at(ws, spec).has(c));
        if (!found) continue;
        ws.getRow(spec.headerRow).getCell(at(ws, spec).get(found)).value = null;
        return `"${found}" removed from an included block`;
      }
      return null;
    },
  },
  {
    name: 'the rule that paints editable columns is dropped',
    count: (o) => o.painted.findings.length,
    plant(ws, spec) {
      if (!spec.conditionalFills || !ws.conditionalFormattings?.length) return null;
      const gone = ws.conditionalFormattings.length;
      ws.conditionalFormattings = [];
      return `${gone} conditional format(s) removed`;
    },
  },
  {
    // The one the old check could not see: strip the block entirely and the
    // sheet is perfectly self-consistent -- no divider, no columns. Only the
    // request knows it was asked for.
    name: 'a block that was asked for is missing altogether',
    count: (o) => o.blocks.findings.length,
    plant(ws, spec) {
      for (const block of Object.values(spec.blocks ?? {})) {
        if (!block.requested || !at(ws, spec).has(block.lead)) continue;
        const cols = at(ws, spec);
        for (const name of [block.lead, ...block.columns]) {
          if (cols.has(name)) ws.getRow(spec.headerRow).getCell(cols.get(name)).value = null;
        }
        return 'the whole block removed, divider included';
      }
      return null;
    },
  },
  {
    name: "a value stops satisfying the sheet's own rule",
    count: (o) => o.validations.findings.length,
    plant(ws, spec) {
      const cols = at(ws, spec);
      for (const row of [spec.headerRow + 1, spec.headerRow + 2]) {
        for (const [, n] of cols) {
          const cell = ws.getRow(row).getCell(n);
          const dv = cell.dataValidation;
          if (!dv || dv.type === 'custom') continue;
          cell.value = dv.type === 'list' ? 'Perhaps' : -999999;
          return `${cell.address} set to something its ${dv.type} rule rejects`;
        }
      }
      return null;
    },
  },
  {
    // The one only the golden can see: change a column no source speaks for.
    // Every other check passes, because none of them is looking there.
    name: 'a column no source can verify drifts',
    count: (o) => o.baseline?.findings.length ?? 0,
    plant(ws, spec, checked) {
      const known = new Set([...checked, ...(spec.derived ?? []).map((d) => d.column)]);
      for (const [name, n] of at(ws, spec)) {
        if (known.has(name)) continue;
        const cell = ws.getRow(spec.headerRow + 1).getCell(n);
        if (cell.value === null || cell.value === undefined || cell.value === '') continue;
        cell.value = typeof cell.value === 'number' ? cell.value + 1 : `${text(cell.value)} drifted`;
        return `"${name}" changed, which no source checks`;
      }
      return null;
    },
  },
  {
    // Nothing in this tool models a merged cell. The archive comparison does
    // not model it either -- it just notices the bytes moved, which is the
    // whole point of having it.
    name: 'something no check understands changes',
    count: (o) => (o.baseline?.parts?.findings.length ?? 0),
    plant(ws, spec) {
      ws.mergeCells(`A${spec.headerRow + 14}:B${spec.headerRow + 14}`);
      return 'a merged cell added, which no check here models';
    },
  },
  {
    name: 'a column nobody checks appears',
    count: (o) => o.coverage.unchecked.length,
    plant(ws, spec) {
      ws.getRow(spec.headerRow).getCell(ws.columnCount + 2).value = 'Planted Unchecked Column';
      return 'one new header written past the last column';
    },
  },
];

/** Header lookup, recomputed each time because planting can move things. */
const at = (ws, spec) => headers(ws, spec.headerRow);

async function checkCase(kind, name, caseDir, descriptor) {
  const { wb, file } = await openTemplate(folders(caseDir).current);
  const { spec, ws } = resolveVariant(wb, descriptor);

  // Which columns a source actually compares, so the "value changed" fault
  // lands somewhere a check is looking rather than somewhere it is not.
  const base = await compareCase(caseDir, descriptor);
  const checked = base.results.flatMap((r) => r.names ?? []);

  const results = [];
  for (const fault of FAULTS) {
    const work = await mkdtemp(join(tmpdir(), 'sheet-verify-selfcheck-'));
    try {
      await cp(caseDir, join(work, 'case'), { recursive: true });
      await rm(join(work, 'case', 'results'), { recursive: true, force: true });

      const at2 = folders(join(work, 'case'));
      const path = join(at2.current, file);

      // Planting rewrites `current` through ExcelJS, which changes the file's
      // producer. Rewrite the golden the same way, unchanged, so the archive
      // comparison is between two ExcelJS files and reports the planted
      // difference instead of reporting the two writers.
      if (at2.golden && existsSync(at2.golden)) {
        const g = await openTemplate(at2.golden);
        await g.wb.xlsx.writeFile(join(at2.golden, g.file));
      }

      const copy = new ExcelJS.Workbook();
      await copy.xlsx.readFile(path);
      const sheet = resolveVariant(copy, descriptor).ws;

      const note = fault.plant(sheet, spec, checked);
      if (note === null) { results.push({ fault: fault.name, applies: false }); continue; }
      await copy.xlsx.writeFile(path);

      let outcome;
      try {
        outcome = await compareCase(join(work, 'case'), descriptor);
      } catch {
        // A fault severe enough to stop the run is still a fault that was noticed.
        results.push({ fault: fault.name, applies: true, caught: true, note });
        continue;
      }
      const before = fault.count(base);
      const after = fault.count(outcome);
      results.push({
        fault: fault.name, applies: true, caught: after > before, note,
        detail: after > before ? undefined : `area went ${before} -> ${after}`,
      });
    } finally {
      await rm(work, { recursive: true, force: true });
    }
  }
  return { results };
}

if (!existsSync(ROOT)) {
  console.error(`template-verify: no folder at ${ROOT}`);
  process.exit(1);
}

const descriptorFor = (kind) => join(HERE, 'templates', kind, 'index.mjs');
const named = basename(resolve(ROOT));
const kinds = existsSync(descriptorFor(named))
  ? [{ kind: named, dir: ROOT }]
  : (await dirs(ROOT)).map((kind) => ({ kind, dir: join(ROOT, kind) }));

let planted = 0;
let missed = 0;
let cases = 0;

for (const { kind, dir } of kinds) {
  if (!existsSync(descriptorFor(kind))) continue;
  const descriptor = (await import(`./templates/${kind}/index.mjs`)).default;

  for (const name of await dirs(dir)) {
    const caseDir = join(dir, name);
    if (!folders(caseDir)) continue;
    cases++;

    const { skipped, results } = await checkCase(kind, name, caseDir, descriptor);
    if (skipped) { console.log(`- ${kind} · ${name}: skipped, ${skipped}`); continue; }

    const applied = results.filter((r) => r.applies);
    const bad = applied.filter((r) => !r.caught);
    planted += applied.length;
    missed += bad.length;

    console.log(`${bad.length ? '✗' : '✓'} ${kind} · ${name}`);
    for (const r of results) {
      if (!r.applies) console.log(`      –  ${r.fault} (does not apply here)`);
      else console.log(`      ${r.caught ? '✓' : '✗'}  ${r.fault} — ${r.note}${r.detail ? ` [${r.detail}]` : ''}`);
    }
  }
}

console.log('');
console.log(`${cases} case(s), ${planted} fault(s) planted, ${missed} not noticed`);
process.exit(missed ? 1 : 0);
