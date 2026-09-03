/**
 * What the captured cases have never exercised.
 *
 *   npm run gaps:templates              across ./edison_output_comparison/templates
 *   npm run gaps:templates -- <folder>  across somewhere else
 *
 * A green run says the cases that exist all pass. It says nothing about the
 * ones that do not exist, and "we have cases" reads exactly like "we have
 * coverage" until somebody checks. This reads every case of a kind together and
 * reports what none of them contains -- which is the list of downloads still
 * worth capturing.
 *
 * Everything here is measured from the files. Nothing is declared, so nothing
 * goes stale: as cases are added the gaps close on their own, and a column that
 * starts varying stops being reported without anyone editing a list.
 *
 * Three questions, in the order they are worth answering:
 *
 *   1. A column carries a list rule -- "Yes,No" -- and only one of its options
 *      has ever appeared. The other is a path through the app that no capture
 *      has ever taken.
 *   2. A column is empty in every row of every case. Nothing it does has been
 *      seen at all.
 *   3. A column holds the same single value everywhere. Its comparison has
 *      never had to distinguish anything, so it would pass while broken.
 *
 * This never fails a build. It is a worklist, not a verdict.
 */
import { readdir, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openTemplate, resolveVariant, folders } from './compare.mjs';

const args = process.argv.slice(2);
const ROOT = args.find((a) => !a.startsWith('-')) ?? 'edison_output_comparison/templates';
const HERE = fileURLToPath(new URL('.', import.meta.url));

const dirs = async (p) =>
  (await readdir(p, { withFileTypes: true })).filter((e) => e.isDirectory()).map((e) => e.name);

const text = (v) => {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.result !== undefined) return String(v.result);
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if (v.text !== undefined) return String(v.text);
    if (v instanceof Date) return v.toISOString().slice(0, 10);
  }
  return String(v);
};

/** `"Yes,No"` -> ['Yes', 'No'] */
const listItems = (f) => String(f ?? '').trim().replace(/^"(.*)"$/s, '$1')
  .split(',').map((x) => x.trim()).filter(Boolean);

const n = (v) => v.toLocaleString('en-US');

if (!existsSync(ROOT)) {
  console.error(`template-verify: no folder at ${ROOT}`);
  process.exit(1);
}

const descriptorFor = (kind) => join(HERE, 'templates', kind, 'index.mjs');
const named = basename(resolve(ROOT));
const kinds = existsSync(descriptorFor(named))
  ? [{ kind: named, dir: ROOT }]
  : (await dirs(ROOT)).map((kind) => ({ kind, dir: join(ROOT, kind) }));

for (const { kind, dir } of kinds) {
  if (!existsSync(descriptorFor(kind))) continue;
  const descriptor = (await import(`./templates/${kind}/index.mjs`)).default;

  // One bucket per sheet shape. A MarketPlace Layer case says nothing about
  // what a MetaRisk Treaty case has covered, so they are never pooled.
  const shapes = new Map();

  for (const name of await dirs(dir)) {
    const caseDir = join(dir, name);
    const where = folders(caseDir);
    if (!where) continue;

    const { wb } = await openTemplate(where.current);
    const { spec, ws } = resolveVariant(wb, descriptor);

    const at = new Map();
    ws.getRow(spec.headerRow).eachCell({ includeEmpty: false }, (c, i) => {
      const k = text(c.value).trim();
      if (k && !at.has(k)) at.set(k, i);
    });
    const rows = [];
    for (let r = spec.headerRow + 1; r <= ws.rowCount; r++) {
      const v = text(ws.getRow(r).getCell(at.get(spec.rowMarker)).value);
      if (v !== '') rows.push(r);
    }

    if (!shapes.has(ws.name)) {
      shapes.set(ws.name, {
        cases: [], seen: new Map(), allowed: new Map(), rowCounts: [],
        // The overrides sheet titles an editable column "X *".
        markers: (spec.markers?.columns ?? []).map((c) => `${c}${spec.markers.suffix}`),
      });
    }
    const shape = shapes.get(ws.name);
    shape.cases.push(name);
    shape.rowCounts.push(rows.length);

    for (const [column, i] of at) {
      if (!shape.seen.has(column)) shape.seen.set(column, new Set());
      const bag = shape.seen.get(column);
      for (const r of rows) bag.add(text(ws.getRow(r).getCell(i).value));

      // What the sheet itself says this column is allowed to hold.
      if (!shape.allowed.has(column) && rows.length) {
        const dv = ws.getRow(rows[0]).getCell(i).dataValidation;
        if (dv?.type === 'list') shape.allowed.set(column, listItems(dv.formulae?.[0]));
      }
    }

    // A payload flag that only ever has one value is a branch never taken.
    const payload = join(caseDir, 'data', 'payload_data.json');
    if (existsSync(payload)) {
      const data = JSON.parse(await readFile(payload, 'utf8'));
      for (const [k, v] of Object.entries(data)) {
        if (typeof v !== 'boolean') continue;
        const key = `payload: ${k}`;
        if (!shape.seen.has(key)) shape.seen.set(key, new Set());
        shape.seen.get(key).add(String(v));
      }
    }
  }

  for (const [sheet, shape] of shapes) {
    console.log(`\n${'='.repeat(72)}`);
    console.log(`${kind} · ${sheet}`);
    console.log(`${shape.cases.length} case(s): ${shape.cases.join(', ')}`);
    console.log(`rows per case: ${shape.rowCounts.sort((a, b) => a - b).join(', ')}`);

    const unusedOption = [];
    const neverFilled = [];
    const neverVaries = [];

    // The columns a user may edit are the ones a gap actually costs something
    // on: an untouched editable column is a field the round trip has never
    // carried. Everything else is read-only and a constant there is usually
    // just what the test data happens to be.
    // A fill on the DATA rows says "you may edit this"; one on the header row
    // marks the column itself -- a divider, or a header band. Only the first
    // kind is a field somebody fills in, which is what a gap here is about.
    const editable = new Set([
      ...Object.values(descriptor.fills ?? {})
        .filter((g) => g.row !== 'header')
        .flatMap((g) => g.columns ?? []),
      ...(shape.markers ?? []),
    ]);

    for (const [column, bag] of shape.seen) {
      const values = [...bag];
      const real = values.filter((v) => v !== '');

      const allowed = shape.allowed.get(column);
      if (allowed?.length) {
        const missing = allowed.filter((o) => !real.includes(o));
        if (missing.length) unusedOption.push({ column, missing, allowed, seen: real });
      }

      if (!real.length) { neverFilled.push(column); continue; }
      if (new Set(real).size === 1 && !column.startsWith('payload: ')) {
        neverVaries.push({ column, value: real[0] });
      }
      if (column.startsWith('payload: ') && new Set(real).size === 1) {
        neverVaries.push({ column, value: real[0] });
      }
    }

    // Said first and said short, because a worklist nobody reads to the end is
    // a worklist that changes nothing.
    const priority = [];
    for (const u of unusedOption) {
      priority.push(`${u.column}: the sheet allows ${u.allowed.join('/')}, only ${u.seen.join('/')} seen`);
    }
    for (const c of neverFilled) {
      if (editable.has(c)) priority.push(`${c}: editable, and empty in every row of every case`);
    }
    for (const c of neverVaries) {
      if (editable.has(c.column)) {
        priority.push(`${c.column}: editable, always ${JSON.stringify(c.value).slice(0, 30)}`);
      }
      if (c.column.startsWith('payload: ')) {
        priority.push(`${c.column}: always ${c.value} -- the other branch is never taken`);
      }
    }

    if (priority.length) {
      console.log(`
  WORTH CAPTURING (${n(priority.length)}):`);
      for (const p2 of priority) console.log(`    * ${p2}`);
    }

    if (unusedOption.length) {
      console.log(`\n  The sheet allows a value no case has ever held (${n(unusedOption.length)}):`);
      for (const u of unusedOption) {
        console.log(`    ${u.column}  allows ${u.allowed.join('/')}, only ever ${u.seen.join('/')}`);
      }
    }

    if (neverFilled.length) {
      console.log(`\n  Empty in every row of every case (${n(neverFilled.length)}):`);
      for (const c of neverFilled.slice(0, 20)) console.log(`    ${c}`);
      if (neverFilled.length > 20) console.log(`    ... and ${n(neverFilled.length - 20)} more`);
    }

    if (neverVaries.length) {
      console.log(`\n  One value everywhere, so nothing has had to be distinguished (${n(neverVaries.length)}):`);
      for (const c of neverVaries.slice(0, 20)) {
        console.log(`    ${c.column} = ${JSON.stringify(c.value).slice(0, 50)}`);
      }
      if (neverVaries.length > 20) console.log(`    ... and ${n(neverVaries.length - 20)} more`);
    }

    if (!unusedOption.length && !neverFilled.length && !neverVaries.length) {
      console.log('\n  Every column varies and every allowed value has been seen.');
    }
  }
}

console.log('');
console.log('A worklist, not a verdict -- nothing here fails a run.');
