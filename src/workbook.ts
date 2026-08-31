import type {
  DiffResult, MovedSheet, ResolvedSpec, SheetModel, SheetOutcome, TableRequest, TableSpec,
  WorkbookDiffResult, WorkbookReader, WorkbookSheetSpec, WorkbookSpec,
} from './types.js';
import { resolveSpec, tableRange } from './model.js';
import { columnRange } from './a1.js';
import { compare } from './compare.js';
import { readerFor } from './verify.js';
import { metadataAddressesFor, parseMetadata } from './metadata.js';

/** Excel forbids two sheets whose names differ only in case, so this is safe. */
const canonSheet = (s: string): string => s.trim().toLowerCase();

/**
 * Folds `metadata` labels into `ignoreRows`, so layer 1 skips them too.
 *
 * A key-value block -- "Report ID" in column A, its value in column B -- is a
 * table like any other once a header is found above it, and its labels are its
 * row keys. Without this, layer 2 would list the report id as metadata while
 * layer 1 reported it as a value change: the same cell, two verdicts.
 *
 * Address-form entries are left out. They name a cell, not a row, and a stray
 * `"A2"` in a key column should not silently drop that row from comparison.
 */
function withMetadataRows(spec: WorkbookSpec): WorkbookSpec {
  const { labels } = parseMetadata(spec.metadata);
  if (!labels.length) return spec;

  const defaults = spec.defaults ?? {};
  const global = labels.filter((l) => !l.sheet).map((l) => l.label);
  const out: WorkbookSpec = {
    ...spec,
    ...(global.length
      ? { defaults: { ...defaults, ignoreRows: [...(defaults.ignoreRows ?? []), ...global] } }
      : {}),
  };

  // A label confined to one sheet stays confined here too, or qualifying it
  // would have bought nothing.
  const confined = labels.filter((l) => l.sheet);
  if (!confined.length) return out;

  const sheets = { ...(out.sheets ?? {}) };
  for (const { sheet, label } of confined) {
    const found = Object.keys(sheets).find((k) => canonSheet(k) === sheet) ?? sheet;
    const entry = sheets[found] ?? {};
    sheets[found] = { ...entry, ignoreRows: [...(entry.ignoreRows ?? []), label] };
  }
  return { ...out, sheets };
}

function isWorkbookReader(r: unknown): r is WorkbookReader {
  return typeof (r as WorkbookReader).readWorkbook === 'function';
}

export function workbookReaderFor(path: string): WorkbookReader {
  const reader = readerFor(path);
  if (!isWorkbookReader(reader)) {
    throw new Error(
      `sheet-verify: the reader for "${path}" cannot read a whole workbook. ` +
      'CSV holds a single table, so use verifySheet() or toMatchSheetBaseline ' +
      'for it; a custom reader must implement readWorkbook() to take part in ' +
      'workbook and case comparisons.',
    );
  }
  return reader;
}

const asRecord = (t: WorkbookSheetSpec['tolerance']): Record<string, number> =>
  typeof t === 'number' ? { '*': t } : { ...(t ?? {}) };

/**
 * Merges the workbook defaults with one sheet's overrides. Per-field, because
 * a blanket override is a footgun: a per-sheet `tolerance` replacing the
 * default record would silently drop the `*` fallback, and a per-sheet
 * `ignoreColumns` would drop the globally ignored timestamp columns.
 */
export function mergeSheetSpec(
  defaults: WorkbookSheetSpec | undefined,
  sheet: WorkbookSheetSpec | undefined,
): WorkbookSheetSpec {
  const d = defaults ?? {};
  const s = sheet ?? {};
  const tables = mergeTables(d.tables, s.tables);
  return {
    ...d,
    ...s,
    ...(tables ? { tables } : {}),
    tolerance: { ...asRecord(d.tolerance), ...asRecord(s.tolerance) },
    relativeTolerance: {
      ...asRecord(d.relativeTolerance), ...asRecord(s.relativeTolerance),
    },
    ignoreColumns: [...(d.ignoreColumns ?? []), ...(s.ignoreColumns ?? [])],
    ignoreRows: [...(d.ignoreRows ?? []), ...(s.ignoreRows ?? [])],
    invariants: [...(d.invariants ?? []), ...(s.invariants ?? [])],
  };
}

/**
 * Tables merge per name, for the same reason the fields above do: overriding
 * one table's tolerance should not delete every other table on the sheet.
 * Names match case-insensitively, as sheet names do.
 */
function mergeTables(
  a: Record<string, TableSpec> | undefined,
  b: Record<string, TableSpec> | undefined,
): Record<string, TableSpec> | undefined {
  if (!a && !b) return undefined;
  const out: Record<string, TableSpec> = { ...(a ?? {}) };
  for (const [name, override] of Object.entries(b ?? {})) {
    const found = Object.keys(out).find((k) => k.toLowerCase() === name.toLowerCase());
    const merged = mergeSheetSpec(found ? out[found] : undefined, override);
    delete (merged as WorkbookSheetSpec).tables;
    out[found ?? name] = merged;
  }
  return out;
}

/**
 * A sheet's spec, with its metadata addresses attached.
 *
 * Every path that resolves a sheet goes through here, because an address rule
 * that reaches layer 2 and not layer 1 puts the same cell in the report twice:
 * once under "what changed" and once under "not verified".
 */
function sheetSpecFor(spec: WorkbookSpec, sheet: string): WorkbookSheetSpec {
  const merged = mergeSheetSpec(spec.defaults, entryFor(spec.sheets, sheet));
  const cells = metadataAddressesFor(spec.metadata, sheet);
  if (!cells.length) return merged;
  return { ...merged, metadataCells: [...(merged.metadataCells ?? []), ...cells] };
}

/** Looks a sheet's entry up case-insensitively. */
function entryFor(
  sheets: WorkbookSpec['sheets'],
  name: string,
): WorkbookSheetSpec | undefined {
  if (!sheets) return undefined;
  const want = canonSheet(name);
  for (const [k, v] of Object.entries(sheets)) {
    if (canonSheet(k) === want) return v;
  }
  return undefined;
}

/**
 * Whether a table with no key should be compared by row position rather than
 * left unchecked. On by default: not comparing is the worse of the two, and
 * the report names every table matched this way.
 */
const positionalFallback = (spec: WorkbookSpec) => spec.matchUnkeyedRowsByPosition ?? true;

/**
 * Whether two tables occupy any column in common. Unbounded means the whole
 * width, so a table with no `columns` shares them with everything -- which is
 * right: on a sheet nobody has split by column, every table is under the last.
 */
function sharesColumns(a: WorkbookSheetSpec, b: WorkbookSheetSpec): boolean {
  if (!a.columns || !b.columns) return true;
  const x = columnRange(a.columns, 0);
  const y = columnRange(b.columns, 0);
  return x.from <= y.to && y.from <= x.to;
}

/** Resolves one sheet's spec, or null when it has no key and cannot be compared. */
export function resolveSheetSpec(spec: WorkbookSpec, sheet: string): ResolvedSpec | null {
  const merged = sheetSpecFor(spec, sheet);
  if (!merged.keyColumns?.length) {
    if (!positionalFallback(spec)) return null;
    return resolveSpec({ ...merged, keyColumns: [], matchRowsByPosition: true, sheet });
  }
  return resolveSpec({ ...merged, keyColumns: merged.keyColumns, sheet });
}

/**
 * Separates the two halves of a model key. Excel forbids control characters in
 * sheet names, so no sheet name can contain one and collide with a key built
 * from a different pair. Written as an escape rather than a literal character,
 * which would make this source file read as binary to grep and diff.
 */
const TABLE_KEY_SEP = '\u0000';

/** Model key. Never shown; `label` is what reports print. */
const tableKey = (sheet: string, table: string) =>
  `${canonSheet(sheet)}${TABLE_KEY_SEP}${table}`;

const labelFor = (sheet: string, table: string) =>
  table === sheet ? sheet : `${sheet} · ${table}`;

export interface ResolvedTable {
  table: string;
  key: string;
  label: string;
  spec: ResolvedSpec | null;
  reason?: string;
}

/**
 * Resolves every table on a sheet. A sheet declaring `tables` yields one entry
 * per table, each bounded by the next table's header row so an info block
 * cannot run on into the data table below it. Otherwise the sheet is one
 * table named after itself.
 */
export function resolveTables(spec: WorkbookSpec, sheet: string): ResolvedTable[] {
  const merged = sheetSpecFor(spec, sheet);
  const declared = Object.entries(merged.tables ?? {});

  if (!declared.length) {
    const resolved = resolveSheetSpec(spec, sheet);
    return [{
      table: sheet,
      key: tableKey(sheet, sheet),
      label: sheet,
      spec: resolved,
      reason: resolved
        ? (resolved.matchRowsByPosition ? 'rows matched by position — no row key found' : undefined)
        : 'no keyColumns configured for this sheet',
    }];
  }

  // `tables` itself must not be inherited down into each table.
  const { tables: _drop, ...sheetLevel } = merged;

  const ordered = declared
    .map(([table, t]) => ({ table, merged: mergeSheetSpec(sheetLevel, t) }))
    .sort((a, b) => (a.merged.headerRow ?? 1) - (b.merged.headerRow ?? 1));

  return ordered.map((entry, i) => {
    // The next table *in the same columns*. A sheet that prints two tables
    // side by side has one starting below another without being under it, and
    // bounding a table at a neighbour's header row would cut it off at the
    // neighbour's first row for no reason.
    const next = ordered.slice(i + 1).find((o) => sharesColumns(entry.merged, o.merged));
    // An explicit endRow wins; otherwise stop just above the next table.
    const endRow = entry.merged.endRow ?? (next ? (next.merged.headerRow ?? 1) - 1 : 0);
    const keyColumns = entry.merged.keyColumns;
    const byPosition = !keyColumns?.length && positionalFallback(spec);

    return {
      table: entry.table,
      key: tableKey(sheet, entry.table),
      label: labelFor(sheet, entry.table),
      spec: keyColumns?.length
        ? resolveSpec({ ...entry.merged, keyColumns, endRow, sheet })
        : byPosition
          ? resolveSpec({ ...entry.merged, keyColumns: [], matchRowsByPosition: true, endRow, sheet })
          : null,
      reason: keyColumns?.length
        ? undefined
        : byPosition
          ? 'rows matched by position — no row key found'
          : 'no keyColumns configured for this table',
    };
  });
}

/**
 * Checks `expect` against what layer 1 actually compared.
 *
 * A miss is an integrity error rather than a review item, deliberately: the
 * whole point is that coverage shrinking should stop a run, and a review item
 * would leave it green. Nothing about the comparison changes either way --
 * this reads the outcome, it does not steer it.
 */
function expectationErrors(spec: WorkbookSpec, outcomes: SheetOutcome[]): string[] {
  if (!spec.expect) return [];
  const out: string[] = [];

  for (const [name, want] of Object.entries(spec.expect)) {
    const on = outcomes.filter(
      (o) => o.status === 'compared' && canonSheet(o.sheet) === canonSheet(name),
    );

    if (typeof want === 'number') {
      if (on.length !== want) {
        out.push(
          `${name}: expected ${want} table(s) compared by name and key, found ${on.length}`,
        );
      }
      continue;
    }

    // Ranges are a set, not a sequence: a sheet whose tables are found in a
    // different order is not a finding, and one whose table moved is.
    const found = on.map((o) => o.range?.base ?? '(none)');
    const missing = want.filter((w) => !found.includes(w));
    const extra = found.filter((f) => !want.includes(f));
    if (!missing.length && !extra.length) continue;

    const bits = [`${name}: expected ${want.length} table(s), compared ${found.length}`];
    if (missing.length) bits.push(`not compared: ${missing.join(', ')}`);
    if (extra.length) bits.push(`unexpected: ${extra.join(', ')}`);
    out.push(bits.join(' — '));
  }
  return out;
}

/**
 * Names in `spec.sheets` that match no sheet in either file -- almost always a
 * typo, and worth stopping for, because a misspelled name means the sheet you
 * meant to configure is quietly being compared without your settings.
 *
 * A sheet marked `optional` is exempt: some reports of a type carry an extra
 * tab and some do not, and that is not a mistake. Marking it is what separates
 * the two, since the absence looks identical either way.
 */
function strayEntries(spec: WorkbookSpec, present: Set<string>): string[] {
  return Object.entries(spec.sheets ?? {})
    .filter(([k, v]) => !present.has(canonSheet(k)) && !v?.optional)
    .map(([k]) => k);
}

/**
 * Compares every sheet of a workbook against its baseline.
 *
 * The baseline drives the sheet list, so the golden file is the contract:
 * a sheet in the new output only is *noted* rather than compared -- there is
 * nothing to compare it against -- while a sheet that has disappeared is a
 * defect, because output the consumers expect is no longer produced.
 */
export async function verifyWorkbook(
  baselinePath: string,
  actualPath: string,
  spec: WorkbookSpec,
): Promise<WorkbookDiffResult> {
  return (await runWorkbook(baselinePath, actualPath, spec)).diff;
}

/** One compared table, kept so a cell-level ledger can be written afterwards. */
export interface ComparedTable {
  label: string;
  sheet: string;
  table: string;
  base: SheetModel;
  next: SheetModel;
  spec: ResolvedSpec;
  /** The comparison itself, so the ledger can mark root causes as such. */
  diff: DiffResult;
}

export interface WorkbookRun {
  diff: WorkbookDiffResult;
  /** Only the tables actually compared, in report order. */
  compared: ComparedTable[];
}

/**
 * The comparison, plus the models it ran on. `verifyWorkbook` discards the
 * models; a cell ledger needs them, and re-reading both files to get them
 * back would double the most expensive part of the run.
 */
export async function runWorkbook(
  baselinePath: string,
  actualPath: string,
  rawSpec: WorkbookSpec,
): Promise<WorkbookRun> {
  const spec = withMetadataRows(rawSpec);
  const ignored = new Set((spec.ignoreSheets ?? []).map(canonSheet));
  const compared: ComparedTable[] = [];

  // Cached because the readers call it once per sheet per file, and resolving
  // rebuilds the tolerance record and invariant list each time.
  const cache = new Map<string, ResolvedTable[]>();
  const tablesOn = (sheet: string): ResolvedTable[] => {
    const k = canonSheet(sheet);
    if (ignored.has(k)) return [];
    if (!cache.has(k)) cache.set(k, resolveTables(spec, sheet));
    return cache.get(k)!;
  };

  const requests = (sheet: string): TableRequest[] =>
    tablesOn(sheet)
      .filter((t) => t.spec)
      .map((t) => ({ table: t.table, key: t.key, spec: t.spec! }));

  const [base, next] = await Promise.all([
    workbookReaderFor(baselinePath).readWorkbook(baselinePath, requests),
    workbookReaderFor(actualPath).readWorkbook(actualPath, requests),
  ]);

  const baseByCanon = new Map(base.sheets.map((s) => [canonSheet(s), s]));
  const nextByCanon = new Map(next.sheets.map((s) => [canonSheet(s), s]));

  // Ignored tabs are out of scope entirely: they must not show up as added,
  // and a removed one must not fail the run. Positions still come from the
  // full lists, so a move is reported relative to the real workbook.
  const inScope = (s: string) => !ignored.has(canonSheet(s));

  const added = next.sheets.filter((s) => inScope(s) && !baseByCanon.has(canonSheet(s)));
  const removed = base.sheets.filter((s) => inScope(s) && !nextByCanon.has(canonSheet(s)));
  const moved: MovedSheet[] = base.sheets
    .filter(inScope)
    .map((s) => ({
      sheet: nextByCanon.get(canonSheet(s)) ?? s,
      from: base.sheets.indexOf(s) + 1,
      to: next.sheets.findIndex((n) => canonSheet(n) === canonSheet(s)) + 1,
    }))
    .filter((m) => m.to > 0 && m.from !== m.to);

  const errors: string[] = [];
  const outcomes: SheetOutcome[] = [];

  for (const sheet of base.sheets) {
    const k = canonSheet(sheet);

    if (ignored.has(k)) {
      outcomes.push({
        sheet, table: sheet, label: sheet,
        status: 'ignored', reason: 'excluded by ignoreSheets',
      });
      continue;
    }
    if (!nextByCanon.has(k)) {
      outcomes.push({
        sheet, table: sheet, label: sheet,
        status: 'removed', reason: 'present in the baseline only',
      });
      continue;
    }

    for (const t of tablesOn(sheet)) {
      const { table, label } = t;
      if (!t.spec) {
        outcomes.push({ sheet, table, label, status: 'skipped', reason: t.reason });
        continue;
      }

      const baseModel = base.models.get(t.key);
      const nextModel = next.models.get(t.key);
      if (!baseModel || !nextModel) {
        // Only reachable if a reader ignores the requests it was handed;
        // worth saying so rather than dereferencing undefined.
        errors.push(`${label}: reader returned no model despite a resolved spec`);
        continue;
      }

      const diff = compare(baseModel, nextModel, t.spec);
      compared.push({
        label, sheet, table, base: baseModel, next: nextModel, spec: t.spec, diff,
      });
      // The reason carries through for a compared table too, so a table
      // matched by position rather than by key says so in the report.
      outcomes.push({
        sheet, table, label, status: 'compared', diff, reason: t.reason,
        range: {
          base: tableRange(baseModel, t.spec.headerRow),
          next: tableRange(nextModel, t.spec.headerRow),
        },
      });
    }
  }

  // Every sheet in the new output that the baseline lacks, including ignored
  // ones -- listing those as 'ignored' says the tab was deliberately skipped,
  // which is more useful in the report than saying nothing about it.
  for (const sheet of next.sheets.filter((s) => !baseByCanon.has(canonSheet(s)))) {
    outcomes.push(
      ignored.has(canonSheet(sheet))
        ? { sheet, table: sheet, label: sheet, status: 'ignored', reason: 'excluded by ignoreSheets' }
        : {
            sheet, table: sheet, label: sheet,
            status: 'added',
            reason: 'new sheet — not in the baseline, so nothing to compare against',
          },
    );
  }

  const present = new Set([...base.sheets, ...next.sheets].map(canonSheet));
  for (const stray of strayEntries(spec, present)) {
    errors.push(`spec configures sheet "${stray}", which exists in neither file`);
  }

  errors.push(...expectationErrors(spec, outcomes));

  const skipped = outcomes.filter((o) => o.status === 'skipped');
  const defects =
    errors.length > 0 ||
    removed.length > 0 ||
    outcomes.some((o) => o.diff && !o.diff.ok);
  const strictFail = Boolean(spec.strictSheets) && (added.length > 0 || skipped.length > 0);
  const changed =
    added.length > 0 ||
    moved.length > 0 ||
    skipped.length > 0 ||
    outcomes.some((o) => o.diff?.reviewOnly);

  const ok = !defects && !strictFail;
  return {
    diff: {
      ok,
      reviewOnly: ok && changed,
      base: { source: baselinePath, sheets: base.sheets },
      next: { source: actualPath, sheets: next.sheets },
      sheetSchema: { added, removed, moved },
      sheets: outcomes,
      errors,
    },
    compared,
  };
}
