import type {
  DiffResult, MovedSheet, ResolvedSpec, SheetModel, SheetOutcome, TableRequest, TableSpec,
  WorkbookDiffResult, WorkbookReader, WorkbookSheetSpec, WorkbookSpec,
} from './types.js';
import { resolveSpec } from './model.js';
import { compare } from './compare.js';
import { readerFor } from './verify.js';

/** Excel forbids two sheets whose names differ only in case, so this is safe. */
const canonSheet = (s: string): string => s.trim().toLowerCase();

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

/** Resolves one sheet's spec, or null when it has no key and cannot be compared. */
export function resolveSheetSpec(spec: WorkbookSpec, sheet: string): ResolvedSpec | null {
  const merged = mergeSheetSpec(spec.defaults, entryFor(spec.sheets, sheet));
  if (!merged.keyColumns?.length) return null;
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
  const merged = mergeSheetSpec(spec.defaults, entryFor(spec.sheets, sheet));
  const declared = Object.entries(merged.tables ?? {});

  if (!declared.length) {
    return [{
      table: sheet,
      key: tableKey(sheet, sheet),
      label: sheet,
      spec: resolveSheetSpec(spec, sheet),
      reason: merged.keyColumns?.length ? undefined : 'no keyColumns configured for this sheet',
    }];
  }

  // `tables` itself must not be inherited down into each table.
  const { tables: _drop, ...sheetLevel } = merged;

  const ordered = declared
    .map(([table, t]) => ({ table, merged: mergeSheetSpec(sheetLevel, t) }))
    .sort((a, b) => (a.merged.headerRow ?? 1) - (b.merged.headerRow ?? 1));

  return ordered.map((entry, i) => {
    const next = ordered[i + 1];
    // An explicit endRow wins; otherwise stop just above the next table.
    const endRow = entry.merged.endRow ?? (next ? (next.merged.headerRow ?? 1) - 1 : 0);
    const keyColumns = entry.merged.keyColumns;

    return {
      table: entry.table,
      key: tableKey(sheet, entry.table),
      label: labelFor(sheet, entry.table),
      spec: keyColumns?.length
        ? resolveSpec({ ...entry.merged, keyColumns, endRow, sheet })
        : null,
      reason: keyColumns?.length ? undefined : 'no keyColumns configured for this table',
    };
  });
}

/** Names in `spec.sheets` that match no sheet in either file -- almost always a typo. */
function strayEntries(spec: WorkbookSpec, present: Set<string>): string[] {
  return Object.keys(spec.sheets ?? {}).filter((k) => !present.has(canonSheet(k)));
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
  spec: WorkbookSpec,
): Promise<WorkbookRun> {
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
      outcomes.push({ sheet, table, label, status: 'compared', diff });
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
