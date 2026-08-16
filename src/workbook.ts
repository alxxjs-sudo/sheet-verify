import type {
  MovedSheet, ResolvedSpec, SheetOutcome, WorkbookDiffResult,
  WorkbookReader, WorkbookSheetSpec, WorkbookSpec,
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
      `sheet-verify: "${path}" is read by a single-sheet reader and has no worksheets. ` +
      'Use verifySheet()/toMatchSheetBaseline for CSV.',
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
  return {
    ...d,
    ...s,
    tolerance: { ...asRecord(d.tolerance), ...asRecord(s.tolerance) },
    ignoreColumns: [...(d.ignoreColumns ?? []), ...(s.ignoreColumns ?? [])],
    invariants: [...(d.invariants ?? []), ...(s.invariants ?? [])],
  };
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
  const ignored = new Set((spec.ignoreSheets ?? []).map(canonSheet));

  // Cached because the readers call it once per sheet per file, and resolving
  // rebuilds the tolerance record and invariant list each time.
  const cache = new Map<string, ResolvedSpec | null>();
  const specFor = (sheet: string): ResolvedSpec | null => {
    const k = canonSheet(sheet);
    if (ignored.has(k)) return null;
    if (!cache.has(k)) cache.set(k, resolveSheetSpec(spec, sheet));
    return cache.get(k)!;
  };

  const [base, next] = await Promise.all([
    workbookReaderFor(baselinePath).readWorkbook(baselinePath, specFor),
    workbookReaderFor(actualPath).readWorkbook(actualPath, specFor),
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
      outcomes.push({ sheet, status: 'ignored', reason: 'excluded by ignoreSheets' });
      continue;
    }
    const nextName = nextByCanon.get(k);
    if (nextName === undefined) {
      outcomes.push({ sheet, status: 'removed', reason: 'present in the baseline only' });
      continue;
    }
    const resolved = specFor(sheet);
    if (!resolved) {
      outcomes.push({
        sheet,
        status: 'skipped',
        reason: 'no keyColumns configured for this sheet',
      });
      continue;
    }

    const baseModel = base.models.get(sheet);
    const nextModel = next.models.get(nextName);
    if (!baseModel || !nextModel) {
      // Only reachable if a reader ignores specFor; worth saying so rather
      // than dereferencing undefined.
      errors.push(`sheet "${sheet}": reader returned no model despite a resolved spec`);
      continue;
    }

    outcomes.push({ sheet, status: 'compared', diff: compare(baseModel, nextModel, resolved) });
  }

  // Every sheet in the new output that the baseline lacks, including ignored
  // ones -- listing those as 'ignored' says the tab was deliberately skipped,
  // which is more useful in the report than saying nothing about it.
  for (const sheet of next.sheets.filter((s) => !baseByCanon.has(canonSheet(s)))) {
    outcomes.push(
      ignored.has(canonSheet(sheet))
        ? { sheet, status: 'ignored', reason: 'excluded by ignoreSheets' }
        : {
            sheet,
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
    ok,
    reviewOnly: ok && changed,
    base: { source: baselinePath, sheets: base.sheets },
    next: { source: actualPath, sheets: next.sheets },
    sheetSchema: { added, removed, moved },
    sheets: outcomes,
    errors,
  };
}
