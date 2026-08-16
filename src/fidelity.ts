/**
 * Template round-trip fidelity gate.
 *
 * Editing a template means reading it and writing it back, and every Excel
 * library silently drops parts of the format when it does. Run this against
 * real templates before trusting any of them: it opens a file, saves it with
 * no edits at all, and reports what did not survive.
 */
import { readFile, writeFile, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

/** Parts Excel regenerates on open. Losing these costs nothing. */
const BENIGN = [/^xl\/calcChain\.xml$/, /\/$/];

/** Parts whose loss silently destroys template content. */
const CRITICAL: { pattern: RegExp; what: string }[] = [
  { pattern: /^xl\/drawings\//, what: 'drawings (shapes, text boxes)' },
  { pattern: /^xl\/charts\//, what: 'charts' },
  { pattern: /^xl\/media\//, what: 'embedded images' },
  { pattern: /^xl\/pivotTables?\//, what: 'pivot tables' },
  { pattern: /^xl\/pivotCache\//, what: 'pivot caches' },
  { pattern: /^xl\/tables\//, what: 'named tables' },
  { pattern: /^xl\/threadedComments\//, what: 'threaded comments' },
  { pattern: /^xl\/ctrlProps\//, what: 'form controls' },
  { pattern: /^xl\/embeddings\//, what: 'embedded objects' },
  { pattern: /^customXml\//, what: 'custom XML (often used by document generators)' },
];

/** In-sheet features that vanish without removing a whole part. */
const FEATURES: { tag: string; what: string }[] = [
  { tag: 'conditionalFormatting', what: 'conditional formatting rules' },
  { tag: 'dataValidation', what: 'data validation (dropdowns)' },
  { tag: 'mergeCell', what: 'merged cells' },
  { tag: 'autoFilter', what: 'auto-filters' },
  { tag: 'sheetProtection', what: 'sheet protection' },
  { tag: 'pane', what: 'frozen panes' },
  { tag: 'hyperlink', what: 'hyperlinks' },
  { tag: 'dataBar', what: 'data bars' },
];

export interface FidelityFinding {
  severity: 'critical' | 'warning' | 'info';
  detail: string;
}

export interface FidelityReport {
  file: string;
  ok: boolean;
  findings: FidelityFinding[];
}

async function partsOf(buf: Buffer): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(buf);
  const out = new Map<string, string>();
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    out.set(name, name.endsWith('.xml') || name.endsWith('.rels') ? await entry.async('string') : '');
  }
  return out;
}

function countTag(parts: Map<string, string>, tag: string): number {
  let n = 0;
  for (const [name, xml] of parts) {
    if (!name.startsWith('xl/worksheets/') || !xml) continue;
    n += (xml.match(new RegExp(`<${tag}[\\s/>]`, 'g')) ?? []).length;
  }
  return n;
}

/**
 * Compares the parts and in-sheet features of a workbook before and after a
 * round-trip. Separated from the I/O so the rules can be tested directly.
 */
export function classify(
  before: Map<string, string>,
  after: Map<string, string>,
): FidelityFinding[] {
  const findings: FidelityFinding[] = [];

  for (const part of before.keys()) {
    if (after.has(part) || BENIGN.some((re) => re.test(part))) continue;
    const crit = CRITICAL.find((c) => c.pattern.test(part));
    findings.push(
      crit
        ? { severity: 'critical', detail: `lost ${crit.what} — ${part}` }
        : { severity: 'warning', detail: `lost part ${part}` },
    );
  }

  for (const f of FEATURES) {
    const b = countTag(before, f.tag);
    const a = countTag(after, f.tag);
    if (b > a) {
      findings.push({
        severity: 'critical',
        detail: `${f.what}: ${b} before, ${a} after — ${b - a} lost`,
      });
    }
  }

  if (before.has('xl/calcChain.xml') && !after.has('xl/calcChain.xml')) {
    findings.push({ severity: 'info', detail: 'calcChain.xml dropped — harmless, Excel rebuilds it' });
  }

  return findings;
}

/** Opens and re-saves the file with no edits, then reports what was lost. */
export async function checkFidelity(path: string): Promise<FidelityReport> {
  const dir = await mkdtemp(join(tmpdir(), 'sheet-fidelity-'));
  const out = join(dir, 'roundtrip.xlsx');
  const findings: FidelityFinding[] = [];

  try {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(path);
      await wb.xlsx.writeFile(out);
    } catch (e) {
      // The worst outcome of all: the library cannot handle this template at
      // all. Report it rather than throwing, so a gate over many files still
      // completes and shows every problem at once.
      return {
        file: path, ok: false,
        findings: [{
          severity: 'critical',
          detail: `the library cannot round-trip this file: ${(e as Error).message}`,
        }],
      };
    }

    const before = await partsOf(await readFile(path));
    const after = await partsOf(await readFile(out));
    findings.push(...classify(before, after));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }

  return { file: path, ok: !findings.some((f) => f.severity === 'critical'), findings };
}

export function formatFidelity(reports: FidelityReport[]): string {
  const lines: string[] = [];
  for (const r of reports) {
    const mark = r.ok ? '✓' : '✗';
    lines.push(`${mark} ${r.file}`);
    if (!r.findings.length) lines.push('    round-tripped with no detectable loss');
    for (const f of r.findings) {
      const icon = f.severity === 'critical' ? '  ✗ ' : f.severity === 'warning' ? '  ! ' : '  · ';
      lines.push(`  ${icon}${f.detail}`);
    }
    lines.push('');
  }
  const bad = reports.filter((r) => !r.ok).length;
  lines.push(
    bad === 0
      ? 'All templates survived the round-trip. Editing them with this library is safe.'
      : `${bad} of ${reports.length} template(s) lose content on save. Do not rewrite these — ` +
        `drive them by writing only into a data sheet their own formulas read from.`,
  );
  return lines.join('\n');
}
