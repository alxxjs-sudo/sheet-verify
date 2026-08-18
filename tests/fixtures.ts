import { mkdir, writeFile, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

export const DIR = join(import.meta.dirname, '.fixtures');

export interface PolicyRow {
  id: string; holder: string; region: string; sumInsured: number; rate: number;
}

export const POLICIES: PolicyRow[] = [
  { id: 'P-1001', holder: 'Ivanov',   region: 'Sofia',   sumInsured: 120000, rate: 0.021 },
  { id: 'P-1002', holder: 'Petrov',   region: 'Plovdiv', sumInsured: 85000,  rate: 0.018 },
  { id: 'P-1003', holder: 'Georgiev', region: 'Varna',   sumInsured: 240000, rate: 0.025 },
  { id: 'P-1004', holder: 'Dimitrov', region: 'Burgas',  sumInsured: 60000,  rate: 0.017 },
  { id: 'P-1005', holder: 'Nikolov',  region: 'Ruse',    sumInsured: 310000, rate: 0.029 },
];

export interface BuildOptions {
  /** Insert a "Premium" column at position 5, shifting later columns right. */
  insertPremium?: boolean;
  /** Policy ids whose Sum Insured is altered, id -> new value. */
  valueDrift?: Record<string, number>;
  /** Policy ids whose commission rate changes from 0.1. */
  rateDrift?: Record<string, number>;
  /** Write Sum Insured as text rather than a number, for these ids. */
  textDrift?: string[];
  /** Omit cached results, mimicking a generator that writes formulas only. */
  omitCachedResults?: boolean;
  /** Policy ids to leave out entirely. */
  dropRows?: string[];
  /** Extra rows appended. */
  extraRows?: PolicyRow[];
}

/**
 * Builds a policy report. Column layout:
 *   PolicyId | Holder | Region | Sum Insured | [Premium] | Rate | Annual Cost | Commission
 * Annual Cost and Commission are formulas, so an inserted column shifts both
 * their position and the references inside them.
 */
export async function buildWorkbook(path: string, opts: BuildOptions = {}): Promise<string> {
  await mkdir(DIR, { recursive: true });
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Policies');

  const headers = ['PolicyId', 'Holder', 'Region', 'Sum Insured'];
  if (opts.insertPremium) headers.push('Premium');
  headers.push('Rate', 'Annual Cost', 'Commission');
  ws.addRow(headers);

  // Column letters shift when Premium is present.
  const P = opts.insertPremium;
  const cSum = 'D';
  const cRate = P ? 'F' : 'E';
  const cAnnual = P ? 'G' : 'F';

  const rows = POLICIES.filter((p) => !opts.dropRows?.includes(p.id)).concat(opts.extraRows ?? []);

  rows.forEach((p, i) => {
    const n = i + 2;
    const sum = opts.valueDrift?.[p.id] ?? p.sumInsured;
    const commRate = opts.rateDrift?.[p.id] ?? 0.1;
    const annual = sum * p.rate;

    const sumCell = opts.textDrift?.includes(p.id) ? String(sum) : sum;
    const mk = (formula: string, result: number) =>
      opts.omitCachedResults ? { formula } : { formula, result };

    const values: any[] = [p.id, p.holder, p.region, sumCell];
    if (P) values.push('Standard');
    values.push(
      p.rate,
      mk(`${cSum}${n}*${cRate}${n}`, annual),
      mk(`${cAnnual}${n}*${commRate}`, annual * commRate),
    );
    ws.addRow(values);
  });

  const full = join(DIR, path);
  await wb.xlsx.writeFile(full);
  return full;
}

export const PERIODS = ['2026-07', '2026-08'];

/**
 * A sheet carrying two tables: an "output info" block on rows 1-5 and a data
 * table whose header sits on row 7. The shape almost every real export has,
 * and the one a single headerRow cannot describe.
 */
export async function buildTwoTableSheet(
  path: string,
  opts: { generatedAt?: string; release?: string; drift?: number; extraInfo?: boolean } = {},
): Promise<string> {
  await mkdir(DIR, { recursive: true });
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Policies');

  ws.addRow(['Field', 'Value']);
  ws.addRow(['Report Name', 'Monthly Policy Export']);
  ws.addRow(['Release', opts.release ?? '4.2.0']);
  ws.addRow(['Generated At', opts.generatedAt ?? '2026-07-15T08:02:11Z']);
  if (opts.extraInfo) ws.addRow(['Source System', 'POLADMIN']);
  while (ws.rowCount < 6) ws.addRow([]);

  ws.addRow(['PolicyId', 'Holder', 'Sum Insured', 'Rate', 'Annual Cost']);
  POLICIES.forEach((p, i) => {
    const n = 8 + i;
    const sum = i === 0 && opts.drift ? opts.drift : p.sumInsured;
    ws.addRow([p.id, p.holder, sum, p.rate, { formula: `C${n}*D${n}`, result: sum * p.rate }]);
  });

  const full = join(DIR, path);
  await wb.xlsx.writeFile(full);
  return full;
}

export interface MultiSheetOptions {
  /** Sheet names, in workbook order. Default Policies, Premiums, Regions. */
  sheets?: string[];
  /** Insert a "Premium" column on the Policies sheet, shifting its formulas. */
  insertPremium?: boolean;
  /** Sum Insured drift on Policies, by policy id. */
  valueDrift?: Record<string, number>;
  /** Amount drift on Premiums, by "id|period". */
  premiumDrift?: Record<string, number>;
}

/**
 * Builds a workbook whose sheets deliberately have little in common -- one
 * keyed by a single column with formulas that shift, one by a composite key,
 * one a small lookup table. That is the shape the workbook layer has to cope
 * with: per-sheet specs, not one spec applied everywhere.
 */
export async function buildMultiSheet(
  path: string,
  opts: MultiSheetOptions = {},
): Promise<string> {
  await mkdir(DIR, { recursive: true });
  const wb = new ExcelJS.Workbook();

  for (const name of opts.sheets ?? ['Policies', 'Premiums', 'Regions']) {
    const ws = wb.addWorksheet(name);

    if (name === 'Policies') {
      const headers = ['PolicyId', 'Holder', 'Region', 'Sum Insured'];
      if (opts.insertPremium) headers.push('Premium');
      headers.push('Rate', 'Annual Cost');
      ws.addRow(headers);

      const cRate = opts.insertPremium ? 'F' : 'E';
      POLICIES.forEach((p, i) => {
        const n = i + 2;
        const sum = opts.valueDrift?.[p.id] ?? p.sumInsured;
        const values: any[] = [p.id, p.holder, p.region, sum];
        if (opts.insertPremium) values.push('Standard');
        values.push(p.rate, { formula: `D${n}*${cRate}${n}`, result: sum * p.rate });
        ws.addRow(values);
      });
    } else if (name === 'Premiums') {
      ws.addRow(['PolicyId', 'Period', 'Amount', 'Tax']);
      let n = 1;
      for (const p of POLICIES) {
        for (const period of PERIODS) {
          n++;
          const amount = opts.premiumDrift?.[`${p.id}|${period}`] ?? p.sumInsured * p.rate;
          ws.addRow([p.id, period, amount, { formula: `C${n}*0.2`, result: amount * 0.2 }]);
        }
      }
    } else if (name === 'Regions') {
      ws.addRow(['Region', 'Office', 'Load']);
      for (const p of POLICIES) ws.addRow([p.region, `${p.region} Branch`, p.rate * 10]);
    } else {
      // Generic filler, used as the unconfigured / newly added sheet.
      ws.addRow(['Id', 'Text']);
      ws.addRow(['N-1', 'generated automatically']);
    }
  }

  const full = join(DIR, path);
  await wb.xlsx.writeFile(full);
  return full;
}

export interface SweepBuildOptions {
  /** Sum Insured drift, by policy id. Lands in a column layer 1 compares. */
  sumDrift?: Record<string, number>;
  /** Run stamp value, meant to be excluded with ignoreColumns. */
  stamp?: string;
  /** Value on the Notes sheet, which has no usable row key. */
  note?: string;
  /** Extends the Annual Cost formula. */
  formulaDrift?: boolean;
  /**
   * Store cached results for the formulas. A report straight from the
   * generator has none; one a person opened in Excel and saved has them.
   */
  cacheResults?: boolean;
  /** Inserts a Premium column at position 3, shifting everything after it. */
  insertColumn?: boolean;
  /** Appends a sheet the other side does not have. */
  extraSheet?: boolean;
}

/**
 * A workbook built for the sweep: one sheet layer 1 can key, and one it cannot.
 *
 * Notes repeats its first column deliberately, so no key detects and layer 1
 * reports the sheet as not compared. A change planted there is invisible to
 * layer 1 by construction -- which is the case the sweep exists to catch.
 *
 * No formula carries a cached result, matching the real reports.
 */
export async function buildSweepWorkbook(
  path: string,
  opts: SweepBuildOptions = {},
): Promise<string> {
  await mkdir(DIR, { recursive: true });
  const wb = new ExcelJS.Workbook();

  const ws = wb.addWorksheet('Policies');
  const headers = ['PolicyId', 'Holder'];
  if (opts.insertColumn) headers.push('Premium');
  headers.push('Sum Insured', 'Rate', 'Annual Cost', 'Run Stamp');
  ws.addRow(headers);

  const cSum = opts.insertColumn ? 'D' : 'C';
  const cRate = opts.insertColumn ? 'E' : 'D';

  POLICIES.forEach((p, i) => {
    const n = i + 2;
    const sum = opts.sumDrift?.[p.id] ?? p.sumInsured;
    const values: any[] = [p.id, p.holder];
    if (opts.insertColumn) values.push('Standard');
    const formula = `${cSum}${n}*${cRate}${n}${opts.formulaDrift ? '*1.05' : ''}`;
    const result = sum * p.rate * (opts.formulaDrift ? 1.05 : 1);
    values.push(
      sum,
      p.rate,
      opts.cacheResults ? { formula, result } : { formula },
      opts.stamp ?? '2026-08-17T09:00:00Z',
    );
    ws.addRow(values);
  });

  const notes = wb.addWorksheet('Notes');
  notes.addRow(['Appendix']);
  notes.addRow(['Note', 'Value']);
  notes.addRow(['Coverage', opts.note ?? 'Standard']);
  notes.addRow(['Coverage', 'Standard']);
  notes.addRow(['Coverage', 'Standard']);

  if (opts.extraSheet) {
    const extra = wb.addWorksheet('Addendum');
    extra.addRow(['Item', 'Amount']);
    extra.addRow(['Reinstatement', 5000]);
  }

  const full = join(DIR, path);
  await wb.xlsx.writeFile(full);
  return full;
}

/**
 * Rewrites a column's individual formulas into an Excel shared formula group:
 * the first cell holds the text, the rest reference it by index. ExcelJS
 * cannot author these, but real Excel files are full of them.
 */
export async function makeSharedFormulas(
  path: string, column: string, firstRow: number, lastRow: number, si: number,
): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(path));
  const file = zip.file('xl/worksheets/sheet1.xml');
  if (!file) throw new Error('sheet1.xml not found');
  let xml = await file.async('string');

  const cellRe = (row: number) =>
    new RegExp(`(<c r="${column}${row}"[^>]*>)<f>([^<]*)</f>`);

  const master = cellRe(firstRow).exec(xml);
  if (!master) throw new Error(`no formula at ${column}${firstRow}`);
  const masterFormula = master[2]!;

  xml = xml.replace(
    cellRe(firstRow),
    `$1<f t="shared" ref="${column}${firstRow}:${column}${lastRow}" si="${si}">${masterFormula}</f>`,
  );
  for (let r = firstRow + 1; r <= lastRow; r++) {
    xml = xml.replace(cellRe(r), `$1<f t="shared" si="${si}"/>`);
  }

  zip.file('xl/worksheets/sheet1.xml', xml);
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }));
}

export async function writeCsv(
  name: string,
  opts: { delimiter?: string; bom?: boolean; crlf?: boolean; thousands?: boolean; drift?: boolean } = {},
): Promise<string> {
  await mkdir(DIR, { recursive: true });
  const d = opts.delimiter ?? ',';
  const eol = opts.crlf ? '\r\n' : '\n';
  const fmt = (n: number) =>
    opts.thousands ? `"${n.toLocaleString('en-US')}"` : String(n);

  const lines = [['PolicyId', 'Holder', 'Region', 'Sum Insured', 'Rate'].join(d)];
  for (const p of POLICIES) {
    const sum = opts.drift && p.id === 'P-1003' ? 249000 : p.sumInsured;
    lines.push([p.id, p.holder, p.region, fmt(sum), String(p.rate)].join(d));
  }
  const text = (opts.bom ? '﻿' : '') + lines.join(eol) + eol;
  const full = join(DIR, name);
  await writeFile(full, text, 'utf8');
  return full;
}
