/**
 * Generates a realistic multi-sheet report, of the shape most real exports
 * actually have: every sheet carries a small "output info" block above the
 * data table, so there are two tables per sheet rather than one.
 *
 *   node --experimental-strip-types examples/generate-report.ts out.xlsx
 *   node --experimental-strip-types examples/generate-report.ts out.xlsx --next
 *
 * `--next` produces the following release: a column inserted mid-table, a new
 * sheet, rows added and dropped, a fresh timestamp in every info block, and
 * two genuine defects planted underneath the noise. `examples/cases.ts`
 * assembles narrower variants from the same parts.
 */
import { mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';
import ExcelJS from 'exceljs';

/** Info block occupies rows 1-6; the data table header sits on row 8. */
export const INFO_HEADER_ROW = 1;
export const DATA_HEADER_ROW = 8;

/** Every way the generated report can be varied, composed per case. */
export interface Variant {
  release?: string;
  generatedAt?: string;
  /** Insert a "Premium" column mid-table, shifting Rate and Annual Cost. */
  insertPremium?: boolean;
  /** A gross premium drifts: a value defect. */
  premiumDefect?: boolean;
  /** An agent's commission formula gains a stray uplift: a formula defect. */
  commissionDefect?: boolean;
  /** One claim closes and another opens: row population churn. */
  claimsChurn?: boolean;
  /** A new sheet appears, with no baseline to judge it against. */
  extraSheet?: boolean;
  /** Sheets left out entirely, as though the generator stopped producing them. */
  dropSheets?: string[];
  /** Write formulas with no cached result, as a broken generator would. */
  omitCachedResults?: boolean;
  /** A rate above 1.0 -- wrong, but wrong in *both* files, so only an
   *  invariant can catch it. */
  impossibleRate?: boolean;
}

/** The release under test: every kind of change at once. */
export const NEXT_RELEASE: Variant = {
  release: '4.3.0',
  generatedAt: '2026-08-16T09:31:44Z',
  insertPremium: true,
  premiumDefect: true,
  commissionDefect: true,
  claimsChurn: true,
  extraSheet: true,
};

const colLetter = (n: number): string => {
  let s = '';
  for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) {
    s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
  }
  return s;
};

/** Column letter for a header name, so formulas survive an inserted column. */
const at = (headers: string[], name: string): string =>
  colLetter(headers.indexOf(name) + 1);

type Row = (string | number)[];

/**
 * A formula column. Both halves matter: nothing here evaluates formulas, so a
 * generator that writes formula text without a cached result leaves value
 * comparison with nothing to check -- which `requireCachedValues` rejects.
 */
interface FormulaDef {
  text: (r: number, h: string[], row: Row) => string;
  value: (row: Row, h: string[]) => number;
}

interface SheetData {
  name: string;
  headers: string[];
  rows: Row[];
  formulas?: Record<string, FormulaDef>;
}

/** Reads a literal out of a row by header name. */
const val = (row: Row, h: string[], name: string): number => Number(row[h.indexOf(name)]);

function writeSheet(
  wb: ExcelJS.Workbook,
  sheet: SheetData,
  info: Record<string, string>,
  v: Variant,
) {
  const ws = wb.addWorksheet(sheet.name);

  // --- table 1: output info -------------------------------------------------
  ws.addRow(['Field', 'Value']);
  for (const [k, value] of Object.entries(info)) ws.addRow([k, value]);
  while (ws.rowCount < DATA_HEADER_ROW - 1) ws.addRow([]);

  // --- table 2: the data ----------------------------------------------------
  ws.addRow(sheet.headers);
  sheet.rows.forEach((row, i) => {
    const rowNum = DATA_HEADER_ROW + 1 + i;
    const values: any[] = [...row];
    for (const [column, def] of Object.entries(sheet.formulas ?? {})) {
      const idx = sheet.headers.indexOf(column);
      const formula = def.text(rowNum, sheet.headers, row);
      values[idx] = v.omitCachedResults
        ? { formula }
        : { formula, result: def.value(row, sheet.headers) };
    }
    ws.addRow(values);
  });

  ws.getRow(INFO_HEADER_ROW).font = { bold: true };
  ws.getRow(DATA_HEADER_ROW).font = { bold: true };
  ws.columns.forEach((c) => { c.width = 16; });
  return ws;
}

const POLICIES = [
  ['P-1001', 'Ivanov', 'Sofia', 120000, 0.021],
  ['P-1002', 'Petrov', 'Plovdiv', 85000, 0.018],
  ['P-1003', 'Georgiev', 'Varna', 240000, 0.025],
  ['P-1004', 'Dimitrov', 'Burgas', 60000, 0.017],
  ['P-1005', 'Nikolov', 'Ruse', 310000, 0.029],
] as const;

const AGENTS = [
  ['A-01', 'Stoyanova', 'Sofia', 1_240_000, 0.035],
  ['A-02', 'Marinov', 'Plovdiv', 860_000, 0.032],
  ['A-03', 'Todorov', 'Varna', 1_510_000, 0.038],
] as const;

const REGIONS = ['Sofia', 'Plovdiv', 'Varna', 'Burgas', 'Ruse'] as const;

export function buildSheets(v: Variant): SheetData[] {
  const policyHeaders = v.insertPremium
    ? ['PolicyId', 'Holder', 'Region', 'Sum Insured', 'Premium', 'Rate', 'Annual Cost']
    : ['PolicyId', 'Holder', 'Region', 'Sum Insured', 'Rate', 'Annual Cost'];

  const rateOf = (id: string, rate: number) =>
    v.impossibleRate && id === 'P-1002' ? 1.4 : rate;

  const policies: SheetData = {
    name: 'Policies',
    headers: policyHeaders,
    rows: POLICIES.map(([id, holder, region, sum, rate]) =>
      v.insertPremium
        ? [id, holder, region, sum, 'Standard', rateOf(id, rate), 0]
        : [id, holder, region, sum, rateOf(id, rate), 0],
    ),
    formulas: {
      'Annual Cost': {
        text: (r, h) => `${at(h, 'Sum Insured')}${r}*${at(h, 'Rate')}${r}`,
        value: (row, h) => val(row, h, 'Sum Insured') * val(row, h, 'Rate'),
      },
    },
  };

  const premiums: SheetData = {
    name: 'Premiums',
    headers: ['PolicyId', 'Period', 'Gross', 'Tax', 'Net'],
    rows: POLICIES.flatMap(([id, , , sum, rate]) =>
      ['2026-07', '2026-08'].map((period) => {
        const drifted = v.premiumDefect && id === 'P-1003' && period === '2026-08';
        const gross = Math.round(sum * rate * (drifted ? 1.15 : 1));
        return [id, period, gross, 0, 0];
      }),
    ),
    formulas: {
      Tax: {
        text: (r, h) => `${at(h, 'Gross')}${r}*0.2`,
        value: (row, h) => val(row, h, 'Gross') * 0.2,
      },
      Net: {
        text: (r, h) => `${at(h, 'Gross')}${r}+${at(h, 'Tax')}${r}`,
        value: (row, h) => val(row, h, 'Gross') * 1.2,
      },
    },
  };

  const claims: SheetData = {
    name: 'Claims',
    headers: ['ClaimId', 'PolicyId', 'Gross Claim', 'Excess', 'Net Claim'],
    rows: [
      ['C-5001', 'P-1001', 4200, 500, 0],
      ['C-5002', 'P-1003', 18700, 1000, 0],
      ...(v.claimsChurn ? [] : [['C-5003', 'P-1004', 3100, 250, 0] as Row]),
      ['C-5004', 'P-1005', 9450, 750, 0],
      ...(v.claimsChurn ? [['C-5005', 'P-1002', 2600, 300, 0] as Row] : []),
    ],
    formulas: {
      'Net Claim': {
        text: (r, h) => `${at(h, 'Gross Claim')}${r}-${at(h, 'Excess')}${r}`,
        value: (row, h) => val(row, h, 'Gross Claim') - val(row, h, 'Excess'),
      },
    },
  };

  const commissions: SheetData = {
    name: 'Commissions',
    headers: ['AgentId', 'Agent', 'Region', 'Volume', 'Rate', 'Commission'],
    rows: AGENTS.map(([id, agent, region, volume, rate]) => [id, agent, region, volume, rate, 0]),
    formulas: {
      Commission: {
        text: (r, h, row) => {
          const base = `${at(h, 'Volume')}${r}*${at(h, 'Rate')}${r}`;
          return v.commissionDefect && row[0] === 'A-02' ? `${base}*1.1` : base;
        },
        value: (row, h) => {
          const base = val(row, h, 'Volume') * val(row, h, 'Rate');
          return v.commissionDefect && row[0] === 'A-02' ? base * 1.1 : base;
        },
      },
    },
  };

  const regions: SheetData = {
    name: 'Regions',
    headers: ['Region', 'Premiums Written', 'Claims Paid', 'Loss Ratio'],
    rows: REGIONS.map((region, i) => [region, 400000 + i * 25000, 90000 + i * 12000, 0]),
    formulas: {
      'Loss Ratio': {
        text: (r, h) => `${at(h, 'Claims Paid')}${r}/${at(h, 'Premiums Written')}${r}`,
        value: (row, h) => val(row, h, 'Claims Paid') / val(row, h, 'Premiums Written'),
      },
    },
  };

  const sheets = [policies, premiums, claims, commissions, regions];

  if (v.extraSheet) {
    sheets.push({
      name: 'Premium Detail',
      headers: ['PolicyId', 'Band', 'Loading', 'Adjusted'],
      rows: POLICIES.map(([id, , , sum]) => [id, 'Standard', 0.05, sum]),
      formulas: {
        Adjusted: {
          text: (r, h) => `${at(h, 'Loading')}${r}*100`,
          value: (row, h) => val(row, h, 'Loading') * 100,
        },
      },
    });
  }

  const dropped = new Set(v.dropSheets ?? []);
  return sheets.filter((s) => !dropped.has(s.name));
}

export async function generate(path: string, v: Variant = {}): Promise<string> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'POLADMIN Export Service';
  wb.created = new Date('2026-08-16T09:00:00Z');

  for (const sheet of buildSheets(v)) {
    writeSheet(wb, sheet, {
      'Report Name': 'Monthly Policy Export',
      Creator: 'POLADMIN Export Service',
      'Source System': 'POLADMIN',
      Release: v.release ?? '4.2.0',
      // Changes on every single run: the classic false-positive source.
      'Generated At': v.generatedAt ?? '2026-07-15T08:02:11Z',
    }, v);
  }

  await mkdir(dirname(path), { recursive: true });
  await wb.xlsx.writeFile(path);
  return path;
}

const target = process.argv[2];
if (target) {
  const isNext = process.argv.includes('--next');
  generate(target, isNext ? NEXT_RELEASE : {}).then((p) =>
    console.log(`wrote ${p}${isNext ? ' (next release)' : ''}`),
  );
}
