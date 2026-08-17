/**
 * The example cases. Each one is a scenario worth being able to recognise in
 * a real run, paired with what it should produce -- so the folder doubles as
 * a reference for reading a diff you have not seen before.
 */
import type { CaseOptions } from '../dist/index.js';
import { invariants as inv } from '../dist/index.js';
import { NEXT_RELEASE, type Variant } from './generate-report.ts';

/** Every sheet carries the same info block, so it lives in one constant. */
const info = {
  headerRow: 1,
  keyColumns: ['Field'],
  // Rewritten on every run: a difference here is never a defect.
  ignoreRows: ['Generated At'],
};

/** The spec the five-sheet report is compared under. */
export const SPEC: CaseOptions = {
  defaults: { headerRow: 8 },
  sheets: {
    Policies: { tables: { Info: info, Detail: { keyColumns: ['PolicyId'] } } },
    Premiums: { tables: { Info: info, Detail: { keyColumns: ['PolicyId', 'Period'] } } },
    Claims: { tables: { Info: info, Detail: { keyColumns: ['ClaimId'] } } },
    Commissions: { tables: { Info: info, Detail: { keyColumns: ['AgentId'] } } },
    Regions: { tables: { Info: info, Detail: { keyColumns: ['Region'] } } },
  },
};

export interface ExampleCase {
  name: string;
  /** What the case demonstrates, printed by the runner. */
  about: string;
  /** How the golden output is generated. */
  golden: Variant;
  /** How the report under test is generated. */
  actual: Variant;
  spec?: CaseOptions;
  /** Whether the comparison is expected to pass, for the runner's summary. */
  expect: 'pass' | 'fail';
}

export const CASES: ExampleCase[] = [
  {
    name: 'monthly-policy-export',
    about: 'the full picture: schema drift, row churn and a new sheet, with two real defects underneath',
    golden: {},
    actual: NEXT_RELEASE,
    expect: 'fail',
  },
  {
    name: 'clean-release',
    about: 'nothing changed — the case that should be the common one',
    golden: {},
    actual: {},
    expect: 'pass',
  },
  {
    name: 'schema-drift-only',
    about: 'a Premium column inserted mid-table and nothing else: reported once, not as churn across every row',
    golden: {},
    actual: { insertPremium: true },
    expect: 'pass',
  },
  {
    name: 'sheet-removed',
    about: 'the Regions sheet stops being produced — a defect with no cells to point at',
    golden: {},
    actual: { dropSheets: ['Regions'] },
    expect: 'fail',
  },
  {
    name: 'uncached-formulas',
    about: 'the generator writes formulas with no cached result, so value comparison has nothing to check',
    golden: {},
    actual: { omitCachedResults: true },
    expect: 'fail',
  },
  {
    name: 'invariant-catch',
    about: 'a rate of 1.4 in BOTH files: the comparison passes, and only an invariant catches it',
    golden: { impossibleRate: true },
    actual: { impossibleRate: true },
    spec: {
      ...SPEC,
      defaults: {
        headerRow: 8,
        invariants: [inv.noErrorValues(), inv.inRange('Rate', 0, 1)],
      },
    },
    expect: 'fail',
  },
];
