import { test, expect } from '@playwright/test';
import { mkdir, readFile, readdir, rm, writeFile, cp } from 'node:fs/promises';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import ExcelJS from 'exceljs';
import { buildMultiSheet, DIR } from './fixtures.js';

/**
 * The CLI over a tree of report types, with configuration inherited down it.
 * Runs the built CLI as a subprocess, so what is asserted is what a user gets.
 */

const ROOT = join(DIR, 'tree');
const cli = (...args: string[]) => {
  const r = spawnSync('node', ['dist/cli.js', ...args], { encoding: 'utf8' });
  return { out: `${r.stdout ?? ''}${r.stderr ?? ''}`, code: r.status };
};

const writeJson = (path: string, value: unknown) =>
  writeFile(path, JSON.stringify(value, null, 2), 'utf8');

/**
 *   tree/
 *     meta.json                          global
 *     reports/
 *       global_standard_cat/
 *         meta.json                      per type
 *         case_001/  case_002/           case_002 has its own case.json
 *       srq/
 *         case_001/
 */
async function buildTree(): Promise<void> {
  await rm(ROOT, { recursive: true, force: true });

  const gsc = join(ROOT, 'reports', 'global_standard_cat');
  const srq = join(ROOT, 'reports', 'srq');
  for (const dir of [join(gsc, 'case_001'), join(gsc, 'case_002'), join(srq, 'case_001')]) {
    await mkdir(dir, { recursive: true });
    const golden = await buildMultiSheet(`tree-${dir.replace(/[^a-z0-9]/gi, '')}-g.xlsx`);
    const actual = await buildMultiSheet(`tree-${dir.replace(/[^a-z0-9]/gi, '')}-a.xlsx`, {
      premiumDrift: { 'P-1003|2026-08': 9999 },
    });
    await cp(golden, join(dir, 'golden.xlsx'));
    await cp(actual, join(dir, 'actual.xlsx'));
  }

  await writeJson(join(ROOT, 'meta.json'), { ignoreSheets: ['Regions'] });
  // A tolerance wide enough to absorb the drift and the formula downstream of
  // it, so the whole type passes...
  await writeJson(join(gsc, 'meta.json'), {
    sheets: { Premiums: { tolerance: { '*': 100000 } } },
  });
  // ...and one case that insists on exactness for itself alone.
  await writeJson(join(gsc, 'case_002', 'case.json'), {
    sheets: { Premiums: { tolerance: { '*': 0 } } },
  });
}

// The tree is shared state on disk, so these must not run against each other.
test.describe.configure({ mode: 'serial' });

test.describe('a tree of report types', () => {
  test.beforeAll(buildTree);

  test('finds cases at any depth and names them by path', async () => {
    const { out } = cli(ROOT);
    expect(out).toContain('reports/global_standard_cat/case_001');
    expect(out).toContain('reports/global_standard_cat/case_002');
    expect(out).toContain('reports/srq/case_001');
    expect(out).toContain('3 cases');
  });

  test('a meta.json applies to every case beneath it', async () => {
    const { out } = cli(ROOT, '--print-spec');
    // The root excludes Regions, so all three inherit it.
    expect(out.match(/"ignoreSheets"/g) ?? []).toHaveLength(3);
  });

  test('a type meta.json reaches its own cases and not a sibling type', async () => {
    const { out } = cli(join(ROOT, 'reports', 'srq'), '--print-spec');
    // srq inherits the root layer but not global_standard_cat's tolerance.
    expect(out).toContain('"ignoreSheets"');
    expect(out).not.toContain('100000');
  });

  test('the type tolerance makes its cases pass, and a case.json overrides it', async () => {
    const { out, code } = cli(ROOT);
    // case_001 inherits the loose tolerance and passes; case_002 sets it back
    // to zero for itself alone and fails on the same drift.
    expect(out).toMatch(/✓ reports\/global_standard_cat\/case_001/);
    expect(out).toMatch(/✗ reports\/global_standard_cat\/case_002/);
    expect(code).toBe(1);
  });

  test('--print-spec names every layer that was applied', async () => {
    const { out } = cli(join(ROOT, 'reports', 'global_standard_cat', 'case_002'), '--print-spec');
    expect(out).toContain('meta.json');
    expect(out).toContain('case.json');
  });

  test('running one type runs only that type', async () => {
    const { out } = cli(join(ROOT, 'reports', 'srq'));
    expect(out).toContain('1 case');
    expect(out).not.toContain('global_standard_cat');
  });

  // These two mutate the tree, so they get their own outside it.
  test('a folder with no cases explains itself rather than failing silently', async () => {
    const empty = join(DIR, 'tree-empty');
    await rm(empty, { recursive: true, force: true });
    await mkdir(join(empty, 'nothing_here'), { recursive: true });

    const { out, code } = cli(empty);
    expect(code).toBe(1);
    expect(out).toContain('no runnable cases');
    expect(out).toContain('nothing_here');
  });

  test('a per-sheet setting written at the top level is refused, not ignored', async () => {
    // The easy mistake, and the worst kind: it parses, it is accepted, and it
    // has no effect, so the sheet goes on being compared without the setting
    // and nothing ever says so.
    const own = join(ROOT, 'reports', 'srq', 'case_001', 'case.json');
    await writeFile(own, '{ "Occupancy": { "keyColumns": ["Portfolio"] } }', 'utf8');
    const { out } = cli(join(ROOT, 'reports', 'srq'));

    expect(out).toContain('would do nothing');
    expect(out).toContain('Occupancy');
    expect(out).toContain('"sheets"');

    await rm(own, { force: true });
  });

  test('notes keyed with // are allowed, and more than one of them', async () => {
    const own = join(ROOT, 'reports', 'srq', 'case_001', 'case.json');
    await writeFile(own, '{ "//": "why", "//keys": "and a second note", "label": "x" }', 'utf8');
    const { out } = cli(join(ROOT, 'reports', 'srq'));

    expect(out).not.toContain('would do nothing');

    await rm(own, { force: true });
  });

  test('a config file saved with a UTF-8 BOM still parses', async () => {
    // PowerShell and several Windows editors add one by default, and it is
    // invisible in every editor that would show you the file.
    const own = join(ROOT, 'reports', 'srq', 'case_001', 'case.json');
    await writeFile(own, '﻿{ "label": "with a byte order mark" }', 'utf8');
    const { out } = cli(join(ROOT, 'reports', 'srq'));

    expect(out).not.toContain('not valid JSON');

    await rm(own, { force: true });
  });

  /**
   * A tree of its own, so naming can be rearranged without disturbing the
   * shared one. CSV keeps it to two files a person can read.
   */
  async function namedTree(meta: unknown, own: unknown): Promise<string> {
    const root = join(DIR, 'tree-named');
    const dir = join(root, 'reports', 'validation', 'case_003');
    await rm(root, { recursive: true, force: true });
    await mkdir(dir, { recursive: true });
    await writeFile(join(dir, 'golden.csv'), 'Region,Premium\nEast,100\n', 'utf8');
    await writeFile(join(dir, 'actual.csv'), 'Region,Premium\nEast,101\n', 'utf8');
    await writeJson(join(root, 'reports', 'validation', 'meta.json'), meta);
    await writeJson(join(dir, 'case.json'), own);
    return root;
  }

  test('a label titles the report and heads the case in the log', async () => {
    const root = await namedTree(
      { reportType: 'Validation Report' },
      { label: 'a peril column added between two others' },
    );
    const { out } = cli(root);

    // What kind of report, which case, and what the case is for -- then the
    // path underneath, since that is how the folder is found rather than what
    // the run is about.
    expect(out).toContain('✗ Validation Report · case_003 · a peril column added between two others');
    expect(out).toContain('    reports/validation/case_003');

    const report = await readFile(
      join(root, 'reports', 'validation', 'case_003', 'results', 'report.md'), 'utf8');
    expect(report).toContain('# a peril column added between two others');
    // The folder name is still on the page: it is what the log calls the case.
    expect(report).toContain('_Validation Report · case_003_');
  });

  test('a label written a folder above is not inherited by the case', async () => {
    // Every other setting inherits, and should. A label describes one case, so
    // one taken from above would head every case beneath it with the same
    // sentence and distinguish none of them.
    const root = await namedTree(
      { reportType: 'Validation Report', label: 'belongs to the folder, not the case' },
      {},
    );
    const { out } = cli(root);

    expect(out).toContain('✗ Validation Report · case_003');
    expect(out).not.toContain('belongs to the folder');
  });

  test('a report type alone still names the case, with no label invented', async () => {
    const root = await namedTree({ reportType: 'Validation Report' }, {});
    const { out } = cli(root);

    expect(out).toContain('✗ Validation Report · case_003');
    const report = await readFile(
      join(root, 'reports', 'validation', 'case_003', 'results', 'report.md'), 'utf8');
    expect(report).toContain('# case_003');
    expect(report).toContain('_Validation Report_');
  });

  test('a label repeating the folder name is not said twice', async () => {
    // How every case in a new tree starts out, before anyone has written a
    // description. "case_003 · case_003" would make the log worse.
    const root = await namedTree({ reportType: 'Validation Report' }, { label: 'case_003' });
    const { out } = cli(root);

    expect(out).toContain('✗ Validation Report · case_003\n');
    expect(out).not.toContain('case_003 · case_003');

    const report = await readFile(
      join(root, 'reports', 'validation', 'case_003', 'results', 'report.md'), 'utf8');
    expect(report).toContain('# case_003');
    expect(report).toContain('_Validation Report_');
  });

  test('a case at the root of the tree does not print its name twice', async () => {
    const root = join(DIR, 'tree-flat');
    const dir = join(root, 'case_001');
    await rm(root, { recursive: true, force: true });
    await mkdir(dir, { recursive: true });
    await writeFile(join(dir, 'golden.csv'), 'Region,Premium\nEast,100\n', 'utf8');
    await writeFile(join(dir, 'actual.csv'), 'Region,Premium\nEast,101\n', 'utf8');
    await writeJson(join(dir, 'case.json'), { label: 'a premium that moved' });

    const { out } = cli(root);
    // The case is its own path, so the line under the headline would repeat it.
    expect(out).toContain('✗ case_001 · a premium that moved\n    1 sheet failing');
  });

  test('a tree that names neither reads as the path alone', async () => {
    const root = await namedTree({}, {});
    const { out } = cli(root);

    expect(out).toContain('✗ reports/validation/case_003');
    // No path line under it: the headline already is the path.
    expect(out).not.toContain('    reports/validation/case_003');
  });

  /**
   * The shape a downloader produces: the role is the folder, and each file
   * keeps the name the source system gave it, timestamp and all.
   *
   *   output/<type>/case_00n/golden/rep_<stamp>.csv
   *                          current/rep_<stamp>.csv
   */
  async function downloadTree(): Promise<string> {
    const root = join(DIR, 'tree-folders');
    await rm(root, { recursive: true, force: true });
    for (const type of ['comparison_report', 'validation_report']) {
      for (const c of ['case_001', 'case_002', 'case_003']) {
        const dir = join(root, type, c);
        await mkdir(join(dir, 'golden'), { recursive: true });
        await mkdir(join(dir, 'current'), { recursive: true });
        await writeFile(join(dir, 'golden', 'rep_1786955263151.csv'), 'Region,Premium\nEast,100\n', 'utf8');
        await writeFile(join(dir, 'current', 'rep_1786957329031.csv'), 'Region,Premium\nEast,101\n', 'utf8');
      }
      await writeJson(join(root, type, 'meta.json'), { reportType: type });
    }
    await writeJson(join(root, 'meta.json'), {});
    return root;
  }

  test('a pair kept in golden/ and current/ folders is a case', async () => {
    const root = await downloadTree();
    const { out } = cli(root);

    expect(out).toContain('6 cases, 6 failing');
    // The files are named by the source system, not by this tool.
    const report = await readFile(
      join(root, 'comparison_report', 'case_001', 'results', 'report.md'), 'utf8');
    expect(report).toContain('rep_1786955263151.csv');
    expect(report).toContain('rep_1786957329031.csv');
  });

  test('two spreadsheets in a role folder stop the run rather than one being picked', async () => {
    const root = await downloadTree();
    const dir = join(root, 'comparison_report', 'case_001', 'current');
    await writeFile(join(dir, 'rep_1786999999999.csv'), 'Region,Premium\nEast,102\n', 'utf8');

    const { out, code } = cli(root);
    expect(out).toContain('current/ holds 2 spreadsheets');
    expect(out).toContain('it must hold exactly one');
    expect(code).toBe(1);
  });

  test('a half-built case is reported even while other cases run', async () => {
    // The downloader wrote golden/ and stopped. Before, this case simply
    // vanished from the run and the total read one lower.
    const root = await downloadTree();
    await rm(join(root, 'validation_report', 'case_002', 'current'), { recursive: true, force: true });

    const { out, code } = cli(root);
    expect(out).toContain('validation_report/case_002: a golden/ folder is here with no current/ folder');
    expect(out).toContain('5 cases, 5 failing, 1 could not be run');
    expect(code).toBe(1);
  });

  test('cases can be selected in a meta.json, and what is left out is counted', async () => {
    const root = await downloadTree();
    await writeJson(join(root, 'meta.json'), {
      cases: ['comparison_report/**', '!comparison_report/case_002'],
    });

    const { out } = cli(root);
    expect(out).toContain('comparison_report · case_001');
    expect(out).toContain('comparison_report · case_003');
    expect(out).not.toContain('case_002');
    expect(out).toContain('2 cases, 2 failing — 4 not selected by "cases"');
  });

  test('a report type can narrow its own cases, on top of the root', async () => {
    const root = await downloadTree();
    await writeJson(join(root, 'validation_report', 'meta.json'), {
      reportType: 'validation_report',
      cases: ['case_003'],
    });

    const { out } = cli(root);
    // Paths are relative to the file that carries them, so the type names its
    // cases directly while the root would name them by type.
    expect(out).toContain('4 cases, 4 failing — 2 not selected by "cases"');
    expect(out).toContain('validation_report · case_003');
  });

  test('a selection that matches nothing says so instead of passing', async () => {
    const root = await downloadTree();
    await writeJson(join(root, 'meta.json'), { cases: ['no_such_type/**'] });

    const { out, code } = cli(root);
    expect(out).toContain('selected none of the 6 cases');
    expect(out).toContain('meta.json');
    expect(code).toBe(1);
  });

  test('blessing a folder pair keeps the name the new report came with', async () => {
    const root = await downloadTree();
    const dir = join(root, 'comparison_report', 'case_001');

    const { out } = cli(dir, '--bless');
    expect(out).toContain('golden replaced by rep_1786957329031.csv, and rep_1786955263151.csv removed');

    const golden = await readdir(join(dir, 'golden'));
    expect(golden).toEqual(['rep_1786957329031.csv']);
    // And the pair now matches, which is what blessing means.
    expect(cli(dir).out).toContain('identical');
  });

  test('clean finds a results folder beside a golden/ folder', async () => {
    const root = await downloadTree();
    cli(root);

    const r = spawnSync('node', ['scripts/clean-results.mjs', root, '--dry'], { encoding: 'utf8' });
    expect(r.stdout).toContain('6 results folder(s) would be removed');
  });

  test('malformed JSON is reported with the file that contains it', async () => {
    const dir = join(DIR, 'tree-bad', 'case_001');
    await rm(join(DIR, 'tree-bad'), { recursive: true, force: true });
    await mkdir(dir, { recursive: true });
    const golden = await buildMultiSheet('tree-bad-g.xlsx');
    await cp(golden, join(dir, 'golden.xlsx'));
    await cp(golden, join(dir, 'actual.xlsx'));
    await writeFile(join(dir, 'case.json'), '{ not json', 'utf8');

    const { out, code } = cli(join(DIR, 'tree-bad'));
    expect(code).toBe(1);
    expect(out).toContain('case.json');
    expect(out).toContain('not valid JSON');
  });

  test('--write-meta writes a starting meta.json instead of one being typed out', async () => {
    const dir = join(DIR, 'tree-meta', 'my_type');
    await rm(join(DIR, 'tree-meta'), { recursive: true, force: true });
    await mkdir(join(dir, 'case_001'), { recursive: true });
    const golden = await buildMultiSheet('tree-meta-g.xlsx');
    await cp(golden, join(dir, 'case_001', 'golden.xlsx'));
    await cp(golden, join(dir, 'case_001', 'actual.xlsx'));

    const { out, code } = cli('--write-meta', dir);
    expect(code).toBe(0);
    expect(out).toContain('written to');

    const written = JSON.parse(await readFile(join(dir, 'meta.json'), 'utf8'));
    expect(written.reportType).toBe('My Type');
    // Per-sheet configuration is exactly what it must NOT invent: those
    // entries go stale when a report changes shape, and a generated one is
    // indistinguishable from one somebody meant.
    expect(written.sheets).toBeUndefined();
    // Notes are keyed with //, which the config reader treats as comments, so
    // the file it writes is a file it can read.
    expect(cli(dir).code).toBe(0);
  });

  test('--write-meta refuses to overwrite a config somebody wrote', async () => {
    const dir = join(DIR, 'tree-meta', 'my_type');
    await writeJson(join(dir, 'meta.json'), { reportType: 'Mine', metadata: ['Report ID'] });

    const { out, code } = cli('--write-meta', dir);
    expect(code).toBe(1);
    expect(out).toContain('already has settings in it');
    expect(out).toContain('reportType, metadata');
    // And left it alone.
    const kept = JSON.parse(await readFile(join(dir, 'meta.json'), 'utf8'));
    expect(kept.reportType).toBe('Mine');
  });

  test('--write-expect records what was verified, and the guard then holds', async () => {
    const dir = join(ROOT, 'reports', 'srq', 'case_001');
    await writeJson(join(dir, 'case.json'), { label: 'a labelled case' });

    const first = cli('--write-expect', join(ROOT, 'reports', 'srq'));
    expect(first.out).toContain('expectations recorded');

    const written = JSON.parse(await readFile(join(dir, 'case.json'), 'utf8'));
    // Written beside what was already there, not over it.
    expect(written.label).toBe('a labelled case');
    expect(Object.keys(written.expect).length).toBeGreaterThan(0);

    // The run it was recorded from still passes its own guard.
    expect(cli(join(ROOT, 'reports', 'srq')).out).not.toContain('expected');

    await rm(join(dir, 'case.json'), { force: true });
  });

  test('a table that stops being compared fails the run and is named', async () => {
    // The failure this exists to prevent: coverage shrinking with nothing to
    // show for it but a smaller number in a summary line.
    const dir = join(ROOT, 'reports', 'srq', 'case_001');
    await writeJson(join(dir, 'case.json'), {
      expect: { Policies: ['A1:F6', 'A40:E60'], Premiums: 2 },
    });

    const { out, code } = cli(join(ROOT, 'reports', 'srq'));
    expect(code).toBe(1);
    expect(out).toContain('integrity');

    const report = await readFile(join(dir, 'results', 'report.md'), 'utf8');
    expect(report).toContain('Policies: expected 2 table(s), compared 1');
    expect(report).toContain('not compared: A40:E60');
    expect(report).toContain('Premiums: expected 2 table(s) compared by name and key, found 1');

    await rm(join(dir, 'case.json'), { force: true });
  });

  test('a table added to a single-table sheet is added, not swapped in', async () => {
    // The sheet carries one table, so its header row and key sit at the sheet
    // level with no "tables" block -- which is what makes reports read
    // "Ledger" rather than "Ledger · Table 1". Declaring a table on such a
    // sheet used to replace it: one entry written to check a title block also
    // stopped the sheet's real table being compared, and nothing said so.
    const dir = join(DIR, 'tree-add', 'case_001');
    await rm(join(DIR, 'tree-add'), { recursive: true, force: true });
    await mkdir(dir, { recursive: true });

    const build = async (name: string, drift: number) => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Ledger');
      ws.addRow(['Ref', 'Name', 'Amount']);
      ws.addRow(['R-1', 'a', 10]);
      ws.addRow(['R-2', 'b', 20]);
      ws.addRow(['R-3', 'c', 30]);
      // Contiguous, so detection reads all of it as one table.
      ws.addRow(['Band', 'Label', 'Value']);
      ws.addRow(['B-1', 'low', 1 + drift]);
      ws.addRow(['B-2', 'high', 2]);
      const path = join(DIR, name);
      await wb.xlsx.writeFile(path);
      return path;
    };
    await cp(await build('tree-add-g.xlsx', 0), join(dir, 'golden.xlsx'));
    await cp(await build('tree-add-a.xlsx', 5), join(dir, 'actual.xlsx'));

    await writeJson(join(dir, 'case.json'), {
      sheets: { Ledger: { tables: { Bands: { headerRow: 5, keyColumns: ['Band'] } } } },
    });

    cli(join(DIR, 'tree-add'));

    const report = await readFile(join(dir, 'results', 'report.md'), 'utf8');
    expect(report).toContain('## What was verified (2)');
    // Both: the sheet's own table, named after the sheet, and the added one.
    expect(report).toContain('| Ledger | `A1:C4`');
    expect(report).toContain('| Ledger · Bands | `A5:C7`');
    // The original is bounded above the added table rather than running on
    // into it, so no cell is compared twice.
    expect(report).toContain('B-1');
  });
});
