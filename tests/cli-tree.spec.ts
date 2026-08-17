import { test, expect } from '@playwright/test';
import { mkdir, rm, writeFile, cp } from 'node:fs/promises';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
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
});
