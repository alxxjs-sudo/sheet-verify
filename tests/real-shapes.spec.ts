import { test, expect } from '@playwright/test';
import { join } from 'node:path';
import { mkdir, readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { detectWorkbook, openWorkbook, withoutDrawings } from '../src/index.js';
import { DIR } from './fixtures.js';

/**
 * Shapes taken from real reports rather than invented ones. Both of these
 * stopped the tool dead on actual files: a workbook ExcelJS refuses to open,
 * and a table sitting under a title row.
 */

/** A 1x1 transparent PNG, enough for ExcelJS to write a real drawing part. */
const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);

async function workbookWithImage(name: string): Promise<string> {
  await mkdir(DIR, { recursive: true });
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow(['Id', 'Name', 'Amount']);
  ws.addRow(['A-1', 'Ivanov', 100]);
  ws.addRow(['A-2', 'Petrov', 200]);

  const id = wb.addImage({ buffer: PNG as any, extension: 'png' });
  ws.addImage(id, { tl: { col: 4, row: 1 }, ext: { width: 40, height: 40 } });

  const path = join(DIR, name);
  await wb.xlsx.writeFile(path);
  return path;
}

const partsOf = async (buf: Buffer) =>
  Object.keys((await JSZip.loadAsync(buf)).files);

test.describe('workbooks ExcelJS cannot open', () => {
  test('withoutDrawings removes every drawing part and leaves the data intact', async () => {
    const path = await workbookWithImage('rs-image.xlsx');

    const before = await partsOf(await readFile(path));
    expect(before.some((p) => p.startsWith('xl/drawings/'))).toBe(true);
    expect(before.some((p) => p.startsWith('xl/media/'))).toBe(true);

    const cleaned = await withoutDrawings(await readFile(path));
    const after = await partsOf(cleaned);
    expect(after.some((p) => p.startsWith('xl/drawings/'))).toBe(false);
    expect(after.some((p) => p.startsWith('xl/media/'))).toBe(false);

    // The point of stripping them: the cells still read.
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(cleaned as unknown as ArrayBuffer);
    const ws = wb.getWorksheet('Data')!;
    expect(ws.getRow(1).getCell(1).value).toBe('Id');
    expect(ws.getRow(3).getCell(3).value).toBe(200);
  });

  test('the sheet no longer references a drawing that is gone', async () => {
    const path = await workbookWithImage('rs-image2.xlsx');
    const zip = await JSZip.loadAsync(await withoutDrawings(await readFile(path)));

    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).not.toContain('<drawing');

    const rels = zip.file('xl/worksheets/_rels/sheet1.xml.rels');
    if (rels) expect(await rels.async('string')).not.toContain('/drawing"');

    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(types).not.toContain('drawings/');
  });

  test('openWorkbook reads a normal workbook without stripping anything', async () => {
    const path = await workbookWithImage('rs-image3.xlsx');
    const wb = await openWorkbook(path);
    expect(wb.getWorksheet('Data')!.getRow(2).getCell(2).value).toBe('Ivanov');
  });

  test('a file that is not a workbook at all fails with a message naming it', async () => {
    const path = join(DIR, 'rs-not-a-workbook.xlsx');
    const { writeFile } = await import('node:fs/promises');
    await writeFile(path, 'this is not a zip', 'utf8');

    await expect(openWorkbook(path)).rejects.toThrow(/could not read .*rs-not-a-workbook/);
  });
});

test.describe('a table under a title row', () => {
  /**
   * The shape that broke on a real report: metadata lines at the top, a title
   * immediately above the header with no blank row between them, and the data
   * below. Taking the first row of the block as the header discards all of it.
   */
  async function titledReport(name: string): Promise<string> {
    await mkdir(DIR, { recursive: true });
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Loss by Contract');

    ws.addRow(['ProformaOverride']);           // 1  metadata, one cell each
    ws.addRow(['All values are in USD']);      // 2
    ws.addRow([]);                             // 3  blank
    ws.addRow(['Loss AEP VaR Contract']);      // 4  title, one cell
    ws.addRow(['Contract ID', 'Company', 'Layer', 'Loss']);  // 5  the header
    ws.addRow(['199293', 'Edison', '35M xs 25M', 1000]);     // 6  data
    ws.addRow(['199294', 'Edison', '105M xs 60M', 2000]);
    ws.addRow(['199295', 'Edison', '197.5M xs 165M', 3000]);

    const path = join(DIR, name);
    await wb.xlsx.writeFile(path);
    return path;
  }

  test('the header is found under the title, not mistaken for it', async () => {
    const path = await titledReport('rs-titled.xlsx');
    const [sheet] = await detectWorkbook(path);

    // Rows 1-2 are a block of one-cell metadata and yield nothing; the block
    // from row 4 is the real table.
    expect(sheet!.tables).toHaveLength(1);
    const t = sheet!.tables[0]!;

    expect(t.headerRow).toBe(5);
    expect(t.headers.filter(Boolean)).toEqual(['Contract ID', 'Company', 'Layer', 'Loss']);
    expect(t.rows).toBe(3);
    expect(t.keyColumns).toEqual(['Contract ID']);
  });

  test('a header on the first row of its block is still preferred over its data', async () => {
    await mkdir(DIR, { recursive: true });
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Plain');
    ws.addRow(['Ref', 'Name', 'Value']);
    ws.addRow(['R-1', 'a', 1]);
    ws.addRow(['R-2', 'b', 2]);
    const path = join(DIR, 'rs-plain.xlsx');
    await wb.xlsx.writeFile(path);

    const [sheet] = await detectWorkbook(path);
    // Header and data rows are equally full, so the earliest wins.
    expect(sheet!.tables[0]!.headerRow).toBe(1);
    expect(sheet!.tables[0]!.keyColumns).toEqual(['Ref']);
  });
});
