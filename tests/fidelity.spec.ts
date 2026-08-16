import { test, expect } from '@playwright/test';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { checkFidelity, classify, formatFidelity } from '../src/fidelity.js';
import { buildWorkbook, DIR } from './fixtures.js';

/** Adds template features by hand, since ExcelJS cannot author all of them. */
async function injectFeatures(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(path));
  let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');

  const extras =
    '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>' +
    '<conditionalFormatting sqref="D2:D6">' +
    '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100000</formula></cfRule>' +
    '</conditionalFormatting>' +
    '<dataValidations count="1">' +
    '<dataValidation type="list" sqref="C2:C6" allowBlank="1"><formula1>"Sofia,Varna,Ruse"</formula1></dataValidation>' +
    '</dataValidations>';
  xml = xml.replace('</sheetData>', `</sheetData>${extras}`);
  zip.file('xl/worksheets/sheet1.xml', xml);

  // The cfRule points at dxfId 0, so styles.xml must define it or the file is
  // malformed and ExcelJS refuses to open it.
  const styles = zip.file('xl/styles.xml');
  if (styles) {
    let s = await styles.async('string');
    const dxfs = '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>';
    s = s.includes('<dxfs')
      ? s.replace(/<dxfs[^>]*\/>|<dxfs[\s\S]*?<\/dxfs>/, dxfs)
      : s.replace('</styleSheet>', `${dxfs}</styleSheet>`);
    zip.file('xl/styles.xml', s);
  }
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }));
}

test.describe('what ExcelJS actually preserves', () => {
  // These are measured, not assumed. They also guard an ExcelJS upgrade: if a
  // future version starts dropping any of this, these fail rather than
  // silently corrupting real templates.
  test('conditional formatting, data validation and merged cells survive', async () => {
    const f = await buildWorkbook('f1.xlsx');
    await injectFeatures(f);
    const r = await checkFidelity(f);
    expect(r.findings.filter((x) => x.severity === 'critical')).toEqual([]);
    expect(r.ok).toBe(true);
  });

  test('images, auto-filters and frozen panes survive', async () => {
    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
      'base64',
    );
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow(['Id', 'Val']);
    ws.addRow(['A', 1]);
    ws.addImage(wb.addImage({ buffer: png, extension: 'png' }), 'D2:E6');
    ws.autoFilter = 'A1:B1';
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    const path = join(DIR, 'f2-rich.xlsx');
    await wb.xlsx.writeFile(path);

    const r = await checkFidelity(path);
    expect(r.findings.filter((x) => x.severity === 'critical')).toEqual([]);
    expect(r.ok).toBe(true);
  });
});

test.describe('the gate reports loss when there is loss', () => {
  const parts = (o: Record<string, string>) => new Map(Object.entries(o));

  test('a dropped drawing part is critical', () => {
    const f = classify(
      parts({ 'xl/drawings/drawing1.xml': '<x/>', 'xl/worksheets/sheet1.xml': '<w/>' }),
      parts({ 'xl/worksheets/sheet1.xml': '<w/>' }),
    );
    expect(f).toHaveLength(1);
    expect(f[0]).toMatchObject({ severity: 'critical' });
    expect(f[0]!.detail).toContain('drawings');
  });

  test('a dropped in-sheet feature is critical and counted', () => {
    const f = classify(
      parts({ 'xl/worksheets/sheet1.xml': '<dataValidation type="list"/><dataValidation type="list"/>' }),
      parts({ 'xl/worksheets/sheet1.xml': '<x/>' }),
    );
    expect(f[0]!.severity).toBe('critical');
    expect(f[0]!.detail).toContain('data validation');
    expect(f[0]!.detail).toContain('2 before, 0 after');
  });

  test('calcChain is reported as harmless, not as a loss', () => {
    const f = classify(parts({ 'xl/calcChain.xml': '<c/>' }), parts({}));
    expect(f).toHaveLength(1);
    expect(f[0]!.severity).toBe('info');
  });

  test('an unrecognised dropped part is a warning, not a failure', () => {
    const f = classify(parts({ 'xl/printerSettings/printerSettings1.bin': '' }), parts({}));
    expect(f[0]!.severity).toBe('warning');
  });

  test('pivot tables are classified as critical', () => {
    const f = classify(parts({ 'xl/pivotTables/pivotTable1.xml': '<p/>' }), parts({}));
    expect(f[0]).toMatchObject({ severity: 'critical' });
    expect(f[0]!.detail).toContain('pivot tables');
  });
});

test.describe('gate output', () => {
  test('a file the library cannot open is the most critical finding of all', async () => {
    const r = await checkFidelity('does-not-exist.xlsx');
    expect(r.ok).toBe(false);
    expect(r.findings[0]!.severity).toBe('critical');
    expect(r.findings[0]!.detail).toContain('cannot round-trip');
  });

  test('a failing report says what to do instead', () => {
    const text = formatFidelity([
      { file: 't.xlsx', ok: false, findings: [{ severity: 'critical', detail: 'lost charts — xl/charts/chart1.xml' }] },
    ]);
    expect(text).toContain('lose content on save');
    expect(text).toContain('data sheet');
  });

  test('a clean report says so plainly', () => {
    const text = formatFidelity([{ file: 't.xlsx', ok: true, findings: [] }]);
    expect(text).toContain('no detectable loss');
    expect(text).toContain('safe');
  });
});
