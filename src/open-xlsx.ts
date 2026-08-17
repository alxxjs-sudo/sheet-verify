import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

/**
 * Opens a workbook, working around ExcelJS failing on some real files.
 *
 * ExcelJS 4.4.0 reconciles every drawing against its relationships after
 * parsing, and on workbooks where that lookup comes back empty it throws
 * `Cannot read properties of undefined (reading 'anchors')` before any cell is
 * reachable. It has had no release since October 2023, so the fix has to be
 * here.
 *
 * Drawings are images and charts. Nothing in a comparison reads them -- only
 * cell values and formulas matter -- so when the direct read fails the file is
 * re-read from an in-memory copy with the drawing parts removed. The file on
 * disk is never touched, and a workbook that opens normally never pays the
 * cost of the second pass.
 */

const DRAWING_PARTS = /^xl\/(drawings|media)\//;
const SHEET_RELS = /^xl\/worksheets\/_rels\/.*\.rels$/;
const SHEET_XML = /^xl\/worksheets\/sheet\d+\.xml$/;

/** An in-memory copy of the workbook with every drawing part taken out. */
export async function withoutDrawings(buffer: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer);

  for (const name of Object.keys(zip.files)) {
    if (DRAWING_PARTS.test(name)) zip.remove(name);
  }

  for (const name of Object.keys(zip.files)) {
    const file = zip.file(name);
    if (!file) continue;

    if (SHEET_RELS.test(name)) {
      // The relationship pointing a sheet at its drawing.
      const xml = await file.async('string');
      zip.file(name, xml.replace(/<Relationship[^>]*\/drawing"[^>]*\/>/g, ''));
    } else if (SHEET_XML.test(name)) {
      // The <drawing r:id="..."/> element that references it.
      const xml = await file.async('string');
      zip.file(name, xml.replace(/<drawing[^>]*\/>/g, ''));
    } else if (name === '[Content_Types].xml') {
      const xml = await file.async('string');
      zip.file(name, xml.replace(/<Override[^>]*drawings\/[^>]*\/>/g, ''));
    }
  }

  return zip.generateAsync({ type: 'nodebuffer' });
}

/** True when a workbook could only be read with its drawings removed. */
export const openedWithoutDrawings = new WeakSet<ExcelJS.Workbook>();

export async function openWorkbook(path: string): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.readFile(path);
    return wb;
  } catch (direct) {
    let cleaned: Buffer;
    try {
      cleaned = await withoutDrawings(await readFile(path));
    } catch {
      throw new Error(`sheet-verify: could not read "${path}" — ${(direct as Error).message}`);
    }

    const retry = new ExcelJS.Workbook();
    try {
      // ExcelJS types `load` against the DOM Buffer rather than Node's.
      await retry.xlsx.load(cleaned as unknown as ArrayBuffer);
    } catch (second) {
      throw new Error(
        `sheet-verify: could not read "${path}". ExcelJS failed with ` +
        `"${(direct as Error).message}", and again with the drawings removed: ` +
        `"${(second as Error).message}". The file may be corrupt, or may use a ` +
        'feature this reader does not handle -- registerReader() replaces it.',
      );
    }
    openedWithoutDrawings.add(retry);
    return retry;
  }
}
