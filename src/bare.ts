import { readFile, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

/**
 * Puts a workbook back the way the generator wrote it: formulas present,
 * results absent.
 *
 * The reports arrive from the generator with no calculated results at all --
 * openpyxl writes `<f>…</f><v />` and sets `fullCalcOnLoad="1"`, which tells
 * Excel to work the numbers out on open. Opening one and saving it makes Excel
 * write all of them in, so a report someone looked at no longer resembles the
 * file the generator produced.
 *
 * That only matters because a comparison holds two files side by side. A bare
 * baseline against a saved report shows every formula in the file as a value
 * change -- not because anything changed, but because one side was opened. In
 * production neither side is ever opened, so the pair matches; it is planting a
 * test edit that forces a file through Excel and breaks the symmetry.
 *
 * Running this over the edited file restores it, leaving the edit itself as the
 * only difference. It cannot restore the file byte for byte -- Excel rewrites
 * the whole package, and no amount of unpicking undoes that -- but it restores
 * everything a comparison reads.
 *
 * Nothing is lost that Excel will not regenerate: open the file again and the
 * results come straight back.
 */

const SHEET_XML = /^xl\/worksheets\/sheet\d+\.xml$/;
const CALC_CHAIN = 'xl/calcChain.xml';

/** A whole `<c>` element, self-closing or not. */
const CELL = /<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
/** `<f>`, `<f …>` or `<f …/>` -- a formula in any of its forms. */
const HAS_FORMULA = /<f[\s>/]/;
/** A stored result. `<v />` is the empty form the generator writes. */
const STORED_RESULT = /<v>[\s\S]*?<\/v>/;
/** The type attribute describing a stored result, meaningless once it is gone. */
const RESULT_TYPE = /\s+t="(?:str|e|b|s|n)"/;

export interface CachedValueState {
  /** Formula cells carrying a calculated result. */
  cached: number;
  /** Formula cells with no result, as the generator writes them. */
  bare: number;
  /** The `<Application>` that last wrote the file. */
  application: string;
  /** Excel is being told to recalculate on open. */
  fullCalcOnLoad: boolean;
}

/**
 * Whether a file carries calculated results. Used to spot a pair whose two
 * sides were produced differently, which is worth one line in a report rather
 * than a value difference on every formula in the file.
 */
export async function cachedValueState(buffer: Buffer): Promise<CachedValueState> {
  const zip = await JSZip.loadAsync(buffer);
  let cached = 0;
  let bare = 0;

  for (const name of Object.keys(zip.files)) {
    if (!SHEET_XML.test(name)) continue;
    const xml = await zip.file(name)!.async('string');
    for (const [cell] of xml.matchAll(CELL)) {
      if (!HAS_FORMULA.test(cell)) continue;
      if (STORED_RESULT.test(cell)) cached++;
      else bare++;
    }
  }

  const workbook = (await zip.file('xl/workbook.xml')?.async('string')) ?? '';
  const app = (await zip.file('docProps/app.xml')?.async('string')) ?? '';

  return {
    cached,
    bare,
    application: /<Application>([^<]*)</.exec(app)?.[1] ?? '',
    fullCalcOnLoad: /fullCalcOnLoad="1"/.test(workbook),
  };
}

/** Strips stored results from one sheet, returning how many it removed. */
function bareSheet(xml: string): { xml: string; stripped: number } {
  let stripped = 0;
  const out = xml.replace(CELL, (cell) => {
    if (!HAS_FORMULA.test(cell) || !STORED_RESULT.test(cell)) return cell;
    stripped++;
    // `<v />` rather than nothing: that is what the generator writes, and a
    // formula cell with no <v> at all is a shape neither tool produces.
    return cell.replace(STORED_RESULT, '<v />').replace(RESULT_TYPE, '');
  });
  return { xml: out, stripped };
}

export interface BareResult {
  buffer: Buffer;
  /** Formula cells whose result was removed. */
  stripped: number;
  /** The calculation chain was present and has been dropped. */
  droppedCalcChain: boolean;
  /** `fullCalcOnLoad` was missing and has been put back. */
  restoredFullCalc: boolean;
}

/**
 * An in-memory copy of the workbook with every calculated result removed. The
 * file on disk is untouched; `makeBare` writes it back.
 */
export async function stripCachedValues(buffer: Buffer): Promise<BareResult> {
  const zip = await JSZip.loadAsync(buffer);
  let stripped = 0;

  for (const name of Object.keys(zip.files)) {
    if (!SHEET_XML.test(name)) continue;
    const result = bareSheet(await zip.file(name)!.async('string'));
    stripped += result.stripped;
    if (result.stripped) zip.file(name, result.xml);
  }

  // The calculation chain is Excel's cache of what to evaluate in what order.
  // It describes results that no longer exist, and Excel rebuilds it, so it
  // goes -- along with the two places that declare it, since a part referenced
  // but absent makes the package invalid.
  const droppedCalcChain = Boolean(zip.file(CALC_CHAIN));
  if (droppedCalcChain) {
    zip.remove(CALC_CHAIN);

    const types = zip.file('[Content_Types].xml');
    if (types) {
      const xml = await types.async('string');
      zip.file('[Content_Types].xml', xml.replace(/<Override[^>]*calcChain[^>]*\/>/g, ''));
    }
    const rels = zip.file('xl/_rels/workbook.xml.rels');
    if (rels) {
      const xml = await rels.async('string');
      zip.file('xl/_rels/workbook.xml.rels', xml.replace(/<Relationship[^>]*calcChain[^>]*\/>/g, ''));
    }
  }

  // Ask Excel to work the numbers out on open, as the generator does. Without
  // this a reader could see a workbook full of formulas and no results and
  // take the blanks at face value. Anything that rewrites a workbook can drop
  // the flag -- ExcelJS does -- so it is restored whether or not a result was
  // stripped, and reported, since on its own it is the only thing that changed.
  const workbook = zip.file('xl/workbook.xml');
  let restoredFullCalc = false;
  if (workbook) {
    let xml = await workbook.async('string');
    if (!/fullCalcOnLoad="1"/.test(xml)) {
      // Exactly one of these, never two: running the self-closing rewrite and
      // then the open-tag one turns `<calcPr …/>` into
      // `<calcPr …/ fullCalcOnLoad="1">`, which is not XML.
      if (/<calcPr\b[^>]*\/>/.test(xml)) {
        xml = xml.replace(/<calcPr\b([^>]*?)\s*\/>/, '<calcPr$1 fullCalcOnLoad="1"/>');
      } else if (/<calcPr\b[^>]*>/.test(xml)) {
        xml = xml.replace(/<calcPr\b([^>]*?)>/, '<calcPr$1 fullCalcOnLoad="1">');
      } else {
        xml = xml.replace('</workbook>', '<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>');
      }
      restoredFullCalc = true;
    }
    zip.file('xl/workbook.xml', xml);
  }

  return {
    // DEFLATE explicitly: JSZip stores parts uncompressed unless told
    // otherwise, and an xlsx written that way is several times the size of the
    // one it replaced -- a 2.6 MB report came back as 23 MB. Excel reads either
    // happily, so nothing fails; the file just quietly bloats every time it is
    // rewritten.
    buffer: await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    }),
    stripped,
    droppedCalcChain,
    restoredFullCalc,
  };
}

/**
 * Rewrites a workbook in place with its calculated results removed. Returns
 * null when the file already had none, so a caller can say "nothing to do"
 * rather than rewriting a file it did not need to touch.
 */
export async function makeBare(path: string): Promise<BareResult | null> {
  const result = await stripCachedValues(await readFile(path));
  if (!result.stripped && !result.droppedCalcChain && !result.restoredFullCalc) return null;
  await writeFile(path, result.buffer);
  return result;
}
