/**
 * The two downloads compared as archives, not as spreadsheets.
 *
 * An .xlsx is a zip of XML parts. Every other check in this tool reads it
 * through a spreadsheet model -- cells, fills, validations -- and a model only
 * sees what it was built to see. Drawings, charts, merged cells, column widths,
 * defined names, sheet protection, themes and print setup all live in those
 * parts, and every one of them could change without a single check noticing.
 *
 * Comparing the parts byte for byte needs no model and no descriptor. It cannot
 * say what changed in a way anybody would enjoy reading, but it cannot miss
 * anything either, and it costs nothing to be sure.
 *
 * This is also what makes the capture's openpyxl round trip survivable. The
 * files under test are re-saves rather than the bytes the app produced, so a
 * feature openpyxl drops is invisible -- but BOTH sides go through the same
 * pipeline and lose the same things, so drift in whatever survives is still
 * caught in full. What it cannot tell you is whether something was already
 * missing when the golden was blessed. That needs the raw download.
 *
 * Measured before it was written: across ten cases the two files differ in
 * exactly one part, docProps/core.xml, and only in its timestamps.
 */
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

/**
 * Parts whose bytes move on their own, and how to read past that.
 *
 * Only the timestamps are dropped from core.xml -- not the whole part -- so a
 * changed title, creator or category is still compared.
 */
const NORMALISE = {
  'docProps/core.xml': (text) => text
    .replace(/<dcterms:created[^>]*>[^<]*<\/dcterms:created>/g, '<dcterms:created/>')
    .replace(/<dcterms:modified[^>]*>[^<]*<\/dcterms:modified>/g, '<dcterms:modified/>'),
};

const TEXT = /\.(xml|rels)$/i;

async function parts(path) {
  const zip = await JSZip.loadAsync(await readFile(path));
  const out = new Map();
  for (const name of Object.keys(zip.files)) {
    const entry = zip.files[name];
    if (entry.dir) continue;
    if (TEXT.test(name)) {
      const text = await entry.async('string');
      out.set(name, (NORMALISE[name] ?? ((t) => t))(text));
    } else {
      out.set(name, await entry.async('base64'));
    }
  }
  return out;
}

/** Who wrote the file, as the file itself claims. */
function producer(app) {
  const m = /<Application>([^<]*)<\/Application>/.exec(app ?? '');
  return m ? m[1] : null;
}

/**
 * @param currentPath the download under test
 * @param goldenPath  the download it is judged against
 * @param ignore      extra part names to skip, for a producer with its own noise
 */
export async function checkParts(currentPath, goldenPath, ignore = []) {
  const skip = new Set(ignore);
  const [a, b] = await Promise.all([parts(goldenPath), parts(currentPath)]);

  const findings = [];
  const names = new Set([...a.keys(), ...b.keys()]);
  let compared = 0;

  // Both files must have come off the same pipeline, or a byte comparison
  // between them compares the writers as much as the downloads. Reported as the
  // single fact it is: a re-save changes every part, and listing all of them
  // would bury the one line that explains why.
  const wrote = {
    golden: producer(a.get('docProps/app.xml')),
    current: producer(b.get('docProps/app.xml')),
  };
  if (wrote.golden !== wrote.current) {
    return {
      ok: false,
      total: names.size,
      compared: 0,
      producer: wrote.current,
      findings: [{
        part: 'docProps/app.xml',
        problem: 'the two files were written by different producers, so every part differs for '
          + 'one reason and none of them was compared',
        golden: wrote.golden ?? 'unstated',
        current: wrote.current ?? 'unstated',
      }],
    };
  }

  for (const name of [...names].sort()) {
    if (skip.has(name)) continue;
    const g = a.get(name);
    const c = b.get(name);
    if (g === undefined) { findings.push({ part: name, problem: 'new since the golden' }); continue; }
    if (c === undefined) { findings.push({ part: name, problem: 'gone since the golden' }); continue; }
    compared++;
    if (g !== c) {
      findings.push({
        part: name,
        problem: 'differs',
        golden: `${g.length} chars`,
        current: `${c.length} chars`,
      });
    }
  }

  return { ok: findings.length === 0, total: names.size, compared, producer: wrote.current, findings };
}
