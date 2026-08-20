import { colToNum, numToCol } from './a1.js';

/**
 * Report metadata: the cells that identify a run rather than describe it.
 *
 * A report name, a report id, a creation timestamp -- these differ between any
 * two runs by construction. Comparing them produces a difference on every
 * single run, which is the worst possible kind of finding: always present,
 * never actionable, and after the third report nobody reads past it. Worse, it
 * makes a clean run impossible, so "did anything change?" stops having a yes.
 *
 * So they are read and skipped, and the report says which ones and why. Skipped
 * is not the same as unseen: every one is listed with both values, so a reader
 * can still notice that an id moved when they did not expect it to.
 *
 * "By construction" is the whole test, and it is narrower than it sounds. A
 * creator name reads like run identity and is not one: when the same account
 * generates every report it is expected to hold still, so a change there says
 * the wrong account ran it. Same for anything the numbers depend on -- the view
 * of risk, the currency, the model version, the as-at date of the data. Those
 * look like header furniture and are not: if one moves, the figures underneath
 * it should have moved too. All of it belongs in the verdict.
 */

/** A label to look for, optionally confined to one sheet. */
export interface MetadataLabel {
  /** Canonical sheet name, or "" for every sheet. */
  sheet: string;
  /** Normalised label text, as written -- kept for the report. */
  label: string;
  /** Set when the label holds a `*`; matches the leading run of text. */
  glob?: RegExp;
}

export interface MetadataRules {
  /** Labels matched against a cell's text. */
  labels: MetadataLabel[];
  /** Addresses, as `sheet` (canonical, or "" for any) -> set of addresses. */
  addresses: Map<string, Set<string>>;
  /** Whether anything at all was configured. */
  any: boolean;
}

const ADDRESS = /^(?:(.+)!)?\$?([A-Z]{1,3})\$?(\d+)$/i;

/** Lowercase, collapse whitespace, drop the quotes a fused formula carries. */
const norm = (s: string): string =>
  s.replace(/["']/g, '').replace(/\s+/g, ' ').trim().toLowerCase();

const canonSheet = (s: string): string => s.trim().toLowerCase();

/**
 * Splits the configured patterns into labels and addresses.
 *
 * An entry that parses as a cell reference is one: `A2`, `Cover!A2`. Anything
 * else is a label, matched against the text of a cell. `Report ID` and
 * `Report ID:` mean the same thing, since whether the sheet writes the colon
 * is a styling choice.
 */
export function parseMetadata(patterns: string[] | undefined): MetadataRules {
  const labels: MetadataLabel[] = [];
  const addresses = new Map<string, Set<string>>();

  for (const raw of patterns ?? []) {
    const pattern = raw.trim();
    if (!pattern) continue;

    const m = ADDRESS.exec(pattern);
    if (m) {
      const [, qualifier, letters, digits] = m;
      const sheet = canonSheet((qualifier ?? '').replace(/^'|'$/g, ''));
      const address = `${letters!.toUpperCase()}${digits}`;
      const set = addresses.get(sheet) ?? new Set<string>();
      set.add(address);
      addresses.set(sheet, set);
      continue;
    }
    // "Report Info!Report ID" confines a label to one sheet, for a word that
    // means run identity in a header block and a column heading elsewhere.
    const bang = pattern.lastIndexOf('!');
    const sheet = bang > 0 ? canonSheet(pattern.slice(0, bang).replace(/^'|'$/g, '')) : '';
    const label = norm(bang > 0 ? pattern.slice(bang + 1) : pattern).replace(/:$/, '');
    if (label) labels.push({ sheet, label, ...(label.includes('*') ? { glob: globOf(label) } : {}) });
  }

  return { labels, addresses, any: labels.length > 0 || addresses.size > 0 };
}

/**
 * A label with a `*` in it, as a regex over the start of a cell's text.
 *
 * Every report type spells its own name differently -- "Summary Report Name",
 * "Regional Report Name", "Quarterly Report Name" -- and a config that has to
 * list each one is a config that silently misses the next report type someone
 * adds. `*Report Name` covers them, and nothing else in these headers.
 */
function globOf(label: string): RegExp {
  const body = label
    .split('*')
    .map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
    .join('[^:]*?');
  return new RegExp(`^${body}`);
}

/**
 * The addresses configured for one sheet, unqualified entries included.
 *
 * Layer 1 needs these as well as layer 2. It has no notion of an address rule
 * -- it aligns by header and key -- so without being told, it compares a cell
 * layer 2 has set aside and the same cell arrives in the report twice, once as
 * a defect and once as metadata.
 */
export function metadataAddressesFor(
  patterns: string[] | undefined,
  sheet: string,
): string[] {
  const { addresses } = parseMetadata(patterns);
  return [
    ...(addresses.get(canonSheet(sheet)) ?? []),
    ...(addresses.get('') ?? []),
  ];
}

const rowOf = (address: string): number => Number(address.replace(/^[A-Z$]+/i, ''));
const colOf = (address: string): number => colToNum(address.replace(/[^A-Z]/gi, '').toUpperCase());

/**
 * Which label a cell's text is, if any, and whether the value sits beside it.
 *
 * Two shapes occur in practice. The sheet writes the label in one cell and the
 * value in the next -- `A1 "Report ID"`, `B1 4542` -- in which case both are
 * metadata. Or it fuses them into one string, usually a formula:
 * `="Program name: " & <name>`. Then only that cell is.
 */
function labelHit(
  text: string,
  labels: MetadataLabel[],
  sheet: string,
): { rule: string; pair: boolean } | null {
  const t = norm(text);
  if (!t) return null;
  for (const { sheet: only, label, glob } of labels) {
    if (only && only !== sheet) continue;

    let width: number;
    if (glob) {
      const m = glob.exec(t);
      if (!m) continue;
      width = m[0].length;
    } else {
      if (!t.startsWith(label)) continue;
      width = label.length;
    }

    const rest = t.slice(width).replace(/^:/, '').trim();
    if (!rest) return { rule: label, pair: true };
    // "Report IDs by region" starts with "report id" but is not it; only a
    // separator may follow the label for the fused form to count.
    if (t[width] === ':') return { rule: label, pair: false };
  }
  return null;
}

/**
 * Every metadata cell on one sheet, from both sides of the comparison.
 *
 * `text` renders a cell's stored token as a person would see it; a formula
 * contributes its cached result when it has one, since that is what the label
 * reads as. Both files are scanned because a metadata block can gain a row.
 */
export function metadataOn(
  sheet: string,
  sides: (Map<string, string> | undefined)[],
  rules: MetadataRules,
  text: (token: string) => string,
): Map<string, string> {
  const out = new Map<string, string>();
  if (!rules.any) return out;

  const byAddress = rules.addresses.get(canonSheet(sheet));
  const anySheet = rules.addresses.get('');

  for (const cells of sides) {
    if (!cells) continue;
    for (const [address, token] of cells) {
      if (byAddress?.has(address)) { out.set(address, address); continue; }
      if (anySheet?.has(address)) { out.set(address, address); continue; }

      const hit = labelHit(text(token), rules.labels, canonSheet(sheet));
      if (!hit) continue;
      out.set(address, hit.rule);
      if (!hit.pair) continue;

      // The value beside the label. Only the next column: a label with three
      // blank cells after it and a number in the fifth is not a pair, it is a
      // table, and swallowing the row would hide real figures.
      const right = `${numToCol(colOf(address) + 1)}${rowOf(address)}`;
      if (sides.some((s) => s?.has(right))) out.set(right, hit.rule);
    }
  }

  return out;
}
