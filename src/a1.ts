/**
 * Formula reference rewriting.
 *
 * The comparison hinges on this: two formulas that compute the same thing must
 * normalise to the same string even when columns have moved between releases.
 *
 * A naive regex over cell-looking substrings is wrong in ways that matter --
 * it rewrites the "G10" inside LOG10(), mangles sheet-qualified names, and
 * corrupts text inside string literals. So this is a small scanner instead.
 */

export function colToNum(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}

const MAX_COL = 16384; // XFD
const MAX_ROW = 1048576;

interface ParsedRef {
  colAbs: boolean;
  col: number;
  rowAbs: boolean;
  row: number;
}

/** Parses a bare A1 reference. Returns null if the text is not a valid ref. */
function parseRef(text: string): ParsedRef | null {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/.exec(text);
  if (!m) return null;
  const col = colToNum(m[2]!.toUpperCase());
  const row = Number(m[4]);
  if (col < 1 || col > MAX_COL || row < 1 || row > MAX_ROW) return null;
  return { colAbs: m[1] === '$', col, rowAbs: m[3] === '$', row };
}

const isIdentStart = (c: string) => /[A-Za-z_\\$]/.test(c);
const isIdentPart = (c: string) => /[A-Za-z0-9_.$]/.test(c);

type Rewriter = (ref: ParsedRef, qualified: boolean) => string;

/**
 * Walks the formula, rewriting only genuine unqualified cell references.
 * String literals, function names, defined names, structured table refs and
 * sheet-qualified references are passed through untouched.
 */
function rewrite(formula: string, fn: Rewriter): string {
  let out = '';
  let i = 0;
  const n = formula.length;

  while (i < n) {
    const c = formula[i]!;

    // "text" -- doubled quotes escape
    if (c === '"') {
      let j = i + 1;
      while (j < n) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') j += 2;
          else { j++; break; }
        } else j++;
      }
      out += formula.slice(i, j);
      i = j;
      continue;
    }

    // 'Sheet name'!  -- quoted sheet prefix
    if (c === "'") {
      let j = i + 1;
      while (j < n) {
        if (formula[j] === "'") {
          if (formula[j + 1] === "'") j += 2;
          else { j++; break; }
        } else j++;
      }
      out += formula.slice(i, j);
      i = j;
      // fall through so the ref after "!" is seen as qualified
      if (formula[i] === '!') { out += '!'; i++; out += consumeQualified(); }
      continue;
    }

    // #REF!, #N/A and friends -- never rewrite
    if (c === '#') {
      const m = /^#[A-Za-z0-9\/!?]+/.exec(formula.slice(i));
      if (m) { out += m[0]; i += m[0].length; continue; }
    }

    if (isIdentStart(c)) {
      let j = i;
      while (j < n && isIdentPart(formula[j]!)) j++;
      const word = formula.slice(i, j);
      const after = formula[j];

      // Table1[Col] -- structured reference, skip the whole bracket group
      if (after === '[') {
        let depth = 0, k = j;
        while (k < n) {
          if (formula[k] === '[') depth++;
          else if (formula[k] === ']') { depth--; if (depth === 0) { k++; break; } }
          k++;
        }
        out += formula.slice(i, k);
        i = k;
        continue;
      }

      // Sheet1!A1 -- unquoted sheet prefix
      if (after === '!') {
        out += word + '!';
        i = j + 1;
        out += consumeQualified();
        continue;
      }

      // SUM( -- function call, not a reference
      if (after === '(') { out += word; i = j; continue; }

      const ref = parseRef(word);
      out += ref ? fn(ref, false) : word;
      i = j;
      continue;
    }

    out += c;
    i++;
  }

  return out;

  /** Reads the reference immediately after a sheet prefix, leaving it as-is. */
  function consumeQualified(): string {
    let j = i;
    while (j < n && isIdentPart(formula[j]!)) j++;
    const word = formula.slice(i, j);
    const ref = parseRef(word);
    const text = ref ? fn(ref, true) : word;
    i = j;
    return text;
  }
}

/** A1 -> R1C1, relative to the cell the formula lives in. */
export function toR1C1(formula: string, row: number, col: number): string {
  return rewrite(formula, (r, qualified) => {
    if (qualified) return refToA1(r);
    const rp = r.rowAbs ? `R${r.row}` : r.row === row ? 'R' : `R[${r.row - row}]`;
    const cp = r.colAbs ? `C${r.col}` : r.col === col ? 'C' : `C[${r.col - col}]`;
    return rp + cp;
  });
}

/**
 * A1 -> header-resolved. Column letters become the header name of the column
 * they point at, so the result is invariant under any column insertion,
 * deletion or reorder.
 *
 * Cross-sheet references are left in A1 form: this sheet's headers say nothing
 * about another sheet's layout.
 */
export function toHeaderRef(
  formula: string,
  row: number,
  col: number,
  headerByCol: ReadonlyMap<number, string>,
): string {
  return rewrite(formula, (r, qualified) => {
    if (qualified) return refToA1(r);
    const name = headerByCol.get(r.col) ?? `col${r.col}`;
    const rp = r.rowAbs
      ? `@R${r.row}`
      : r.row === row
        ? '@row'
        : `@row${r.row - row > 0 ? '+' : ''}${r.row - row}`;
    return `[${name}]${rp}`;
  });
}

function refToA1(r: ParsedRef): string {
  return `${r.colAbs ? '$' : ''}${numToCol(r.col)}${r.rowAbs ? '$' : ''}${r.row}`;
}
