/**
 * Checks a template against the rules it carries about itself.
 *
 * These sheets ship data validations: "Include" is Yes or No, "Treaty
 * Participation" is a decimal between 0 and 1, "Treaty Premium" is at least 0.
 * They are a contract with whoever edits the file -- Excel refuses anything
 * else on the way back in -- so a wrong rule breaks the upload half of the
 * round trip while the download looks perfect.
 *
 * Two things are checked, and neither needs a source of any kind. The template
 * states both sides of the argument itself:
 *
 *   1. Every value the template wrote satisfies the rule it wrote beside it. A
 *      sheet that ships a figure its own rule rejects is a sheet nobody can
 *      edit at that cell.
 *   2. A rule written per row refers to its own row. A formula copied down
 *      without being advanced validates somebody else's cell, which passes for
 *      as long as the two happen to agree.
 *
 * `custom` rules are reported as unevaluated rather than assumed good: this is
 * not an Excel formula engine, and claiming to have checked something that was
 * skipped is the failure this whole tool exists to avoid. Their row references
 * are still checked, which is what (2) is for.
 */

/** openpyxl writes a missing bound as the string "None". */
const bound = (v) => {
  if (v === null || v === undefined) return null;
  if (typeof v === 'string' && (v === 'None' || v.trim() === '')) return null;
  return v;
};

const asNumber = (v) => {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  if (v instanceof Date) return v.getTime();
  const n = Number(String(v).trim());
  return Number.isNaN(n) ? null : n;
};

/** `"Yes,No"` -> ['Yes', 'No'] */
const listItems = (f) => {
  const s = String(f ?? '').trim().replace(/^"(.*)"$/s, '$1');
  return s.split(',').map((x) => x.trim()).filter((x) => x !== '');
};

const COMPARE = {
  between: (v, a, b) => a !== null && b !== null && v >= Math.min(a, b) && v <= Math.max(a, b),
  notBetween: (v, a, b) => a === null || b === null || v < Math.min(a, b) || v > Math.max(a, b),
  equal: (v, a) => v === a,
  notEqual: (v, a) => v !== a,
  greaterThan: (v, a) => v > a,
  lessThan: (v, a) => v < a,
  greaterThanOrEqual: (v, a) => v >= a,
  lessThanOrEqual: (v, a) => v <= a,
};

/**
 * Cell references a formula makes, once each.
 *
 * Deduplicated because a rule naming the same cell twice --
 * OR(Q5="Unlimited", ISNUMBER(Q5)) -- is one mistake, not two.
 */
function references(formula) {
  const seen = new Map();
  for (const m of String(formula ?? '').matchAll(/\$?([A-Z]{1,3})\$?(\d+)/g)) {
    if (!seen.has(m[0])) seen.set(m[0], { column: m[1], row: Number(m[2]), text: m[0] });
  }
  return [...seen.values()];
}

/** The resolved value of a cell, formulas included. */
function valueOf(cell) {
  let v = cell.value;
  if (v && typeof v === 'object') {
    if (v.result !== undefined) v = v.result;
    else if (v.richText) v = v.richText.map((t) => t.text).join('');
    else if (v.text !== undefined) v = v.text;
  }
  return v;
}

/**
 * @param ws      the worksheet
 * @param rows    1-based data row numbers
 * @param columns Map of column name -> column number
 */
export function checkValidations(ws, rows, columns) {
  const findings = [];
  let checked = 0;
  let custom = 0;

  const nameAt = new Map([...columns].map(([name, n]) => [n, name]));

  for (const row of rows) {
    for (const [n, name] of nameAt) {
      const cell = ws.getRow(row).getCell(n);
      const dv = cell.dataValidation;
      if (!dv) continue;

      const [f1, f2] = (dv.formulae ?? []).map(bound);

      // (2) A per-row rule that names a different row is validating somebody
      // else's cell. Excel does not complain; it simply checks the wrong thing.
      for (const ref of references(f1)) {
        if (ref.row !== row) {
          findings.push({
            row,
            column: name,
            address: cell.address,
            value: valueOf(cell),
            problem: `its rule refers to ${ref.text}, which is row ${ref.row}, not row ${row}`,
          });
        }
      }

      if (dv.type === 'custom') { custom++; continue; }

      checked++;
      const value = valueOf(cell);
      const blank = value === null || value === undefined || value === '';
      if (blank) {
        if (dv.allowBlank === false) {
          findings.push({ row, column: name, address: cell.address, problem: 'empty, and the rule does not allow blank' });
        }
        continue;
      }

      if (dv.type === 'list') {
        const items = listItems(f1);
        if (items.length && !items.includes(String(value).trim())) {
          findings.push({
            row, column: name, address: cell.address, value,
            problem: `is not one of ${items.map((i) => `"${i}"`).join(', ')}`,
          });
        }
        continue;
      }

      if (dv.type === 'decimal' || dv.type === 'whole' || dv.type === 'textLength') {
        const v = dv.type === 'textLength' ? String(value).length : asNumber(value);
        if (v === null) {
          findings.push({ row, column: name, address: cell.address, value, problem: 'is not a number, and the rule wants one' });
          continue;
        }
        if (dv.type === 'whole' && !Number.isInteger(v)) {
          findings.push({ row, column: name, address: cell.address, value, problem: 'is not a whole number' });
          continue;
        }
        const op = COMPARE[dv.operator ?? 'between'];
        const a = asNumber(f1);
        const b = asNumber(f2);
        if (op && a !== null && !op(v, a, b)) {
          const said = dv.operator === 'between' || !dv.operator
            ? `between ${a} and ${b}`
            : `${String(dv.operator).replace(/([A-Z])/g, ' $1').toLowerCase()} ${a}`;
          findings.push({ row, column: name, address: cell.address, value, problem: `is not ${said}` });
        }
      }
    }
  }

  return { ok: findings.length === 0, checked, custom, findings };
}
