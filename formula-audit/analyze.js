// @ts-check
/**
 * Checks for the formula problems that cause most spreadsheet mistakes. Pure functions over the arrays
 * Excel returns for a sheet's used range, so they can be tested without Excel.
 */

/**
 * @typedef {"error" | "inconsistent" | "hardcoded" | "external" | "volatile"} Category
 * @typedef {{ category: Category, sheet: string, address: string, detail: string, formula: string }} Finding
 * @typedef {{ name: string, top: number, left: number, formulas: unknown[][], formulasR1C1: unknown[][],
 *   values: unknown[][], valueTypes: string[][] }} SheetData
 *   top/left: 0-based row and column of the used range's first cell
 */

export const CATEGORIES = /** @type {Record<Category, { title: string, why: string }>} */ ({
  error: { title: "Errors", why: "Cells showing #REF!, #N/A, #DIV/0! and other errors, which spread to every formula that uses them." },
  inconsistent: {
    title: "Inconsistent formulas",
    why: "A formula that differs from the ones on both sides of it (above and below, or left and right): often a copy-paste slip or a manual override.",
  },
  hardcoded: { title: "Typed-in numbers", why: "Numbers typed into formulas, such as =B2*1.07. When the rate changes, every copy has to be found and edited. Put them in a cell instead." },
  external: { title: "Links to other files", why: "Formulas that read from other workbooks. They break when that file moves or someone else opens this one." },
  volatile: {
    title: "Volatile functions",
    why: "INDIRECT, OFFSET, NOW, TODAY, RAND and similar recalculate on every change, which slows big workbooks. INDIRECT also hides what a formula depends on.",
  },
});

/** Numbers that are almost always fine in a formula: 0, 1, 100 (percentages), and 2 (halving, squaring). */
const HARMLESS = new Set([0, 1, 2, 100]);
const VOLATILE = /\b(INDIRECT|OFFSET|NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|CELL|INFO)\s*\(/gi;
/** [Book.xlsx]Sheet1!A1, '[Book.xlsx]My Sheet'!A1, 'C:\dir\[Book.xlsx]Sheet'!A1, and the stored form [1]Sheet1!A1. */
const EXTERNAL = /\[[^\]]+\][^[\]!(),;=+\-*/&^<>"]*'?!/;

/** A formula with its string literals blanked out, so text like "Q4 2024" isn't mistaken for code. @param {string} f */
export const withoutStrings = (f) => f.replace(/"(?:[^"]|"")*"/g, '""');

/**
 * Numbers typed into a formula, ignoring cell references (A1, $B$2, 1:1), function names (LOG10, ATAN2),
 * structured references (Table1[Col]) and the harmless constants above.
 * @param {string} formula A1-style
 * @returns {number[]}
 */
export function typedNumbers(formula) {
  const code = withoutStrings(formula)
    .replace(/'[^']*'!/g, " ") // quoted sheet names
    .replace(/\[[^\]]*\]/g, " ") // structured references and external workbook names
    .replace(/\$?\b\d+:\$?\d+\b/g, " ") // whole-row ranges like 1:1
    .replace(/\$/g, "")
    .replace(/\b[A-Za-z_][A-Za-z0-9_.]*/g, " "); // names, functions, and A1 references
  const found = [];
  for (const m of code.matchAll(/(?<![\w.])\d+(\.\d+)?([eE][-+]?\d+)?%?/g)) {
    const text = m[0];
    const n = Number(text.replace("%", "")) / (text.endsWith("%") ? 100 : 1);
    if (!HARMLESS.has(n)) found.push(n);
  }
  return found;
}

/** @param {string} formula */
export const isExternal = (formula) => EXTERNAL.test(withoutStrings(formula));

/** @param {string} formula */
export const volatileFunctions = (formula) => [...new Set([...withoutStrings(formula).matchAll(VOLATILE)].map((m) => m[1].toUpperCase()))];

/** A1-style column letters for a 0-based column index. @param {number} col */
export function columnLetters(col) {
  let s = "";
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** @param {unknown} f @returns {f is string} */
const isFormula = (f) => typeof f === "string" && f.startsWith("=");

/**
 * Every finding on one sheet, or on a slice of its rows. A sheet is read in slices when it is large;
 * each slice includes one extra row above and below, so formulas at its edges have neighbours to compare
 * with, and `rows` says which rows of the slice to report on.
 * @param {SheetData} s
 * @param {{ from: number, to: number }} [rows] 0-based rows of `s` to report, `to` exclusive
 * @returns {Finding[]}
 */
export function auditSheet(s, rows = { from: 0, to: s.formulas.length }) {
  /** @type {Finding[]} */
  const out = [];
  /** @param {number} r @param {number} c */
  const addr = (r, c) => `${columnLetters(s.left + c)}${s.top + r + 1}`;
  /** @param {number} r @param {number} c */
  const r1c1 = (r, c) => {
    const f = s.formulasR1C1[r]?.[c];
    return isFormula(f) ? f : null;
  };
  for (let r = rows.from; r < rows.to; r++) {
    for (let c = 0; c < s.formulas[r].length; c++) {
      const formula = s.formulas[r][c];
      const at = { sheet: s.name, address: addr(r, c), formula: isFormula(formula) ? formula : "" };
      if (s.valueTypes[r]?.[c] === "Error") {
        out.push({ ...at, category: "error", detail: `${s.values[r][c]}${isFormula(formula) ? "" : " (typed in, not calculated)"}` });
      }
      if (!isFormula(formula)) continue;

      // Compared in R1C1 form, where copies of one formula are identical text (=R[-1]C+1 everywhere).
      const own = r1c1(r, c);
      const up = r1c1(r - 1, c);
      const down = r1c1(r + 1, c);
      const left = r1c1(r, c - 1);
      const right = r1c1(r, c + 1);
      if ((up && up === down && own !== up) || (left && left === right && own !== left)) {
        const neighbour = up && up === down ? `above and below (${addr(r - 1, c)}, ${addr(r + 1, c)})` : `to the left and right (${addr(r, c - 1)}, ${addr(r, c + 1)})`;
        out.push({ ...at, category: "inconsistent", detail: `Differs from the formulas ${neighbour}` });
      }

      const numbers = typedNumbers(formula);
      if (numbers.length) out.push({ ...at, category: "hardcoded", detail: `Typed-in ${numbers.length === 1 ? "number" : "numbers"}: ${numbers.slice(0, 5).join(", ")}` });
      if (isExternal(formula)) out.push({ ...at, category: "external", detail: "Reads from another workbook" });
      const vol = volatileFunctions(formula);
      if (vol.length) out.push({ ...at, category: "volatile", detail: vol.join(", ") });
    }
  }
  return out;
}

/**
 * Rows for the report sheet. Formulas are written as text (a leading apostrophe), so the report shows
 * them rather than running them.
 * @param {Finding[]} findings
 */
export function reportRows(findings) {
  /** @param {string} s */
  const q = (s) => s.replace(/"/g, '""');
  /** @param {string} sheet */
  const ref = (sheet) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`);
  return [
    ["Category", "Sheet", "Cell", "Detail", "Formula"],
    ...findings.map((f) => [
      CATEGORIES[f.category].title,
      f.sheet,
      `=HYPERLINK("#${q(ref(f.sheet))}!${f.address}","${f.address}")`,
      // "#REF!" or "=…" typed into a cell would become an error value or a formula.
      /^[=#+\-@]/.test(f.detail) ? `'${f.detail}` : f.detail,
      f.formula ? `'${f.formula}` : "",
    ]),
  ];
}
