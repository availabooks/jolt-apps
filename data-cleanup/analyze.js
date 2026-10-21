// @ts-check
/**
 * Finding problems in a block of spreadsheet data. Everything here is pure (no Office calls), so it can
 * be tested on plain arrays. The arrays are what Excel's `range.values` and `range.formulas` return:
 * one inner array per row.
 *
 * Only constant cells are ever changed. Cells with formulas are left alone, because a formula's result
 * is not the place to fix data.
 */

/**
 * @typedef {string | number | boolean} Cell
 * @typedef {{ row: number, col: number, from: Cell, to: Cell, numberFormat?: string }} CellFix
 *   row and col are 0-based within the range; numberFormat is set along with the new value
 * @typedef {{
 *   key: "spaces" | "textNumbers" | "blankRows" | "duplicates" | "mixedTypes",
 *   title: string,
 *   detail: string,
 *   count: number,
 *   cells?: CellFix[],
 *   rows?: number[],
 *   select?: { row: number, col: number }[],
 * }} Issue
 *   cells: value changes to make; rows: whole rows to delete (0-based, within the range);
 *   select: cells to point out when there is no automatic fix
 */

/** Spaces that don't belong: at either end, doubled, or the non-breaking kind pasted from web pages. */
const BAD_SPACE = /^[\s\u00a0]|[\s\u00a0]$|[\s\u00a0]{2,}|\u00a0/;

/** @param {string} s */
export const tidySpaces = (s) => s.replace(/[\s\u00a0]+/g, " ").trim();

/**
 * A number typed or imported as text: "1,234.50", "$99", "(42)", "-3.5", "12%", " 7 ".
 * Text with a leading zero ("00123", a ZIP code or ID) is left alone: the zero is probably meant.
 * @param {string} s
 * @returns {{ value: number, numberFormat?: string } | null}
 */
export function parseTextNumber(s) {
  let t = s.replace(/\u00a0/g, " ").trim();
  if (!t) return null;
  let sign = 1;
  if (/^\(.*\)$/.test(t)) {
    sign = -1;
    t = t.slice(1, -1).trim();
  }
  const symbol = /^[-+]?([$€£¥])/.exec(t)?.[1];
  t = t.replace(/^([-+]?)[$€£¥]\s*/, "$1");
  const percent = t.endsWith("%");
  if (percent) t = t.slice(0, -1).trim();
  if (!/^[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$|^[-+]?\.\d+$/.test(t)) return null;
  const digits = t.replace(/^[-+]/, "");
  if (/^0\d/.test(digits)) return null;
  const value = (sign * Number(t.replace(/,/g, ""))) / (percent ? 100 : 1);
  if (!Number.isFinite(value)) return null;
  // A number format is set too: cells holding numbers-as-text are often formatted as Text ("@").
  return { value, numberFormat: percent ? "0%" : symbol ? `${symbol}#,##0.00` : "General" };
}

/** @param {Cell} v */
const isBlank = (v) => v === "" || v === null || v === undefined;
/** @param {string | undefined} f */
const isFormula = (f) => typeof f === "string" && f.startsWith("=");

/** @param {Cell} v @returns {"number" | "text" | "boolean" | "blank"} */
const kind = (v) => (isBlank(v) ? "blank" : typeof v === "number" ? "number" : typeof v === "boolean" ? "boolean" : "text");

/**
 * Every issue found in the data.
 * @param {Cell[][]} values
 * @param {(string | Cell)[][]} formulas same shape; a string starting with "=" marks a formula cell
 * @param {{ headers: boolean }} options headers: the first row holds column names
 * @returns {Issue[]}
 */
export function findIssues(values, formulas, { headers }) {
  const first = headers ? 1 : 0;
  /** @param {number} r @param {number} c */
  const constant = (r, c) => !isFormula(/** @type {string} */ (formulas[r]?.[c]));
  const width = Math.max(0, ...values.map((row) => row.length));

  /** @type {CellFix[]} */
  const spaces = [];
  /** @type {CellFix[]} */
  const textNumbers = [];
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const v = values[r][c];
      if (typeof v !== "string" || !constant(r, c)) continue;
      if (BAD_SPACE.test(v)) spaces.push({ row: r, col: c, from: v, to: tidySpaces(v) });
      // Header names stay text even when they look like numbers ("2024").
      const n = r >= first ? parseTextNumber(v) : null;
      if (n) textNumbers.push({ row: r, col: c, from: v, to: n.value, numberFormat: n.numberFormat });
    }
  }

  /** @type {number[]} */
  const blankRows = [];
  /** @type {number[]} */
  const duplicates = [];
  const seen = new Set();
  for (let r = first; r < values.length; r++) {
    const row = values[r];
    if (row.every(isBlank)) {
      blankRows.push(r);
      continue;
    }
    // Rows with formulas are compared by their results, but never deleted as duplicates.
    const key = JSON.stringify(row.map((v) => (typeof v === "string" ? tidySpaces(v).toLowerCase() : v)));
    if (seen.has(key) && row.every((_, c) => constant(r, c))) duplicates.push(r);
    else seen.add(key);
  }

  /** @type {{ row: number, col: number }[]} */
  const mixed = [];
  const mixedColumns = [];
  for (let c = 0; c < width; c++) {
    const counts = { number: 0, text: 0, boolean: 0, blank: 0 };
    for (let r = first; r < values.length; r++) counts[kind(values[r][c])]++;
    // A column that is mostly numbers with a little text (or the reverse) usually has typos in it.
    // Text that will become a number after "Numbers stored as text" is fixed doesn't count.
    const pending = new Set(textNumbers.filter((f) => f.col === c).map((f) => f.row));
    const text = counts.text - pending.size;
    const numbers = counts.number + pending.size;
    if (!text || !numbers) continue;
    const minority = text < numbers ? "text" : "number";
    if (Math.min(text, numbers) / (text + numbers) > 0.2) continue;
    mixedColumns.push(c);
    for (let r = first; r < values.length; r++) {
      const k = kind(values[r][c]);
      const effective = k === "text" && pending.has(r) ? "number" : k;
      if (effective === minority) mixed.push({ row: r, col: c });
    }
  }

  /** @type {Issue[]} */
  const issues = [
    {
      key: "spaces",
      title: "Extra spaces",
      detail: "Spaces at the start or end, doubled spaces, or non-breaking spaces. They make lookups and matches fail.",
      count: spaces.length,
      cells: spaces,
    },
    {
      key: "textNumbers",
      title: "Numbers stored as text",
      detail: "These look like numbers but are text, so SUM and charts ignore them. Values with leading zeros (IDs, ZIP codes) are left alone.",
      count: textNumbers.length,
      cells: textNumbers,
    },
    {
      key: "blankRows",
      title: "Blank rows",
      detail: "Empty rows inside the data break sorting, filtering and PivotTables.",
      count: blankRows.length,
      rows: blankRows,
    },
    {
      key: "duplicates",
      title: "Duplicate rows",
      detail: "Rows that repeat an earlier row (ignoring case and extra spaces). The first copy is kept.",
      count: duplicates.length,
      rows: duplicates,
    },
    {
      key: "mixedTypes",
      title: "Odd values in a column",
      detail: `Cells whose type doesn't match the rest of their column (text in a number column, or the reverse) in ${mixedColumns.length} column${mixedColumns.length === 1 ? "" : "s"}. Check them by hand.`,
      count: mixed.length,
      select: mixed,
    },
  ];
  return issues;
}

/** A1-style column letters for a 0-based column index. @param {number} col */
export function columnLetters(col) {
  let s = "";
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/**
 * Apply cell fixes to copies of the range's values and number formats, for writing back in one go when
 * the range has no formulas.
 * @param {Cell[][]} values @param {string[][]} numberFormats @param {CellFix[]} fixes
 */
export function applyFixes(values, numberFormats, fixes) {
  const v = values.map((row) => [...row]);
  const f = numberFormats.map((row) => [...row]);
  for (const fix of fixes) {
    v[fix.row][fix.col] = fix.to;
    if (fix.numberFormat) f[fix.row][fix.col] = fix.numberFormat;
  }
  return { values: v, numberFormats: f };
}
