// @ts-check
/**
 * The workbook as a tree (sheets, with their tables, charts and PivotTables, plus named ranges), and the
 * Contents sheet built from it. Pure functions, so they can be tested without Excel.
 */

/**
 * @typedef {{ name: string, address: string }} Part
 * @typedef {{ name: string, visibility: "Visible" | "Hidden" | "VeryHidden", used: string, rows: number, cols: number,
 *   tables: Part[], charts: Part[], pivots: Part[] }} SheetInfo
 *   used: the used range's address without the sheet ("A1:H200"), or "" for an empty sheet
 * @typedef {{ name: string, formula: string, scope: string, visible: boolean }} NameInfo
 */

/** A sheet name as it must appear in a reference: quoted when it has spaces or punctuation. @param {string} name */
export function quoteSheet(name) {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** Text inside a formula string literal. @param {string} text */
const literal = (text) => `"${text.replace(/"/g, '""')}"`;

/**
 * A HYPERLINK formula that jumps to a place in this workbook.
 * @param {string} sheet @param {string} cell e.g. "A1" @param {string} text
 */
export function linkFormula(sheet, cell, text) {
  return `=HYPERLINK(${literal(`#${quoteSheet(sheet)}!${cell}`)},${literal(text)})`;
}

/** @param {number} n @param {string} word */
const plural = (n, word) => `${n.toLocaleString("en-US")} ${word}${n === 1 ? "" : "s"}`;

/** "A1:H200 · 200 rows × 8 columns", or "empty". @param {SheetInfo} s */
export function sizeLabel(s) {
  return s.used ? `${s.used} · ${plural(s.rows, "row")} × ${plural(s.cols, "column")}` : "empty";
}

/**
 * The Contents sheet: a header row, then one row per sheet with a link to it, its size and what's on it.
 * @param {SheetInfo[]} sheets @param {string} contentsName the sheet being written, which is left out
 * @returns {(string | number)[][]} formulas, ready for range.formulas
 */
export function contentsRows(sheets, contentsName) {
  const rows = /** @type {(string | number)[][]} */ ([["Sheet", "Used range", "Rows", "Columns", "Tables", "Charts", "PivotTables", "Notes"]]);
  for (const s of sheets) {
    if (s.name === contentsName) continue;
    const top = s.used ? s.used.split(":")[0] : "A1";
    rows.push([
      linkFormula(s.name, top, s.name),
      s.used || "(empty)",
      s.rows,
      s.cols,
      s.tables.map((t) => t.name).join(", "),
      s.charts.length,
      s.pivots.length,
      s.visibility === "Visible" ? "" : s.visibility === "Hidden" ? "Hidden sheet" : "Very hidden sheet",
    ]);
  }
  return rows;
}

/**
 * Keep only what matches the search (case-insensitive) in a sheet's name, or in the names of its tables,
 * charts and PivotTables. A sheet whose name matches keeps all of its parts.
 * @param {SheetInfo[]} sheets @param {NameInfo[]} names @param {string} query
 */
export function filterMap(sheets, names, query) {
  const q = query.trim().toLowerCase();
  if (!q) return { sheets, names };
  /** @param {string} s */
  const hit = (s) => s.toLowerCase().includes(q);
  /** @param {Part[]} parts */
  const keep = (parts) => parts.filter((p) => hit(p.name));
  return {
    sheets: sheets
      .map((s) => (hit(s.name) ? s : { ...s, tables: keep(s.tables), charts: keep(s.charts), pivots: keep(s.pivots) }))
      .filter((s) => hit(s.name) || s.tables.length || s.charts.length || s.pivots.length),
    names: names.filter((n) => hit(n.name) || hit(n.formula)),
  };
}

/** A name for the Contents sheet that doesn't clash with an existing sheet. @param {string[]} existing */
export function freeSheetName(existing, base = "Contents") {
  const taken = new Set(existing.map((n) => n.toLowerCase()));
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 2; ; i++) if (!taken.has(`${base} (${i})`.toLowerCase())) return `${base} (${i})`;
}
