// @ts-check

/**
 * Pure functions for moving data between Excel and R, with no Office or webR calls, so they're easy to test.
 *
 * Excel → R: a block of cell values (and their number formats) becomes a list of typed columns.
 * R → Excel: the "blocks" that jolt.R makes from R objects become a grid of cell values plus the styling to apply.
 */

/**
 * @typedef {"number" | "date" | "datetime" | "logical" | "text"} ColumnType
 * @typedef {{ name: string, label: string, type: ColumnType, values: (number | string | boolean | null)[] }} Column
 *   name is a valid R name; label is the original header text. Dates are Excel serial numbers. null is NA.
 * @typedef {{ name: string, rows: number, columns: Column[] }} Frame
 *
 * @typedef {{ title?: string, header?: string[], cols?: unknown[][], fmt?: string[], text?: string[] }} Block
 *   One piece of output: a titled table (header + columns) or lines of text. fmt is an Excel number format per column ("" = General).
 * @typedef {{ name: string, sheet?: string, cell?: string, table: boolean, blocks: Block[] }} Output
 *   An object sent to Excel with excel_out() in R.
 * @typedef {{ row: number, col: number, rows: number, cols: number,
 *   kind: "title" | "header" | "end" | "text" | "format" | "table", format?: string }} Mark
 *   A rectangle of the grid that gets some styling. row/col are 0-based within the grid.
 *   "table" covers a whole table block (header and values), for autofitting or making an Excel Table.
 * @typedef {{ values: (string | number | boolean)[][], width: number, marks: Mark[] }} Grid
 */

/** Excel error values arrive as text, like "#N/A". They become NA in R. */
const ERROR_VALUE = /^#(N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)$/;

/** Words that can't be R names. */
const RESERVED = new Set(
  "if else repeat while function for next break TRUE FALSE NULL Inf NaN NA NA_integer_ NA_real_ NA_character_ NA_complex_ in".split(" "),
);

/** @param {unknown} v */
const isMissing = (v) => v === null || v === undefined || (typeof v === "string" && (v.trim() === "" || ERROR_VALUE.test(v.trim())));

/**
 * Text that is really a number: "1,234", "$5.00", "(12)", "45%". Returns the number, or null.
 * @param {string} s
 */
export function parseNumber(s) {
  let t = s.trim().replace(/[$€£¥,\s]/g, "");
  let negative = false;
  if (/^\(.*\)$/.test(t)) {
    negative = true;
    t = t.slice(1, -1);
  }
  const percent = t.endsWith("%");
  if (percent) t = t.slice(0, -1);
  if (!/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t)) return null;
  const n = Number(t) * (percent ? 0.01 : 1);
  return negative ? -n : n;
}

/**
 * Is this an Excel date or time number format? Quoted text and [colour]/[$-409] sections are ignored.
 * @param {string} format
 * @returns {"date" | "datetime" | null}
 */
export function dateKind(format) {
  if (!format || format === "General") return null;
  const f = format.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "").replace(/\\./g, "");
  const date = /[dy]/i.test(f) || /m{3,}/i.test(f);
  const time = /[hs]/i.test(f);
  if (date) return time ? "datetime" : "date";
  return null;
}

/**
 * Make a valid, readable R name: "Ad Spend ($)" → "Ad_Spend", "2024 Sales" → "X2024_Sales".
 * @param {string} text @param {string} fallback
 */
export function rName(text, fallback) {
  let n = String(text ?? "")
    .normalize("NFKD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Za-z0-9_.]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^[_.]+|[_.]+$/g, "");
  if (!n) n = fallback;
  if (/^[0-9]/.test(n) || /^\.[0-9]/.test(n)) n = "X" + n;
  if (RESERVED.has(n)) n += "_";
  return n;
}

/** Make names unique by adding _2, _3… @param {string[]} names */
export function uniqueNames(names) {
  const seen = new Set();
  return names.map((n) => {
    let name = n;
    for (let i = 2; seen.has(name.toLowerCase()); i++) name = `${n}_${i}`;
    seen.add(name.toLowerCase());
    return name;
  });
}

/** The first row is probably headers if it's all text and some row below has a number in it. @param {unknown[][]} values */
export function looksLikeHeaders(values) {
  if (values.length < 2) return false;
  const first = values[0].every((v) => typeof v === "string" && v.trim() !== "");
  return first && values.slice(1).some((row) => row.some((v) => typeof v === "number"));
}

/**
 * Turn a block of Excel values into typed columns.
 *
 * A column is numeric if every non-blank cell is a number or number-like text; it's a date if most of
 * those cells have a date format. Columns of TRUE/FALSE are logical. Anything else is text.
 * Blanks and error values (#N/A, #DIV/0!…) become NA.
 *
 * @param {unknown[][]} values @param {string[][] | null} formats the cells' number formats, same shape
 * @param {{ name: string, headers: boolean }} options
 * @returns {Frame}
 */
export function toFrame(values, formats, { name, headers }) {
  const width = Math.max(0, ...values.map((r) => r.length));
  const body = headers ? values.slice(1) : values;
  const bodyFormats = formats ? (headers ? formats.slice(1) : formats) : null;
  const labels = Array.from({ length: width }, (_, c) => (headers ? String(values[0]?.[c] ?? "").trim() : ""));
  const names = uniqueNames(labels.map((l, c) => rName(l, `V${c + 1}`)));
  const columns = names.map((n, c) => {
    const cells = body.map((row) => row[c]);
    const cellFormats = bodyFormats?.map((row) => row?.[c] ?? "General") ?? [];
    return { name: n, label: labels[c] || n, ...typeColumn(cells, cellFormats) };
  });
  return { name, rows: body.length, columns };
}

/**
 * @param {unknown[]} cells @param {string[]} formats
 * @returns {{ type: ColumnType, values: (number | string | boolean | null)[] }}
 */
function typeColumn(cells, formats) {
  const present = cells.filter((v) => !isMissing(v));
  if (present.length && present.every((v) => typeof v === "boolean")) {
    return { type: "logical", values: cells.map((v) => (isMissing(v) ? null : /** @type {boolean} */ (v))) };
  }
  const numbers = cells.map((v) => (isMissing(v) ? null : typeof v === "number" ? v : typeof v === "string" ? parseNumber(v) : null));
  const numeric = present.length > 0 && cells.every((v, i) => isMissing(v) || numbers[i] !== null);
  if (numeric) {
    const kinds = cells.map((v, i) => (typeof v === "number" ? dateKind(formats[i] ?? "") : null));
    const dated = kinds.filter(Boolean);
    const numberCount = cells.filter((v) => typeof v === "number").length;
    if (numberCount && dated.length * 2 >= numberCount) {
      const hasTime = dated.includes("datetime") || numbers.some((n) => n !== null && n % 1 !== 0);
      return { type: hasTime ? "datetime" : "date", values: numbers };
    }
    return { type: "number", values: numbers };
  }
  return { type: "text", values: cells.map((v) => (isMissing(v) ? null : typeof v === "string" ? v.trim() : String(v))) };
}

/** Excel's day 0 is 1899-12-30 (allowing for its 1900 leap-year bug). @param {number} serial */
export function serialToISO(serial, withTime = false) {
  const ms = Math.round((serial - 25569) * 86400000);
  const iso = new Date(ms).toISOString();
  return withTime ? iso.slice(0, 19).replace("T", " ") : iso.slice(0, 10);
}

/**
 * Lay out output blocks as cells: each block's title, its header row and its columns, with a blank row
 * between blocks. Text blocks are one line per row in the first column.
 * @param {Block[]} blocks
 * @returns {Grid}
 */
export function blocksToGrid(blocks) {
  /** @type {(string | number | boolean)[][]} */
  const rows = [];
  /** @type {Mark[]} */
  const marks = [];
  let width = 1;
  blocks.forEach((b, i) => {
    if (i > 0) rows.push([]);
    if (b.title) {
      marks.push({ row: rows.length, col: 0, rows: 1, cols: 1, kind: "title" });
      rows.push([b.title]);
    }
    if (b.text) {
      marks.push({ row: rows.length, col: 0, rows: b.text.length, cols: 1, kind: "text" });
      for (const line of b.text) rows.push([cellValue(line)]);
      return;
    }
    const cols = b.cols ?? [];
    const n = Math.max(0, ...cols.map((c) => c.length));
    const w = Math.max(cols.length, b.header?.length ?? 0);
    width = Math.max(width, w);
    const start = rows.length;
    if (b.header) {
      marks.push({ row: rows.length, col: 0, rows: 1, cols: w, kind: "header" });
      rows.push(b.header.map(cellValue));
    }
    const top = rows.length;
    for (let r = 0; r < n; r++) rows.push(cols.map((c) => cellValue(c[r])));
    if (n) marks.push({ row: top + n - 1, col: 0, rows: 1, cols: w, kind: "end" });
    if (w && rows.length > start) marks.push({ row: start, col: 0, rows: rows.length - start, cols: w, kind: "table" });
    b.fmt?.forEach((format, c) => {
      if (format && n) marks.push({ row: top, col: c, rows: n, cols: 1, kind: "format", format });
    });
  });
  return { values: rows.map((r) => [...r, ...Array(width - r.length).fill("")].slice(0, width)), width, marks };
}

/**
 * A value Excel will store as-is. Text that starts with "=" would become a formula, so it gets a leading apostrophe.
 * @param {unknown} v
 * @returns {string | number | boolean}
 */
function cellValue(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return Number.isFinite(v) ? v : String(v);
  if (typeof v === "boolean") return v;
  const s = String(v);
  return /^[=+\-@]/.test(s) && !/^[-+]?\d/.test(s) ? "'" + s : s;
}

/**
 * Tidy the JSON that jolt.R's .jolt_json() writes. R has no scalars, so every value arrives as an array:
 * ["Title"] becomes "Title", and in a column of mixed values [[1.5], ["text"]] becomes [1.5, "text"].
 * @param {any} raw
 * @returns {Output[]}
 */
export function normalizeOutputs(raw) {
  /** @param {any} v */
  const one = (v) => (Array.isArray(v) ? (v[0] ?? null) : v);
  /** @param {any} v */
  const list = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]);
  return list(raw).map((o) => ({
    name: one(o.name) ?? "",
    sheet: one(o.sheet) ?? undefined,
    cell: one(o.cell) ?? undefined,
    table: !!one(o.table),
    blocks: list(o.blocks).map((b) => {
      /** @type {Block} */
      const block = {};
      const title = one(b.title);
      if (title) block.title = String(title);
      if (b.text) block.text = list(b.text).map(String);
      if (b.header) block.header = list(b.header).map((h) => String(h ?? ""));
      if (b.cols) block.cols = list(b.cols).map((c) => list(c).map(one));
      if (b.fmt) block.fmt = list(b.fmt).map((f) => String(f ?? ""));
      return block;
    }),
  }));
}

/** "Sheet1!A1:D20" → ["Sheet1", "A1:D20"], unquoting names like 'My Sheet'. @param {string} address */
export function splitAddress(address) {
  const i = address.lastIndexOf("!");
  return [address.slice(0, i).replace(/^'|'$/g, "").replace(/''/g, "'"), address.slice(i + 1)];
}

/** A sheet name Excel accepts: no []:*?/\ and at most 31 characters. @param {string} name */
export function sheetName(name) {
  const n = name.replace(/[[\]:*?/\\]/g, " ").replace(/^'+|'+$/g, "").trim().slice(0, 31);
  return n || "R Output";
}

/** "R Output", then "R Output 2"… whichever isn't taken. @param {string} base @param {string[]} taken */
export function freeSheetName(base, taken) {
  const used = new Set(taken.map((t) => t.toLowerCase()));
  const b = sheetName(base);
  if (!used.has(b.toLowerCase())) return b;
  for (let i = 2; ; i++) {
    const suffix = ` ${i}`;
    const n = b.slice(0, 31 - suffix.length) + suffix;
    if (!used.has(n.toLowerCase())) return n;
  }
}

// ---- Analysis files --------------------------------------------------------------------------------------

/**
 * @typedef {{ name: string, kind: "table" | "range", table?: string, sheet?: string, address?: string, headers: boolean }} SavedSource
 *   A data source as written to a file: no id, since ids are only unique within one workbook.
 * @typedef {{ name: string, description: string, question: string, includeSample: boolean, code: string, explanation: string, sources: SavedSource[] }} SavedAnalysis
 */

/** Marks a file as one of ours, so other JSON files are rejected with a clear message. */
const ANALYSIS_FORMAT = "r-for-excel-analysis";

/** An analysis as pretty-printed JSON, for downloading or copying. @param {SavedAnalysis} a */
export function serializeAnalysis(a) {
  const sources = a.sources.map(({ name, kind, table, sheet, address, headers }) =>
    kind === "table" ? { name, kind, table, headers } : { name, kind, sheet, address, headers },
  );
  return JSON.stringify(
    { format: ANALYSIS_FORMAT, version: 1, name: a.name, description: a.description, question: a.question, includeSample: a.includeSample, sources, code: a.code, explanation: a.explanation },
    null,
    2,
  );
}

/** Starts the comment that carries an analysis inside an HTML report. */
const COMMENT_START = `<!-- ${ANALYSIS_FORMAT}`;

/**
 * An analysis file as an HTML comment, so a downloaded report can be imported to rebuild the analysis.
 * In JSON, "--" can only appear inside strings, where "--" means the same thing and can't end the comment.
 * @param {string} json from serializeAnalysis
 */
export const analysisComment = (json) => `${COMMENT_START}\n${json.replace(/--/g, "-\\u002d")}\n-->`;

/**
 * Read an analysis file, or the analysis inside an HTML report, checking every field. Sources that make no
 * sense are dropped, and their names are made valid and unique R names. Throws an Error with a message for
 * the user if it isn't an analysis.
 * @param {string} text
 * @returns {SavedAnalysis}
 */
export function parseAnalysis(text) {
  const start = text.indexOf(COMMENT_START);
  const end = start < 0 ? -1 : text.indexOf("-->", start);
  if (end > 0) text = text.slice(start + COMMENT_START.length, end);
  else if (/^\s*</.test(text)) throw new Error("That HTML file isn't an R for Excel report, or it was made before reports included their analysis.");
  /** @type {any} */
  let raw;
  try {
    raw = JSON.parse(text);
  } catch {
    throw new Error("That isn't an analysis file: it isn't valid JSON.");
  }
  if (!raw || typeof raw !== "object" || raw.format !== ANALYSIS_FORMAT) throw new Error("That isn't an R for Excel analysis file.");
  if (raw.version !== 1) throw new Error(`This analysis file is version ${raw.version}. This version of R for Excel reads version 1.`);
  /** @param {unknown} v */
  const str = (v) => (typeof v === "string" ? v : "");
  /** @type {SavedSource[]} */
  const sources = [];
  for (const s of Array.isArray(raw.sources) ? raw.sources : []) {
    if (!s || typeof s !== "object") continue;
    const name = rName(str(s.name), "data");
    if (s.kind === "table" && str(s.table)) sources.push({ name, kind: "table", table: str(s.table), headers: true });
    else if (s.kind === "range" && str(s.sheet) && str(s.address)) sources.push({ name, kind: "range", sheet: str(s.sheet), address: str(s.address), headers: s.headers !== false });
  }
  const names = uniqueNames(sources.map((s) => s.name));
  sources.forEach((s, i) => (s.name = names[i]));
  return {
    name: str(raw.name).trim() || "Imported analysis",
    description: str(raw.description),
    question: str(raw.question),
    includeSample: raw.includeSample !== false,
    code: str(raw.code),
    explanation: str(raw.explanation),
    sources,
  };
}

/** A name that's safe in a file name on any system: "Sales / 2024?" → "Sales 2024". @param {string} name */
export function safeFileName(name) {
  return name.replace(/[<>:"/\\|?*\u0000-\u001f]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 80) || "analysis";
}

/** "Sales / 2024?" → "Sales 2024.r-analysis.json". @param {string} name */
export const analysisFileName = (name) => `${safeFileName(name)}.r-analysis.json`;
