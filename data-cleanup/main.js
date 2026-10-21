// @ts-check
import van from "van";
import { columnLetters, findIssues } from "./analyze.js";

/**
 * Data Cleanup: select some data (or click one cell inside it) and get a list of what's wrong with it,
 * each with a one-click fix and an Undo.
 *
 * How it talks to Excel: every Office call goes through `jolt.run(async (context) => …)`, which is
 * Excel.run. Values are read with `load` + `context.sync()`, and writes are batched into one sync.
 */

const { main, header, h1, p, button, label, input, ul, li, div, span, strong, small } = van.tags;

/** Larger blocks are slow to read and fix in one go. */
const MAX_CELLS = 300_000;

/**
 * @typedef {import("./analyze.js").Issue} Issue
 * @typedef {{ sheet: string, address: string, top: number, left: number, rows: number, cols: number, hasFormulas: boolean }} Target
 *   address is without the sheet name; top/left are the 0-based row and column of its first cell
 * @typedef {{ label: string, sheet: string, cells?: { address: string, formula: unknown, numberFormat: string }[],
 *   block?: { address: string, formulasR1C1: unknown[][], numberFormat: string[][] } }} Undo
 */

const target = van.state(/** @type {Target | null} */ (null));
const issues = van.state(/** @type {Issue[]} */ ([]));
const headers = van.state(true);
const busy = van.state("");
const undo = van.state(/** @type {Undo | null} */ (null));
const stats = van.state(/** @type {{ cells: number, rows: number }} */ ({ cells: 0, rows: 0 }));

/** The block to clean: the selection, or the data around the active cell when only one cell is selected. */
async function selectedBlock() {
  /** @param {(sel: Excel.Range) => Excel.Range} pick */
  const read = (pick) =>
    jolt.run(async (context) => {
      const range = pick(context.workbook.getSelectedRange());
      range.load("address,rowIndex,columnIndex,rowCount,columnCount,cellCount,isNullObject");
      await context.sync();
      if (range.isNullObject) throw new Error("The selection has no data in it.");
      return { address: range.address, top: range.rowIndex, left: range.columnIndex, rows: range.rowCount, cols: range.columnCount, cells: range.cellCount };
    });
  // A whole selected column or row is mostly empty: keep only the part that has data.
  const selection = await read((sel) => sel.getIntersectionOrNullObject(sel.worksheet.getUsedRange()));
  if (selection.cells > 1) return selection;
  try {
    // One cell: use the "current region", the block of filled cells around it (Ctrl+A in Excel).
    return await read((sel) => sel.getSurroundingRegion());
  } catch {
    // Older Excel without getSurroundingRegion: use everything on the sheet.
    return read((sel) => sel.worksheet.getUsedRange());
  }
}

/** @param {string} sheet @param {string} address */
async function readBlock(sheet, address) {
  return jolt.run(async (context) => {
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    range.load("values,formulas");
    await context.sync();
    return { values: range.values, formulas: range.formulas };
  });
}

/** Read the block and look for problems. Guesses whether the first row is headers the first time. */
async function analyze({ guessHeaders = true } = {}) {
  await work("Reading the data…", async () => {
    const block = await selectedBlock();
    if (block.cells > MAX_CELLS) throw new Error(`That's ${block.cells.toLocaleString()} cells. Select up to ${MAX_CELLS.toLocaleString()} at a time.`);
    const [sheet, address] = splitAddress(block.address);
    const { values, formulas } = await readBlock(sheet, address);
    if (guessHeaders) headers.val = looksLikeHeaders(values);
    target.val = { sheet, address, top: block.top, left: block.left, rows: block.rows, cols: block.cols, hasFormulas: formulas.some((r) => r.some(isFormula)) };
    issues.val = findIssues(values, formulas, { headers: headers.val });
  });
}

/**
 * Fix one kind of issue. The data is read again first, so the fix matches what's in the sheet now.
 * @param {Issue["key"]} key
 */
async function fix(key) {
  const t = target.val;
  if (!t) return;
  await work("Fixing…", async () => {
    const { values, formulas } = await readBlock(t.sheet, t.address);
    const issue = findIssues(values, formulas, { headers: headers.val }).find((i) => i.key === key);
    if (!issue?.count) return;
    if (issue.cells) await fixCells(t, issue);
    else if (issue.rows) await removeRows(t, issue);
    await analyze({ guessHeaders: false });
    jolt.notify(`${issue.title}: fixed ${issue.count.toLocaleString()}.`, "success");
  });
}

/** Change individual cells, keeping their old contents for Undo. @param {Target} t @param {Issue} issue */
async function fixCells(t, issue) {
  const fixes = issue.cells ?? [];
  await jolt.run(async (context) => {
    const block = context.workbook.worksheets.getItem(t.sheet).getRange(t.address);
    const cells = fixes.map((f) => block.getCell(f.row, f.col));
    for (const c of cells) c.load("address,formulas,numberFormat");
    await context.sync();
    undo.val = {
      label: issue.title,
      sheet: t.sheet,
      cells: cells.map((c) => ({ address: splitAddress(c.address)[1], formula: c.formulas[0][0], numberFormat: c.numberFormat[0][0] })),
    };
    fixes.forEach((f, i) => {
      cells[i].values = [[f.to]];
      if (f.numberFormat) cells[i].numberFormat = [[f.numberFormat]];
    });
    await context.sync();
  });
  await count({ cells: fixes.length });
}

/**
 * Remove rows by moving the rows below them up, inside the block only, so nothing outside it moves.
 * Formulas are moved as R1C1 text, which keeps references relative to their own row the way Excel does.
 * Cell formatting such as fill colour stays where it is.
 * @param {Target} t @param {Issue} issue
 */
async function removeRows(t, issue) {
  const remove = new Set(issue.rows);
  const firstRow = Math.min(...remove);
  await jolt.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(t.sheet);
    const block = sheet.getRange(t.address);
    // Only the part from the first removed row down changes.
    const changed = block.getRow(firstRow).getResizedRange(t.rows - firstRow - 1, 0);
    changed.load("address,formulasR1C1,numberFormat");
    await context.sync();
    undo.val = {
      label: issue.title,
      sheet: t.sheet,
      block: { address: splitAddress(changed.address)[1], formulasR1C1: changed.formulasR1C1, numberFormat: changed.numberFormat },
    };
    const keep = changed.formulasR1C1.map((row, i) => ({ row, format: changed.numberFormat[i], i: i + firstRow })).filter((r) => !remove.has(r.i));
    const empty = () => Array.from({ length: t.cols }, () => "");
    const general = () => Array.from({ length: t.cols }, () => "General");
    const height = t.rows - firstRow;
    changed.formulasR1C1 = Array.from({ length: height }, (_, i) => keep[i]?.row ?? empty());
    changed.numberFormat = Array.from({ length: height }, (_, i) => keep[i]?.format ?? general());
    await context.sync();
  });
  await count({ rows: remove.size });
}

async function undoLast() {
  const u = undo.val;
  if (!u) return;
  await work("Undoing…", async () => {
    await jolt.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(u.sheet);
      for (const c of u.cells ?? []) {
        const cell = sheet.getRange(c.address);
        cell.formulas = [[c.formula]];
        cell.numberFormat = [[c.numberFormat]];
      }
      if (u.block) {
        const range = sheet.getRange(u.block.address);
        range.formulasR1C1 = u.block.formulasR1C1;
        range.numberFormat = u.block.numberFormat;
      }
      await context.sync();
    });
    undo.val = null;
    await analyze({ guessHeaders: false });
    jolt.notify(`Undid "${u.label}".`, "success");
  });
}

/** Select the cells an issue points to (for issues that need a person to look). @param {Issue} issue */
async function selectCells(issue) {
  const t = target.val;
  if (!t || !issue.select?.length) return;
  await work("", async () => {
    const shown = issue.select.slice(0, 200);
    const addresses = shown.map((c) => `${columnLetters(t.left + c.col)}${t.top + c.row + 1}`);
    await jolt.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(t.sheet);
      sheet.activate();
      // getRanges takes several areas at once ("A2,C7,D9"): ExcelApi 1.9.
      sheet.getRanges(addresses.join(",")).select();
      await context.sync();
    });
    if (issue.select.length > shown.length) jolt.notify(`Selected the first ${shown.length} of ${issue.select.length}.`, "info");
  });
}

/** Turn the block into an Excel Table, with filter buttons and banded rows. */
async function makeTable() {
  const t = target.val;
  if (!t) return;
  await work("Creating the table…", async () => {
    await jolt.run(async (context) => {
      const table = context.workbook.worksheets.getItem(t.sheet).tables.add(t.address, headers.val);
      table.getRange().format.autofitColumns();
      await context.sync();
    });
    jolt.notify("Converted to a table. Use Table Design → Convert to Range to undo.", "success");
  });
}

/** Keep a running total, saved with jolt.storage, of what this app has fixed. @param {{ cells?: number, rows?: number }} n */
async function count({ cells = 0, rows = 0 }) {
  stats.val = { cells: stats.val.cells + cells, rows: stats.val.rows + rows };
  await jolt.storage.set("stats", stats.val);
}

/**
 * Run a step with a status message, and show any error to the user.
 * @param {string} message @param {() => Promise<void>} fn
 */
async function work(message, fn) {
  if (busy.val) return;
  busy.val = message || "…";
  try {
    await fn();
  } catch (e) {
    jolt.notify(e instanceof Error ? e.message : String(e), "error");
  } finally {
    busy.val = "";
  }
}

/** "Sheet1!A1:D20" → ["Sheet1", "A1:D20"], unquoting names like 'My Sheet'. @param {string} address */
function splitAddress(address) {
  const i = address.lastIndexOf("!");
  return [address.slice(0, i).replace(/^'|'$/g, "").replace(/''/g, "'"), address.slice(i + 1)];
}

/** @param {unknown} f */
const isFormula = (f) => typeof f === "string" && f.startsWith("=");

/** The first row is probably headers if it's all text and some row below has a number in it. @param {unknown[][]} values */
function looksLikeHeaders(values) {
  if (values.length < 2) return false;
  const first = values[0].every((v) => typeof v === "string" && v.trim() !== "");
  return first && values.slice(1).some((row) => row.some((v) => typeof v === "number"));
}

// ---- The page ----------------------------------------------------------------------------------------

/** @param {Issue} issue */
const IssueRow = (issue) =>
  li(
    { class: issue.count ? "issue" : "issue ok" },
    div(
      { class: "issue-head" },
      strong(issue.title),
      span({ class: "count" }, issue.count ? issue.count.toLocaleString() : "✓"),
    ),
    issue.count ? p({ class: "detail" }, issue.detail) : null,
    issue.count
      ? issue.select
        ? button({ onclick: () => selectCells(issue), disabled: () => !!busy.val }, "Select them")
        : button({ class: "primary", onclick: () => fix(issue.key), disabled: () => !!busy.val }, issue.rows ? "Remove" : "Fix")
      : null,
  );

const App = () =>
  main(
    { class: "app" },
    header(h1("Data Cleanup"), p("Select a table of data (or click any cell in it), then check it for common problems.")),
    div(
      { class: "toolbar" },
      button({ class: "primary", onclick: () => analyze(), disabled: () => !!busy.val }, () => (target.val ? "Check again" : "Check the selection")),
      label(
        input({
          type: "checkbox",
          checked: headers,
          onchange: (/** @type {Event} */ e) => {
            headers.val = /** @type {HTMLInputElement} */ (e.target).checked;
            if (target.val) analyze({ guessHeaders: false });
          },
        }),
        " First row is headers",
      ),
    ),
    () => (busy.val ? p({ class: "status" }, busy.val) : span()),
    () => {
      const t = target.val;
      if (!t) return div();
      const total = issues.val.reduce((n, i) => n + i.count, 0);
      return div(
        p({ class: "target" }, `${t.sheet}!${t.address} · ${t.rows.toLocaleString()} rows × ${t.cols} columns`, total ? "" : " · no problems found"),
        ul({ class: "issues" }, ...issues.val.map(IssueRow)),
        div(
          { class: "toolbar" },
          () => (undo.val ? button({ onclick: undoLast, disabled: () => !!busy.val }, `Undo "${undo.val.label}"`) : span()),
          button({ onclick: makeTable, disabled: () => !!busy.val }, "Convert to table"),
        ),
      );
    },
    () =>
      stats.val.cells || stats.val.rows
        ? small({ class: "stats" }, `Fixed with Data Cleanup so far:${stats.val.cells.toLocaleString()} cells, ${stats.val.rows.toLocaleString()} rows.`)
        : span(),
  );

export function autoRun() {
  const saved = jolt.storage.get("stats");
  if (saved && typeof saved === "object") stats.val = { cells: 0, rows: 0, .../** @type {object} */ (saved) };
  van.add(/** @type {Element} */ (jolt.$("#app")), App());
}
