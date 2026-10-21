// @ts-check
import van from "van";
import { CATEGORIES, auditSheet, reportRows } from "./analyze.js";

/**
 * Formula Audit: scans every sheet for error values, formulas that break the pattern around them,
 * numbers typed into formulas, links to other workbooks and volatile functions. Click a finding to go
 * to the cell, or write the whole report to a new sheet. The app only reads the workbook, except when
 * you ask it to write the report.
 */

const { main, header, h1, p, button, select, option, ul, li, div, span, small, code, label, input } = van.tags;

/**
 * @typedef {import("./analyze.js").Finding} Finding
 * @typedef {import("./analyze.js").Category} Category
 */

/** Ask before scanning more cells than this; big workbooks take a while. */
const LARGE = 1_000_000;
/** Cells read per request, which keeps each response well under Excel's size limits. */
const CHUNK_CELLS = 20_000;
/** Findings listed at a time. */
const PAGE = 200;

const findings = van.state(/** @type {Finding[]} */ ([]));
const scanned = van.state(/** @type {{ sheets: number, cells: number, at: Date } | null} */ (null));
const progress = van.state("");
const shown = van.state(/** @type {Record<Category, boolean>} */ ({ error: true, inconsistent: true, hardcoded: true, external: true, volatile: true }));
const sheetFilter = van.state("");
const limit = van.state(PAGE);

/** @param {{ activeOnly?: boolean }} [o] */
async function scan({ activeOnly = false } = {}) {
  if (progress.val) return;
  progress.val = "Looking at the workbook…";
  try {
    // First, each sheet's used range: where it is and how big.
    const sheets = await jolt.run(async (context) => {
      const ws = context.workbook.worksheets;
      ws.load("items/name");
      const current = ws.getActiveWorksheet();
      current.load("name");
      await context.sync();
      const list = (activeOnly ? ws.items.filter((s) => s.name === current.name) : ws.items).map((s) => ({ name: s.name, used: s.getUsedRangeOrNullObject() }));
      for (const s of list) s.used.load("rowIndex,columnIndex,rowCount,columnCount");
      await context.sync();
      return list
        .filter((s) => !s.used.isNullObject)
        .map((s) => ({ name: s.name, top: s.used.rowIndex, left: s.used.columnIndex, rows: s.used.rowCount, cols: s.used.columnCount }));
    });
    const total = sheets.reduce((n, s) => n + s.rows * s.cols, 0);
    if (total > LARGE && !(await jolt.confirm(`This will read ${total.toLocaleString()} cells, which may take a minute or two. Continue?`, { title: "Large workbook", ok: "Scan" }))) return;

    /** @type {Finding[]} */
    const all = [];
    for (const [i, s] of sheets.entries()) {
      const step = Math.max(1, Math.floor(CHUNK_CELLS / s.cols));
      for (let from = 0; from < s.rows; from += step) {
        progress.val = `Scanning ${s.name} (${i + 1} of ${sheets.length})${s.rows > step ? `, ${Math.round((100 * from) / s.rows)}%` : ""}…`;
        const to = Math.min(s.rows, from + step);
        // One extra row each side, so formulas at the slice's edges can be compared with their neighbours.
        const start = Math.max(0, from - 1);
        const end = Math.min(s.rows, to + 1);
        const data = await jolt.run(async (context) => {
          const range = context.workbook.worksheets.getItem(s.name).getRangeByIndexes(s.top + start, s.left, end - start, s.cols);
          range.load("formulas,formulasR1C1,values,valueTypes");
          await context.sync();
          return { formulas: range.formulas, formulasR1C1: range.formulasR1C1, values: range.values, valueTypes: /** @type {string[][]} */ (range.valueTypes) };
        });
        all.push(...auditSheet({ name: s.name, top: s.top + start, left: s.left, ...data }, { from: from - start, to: to - start }));
      }
    }
    findings.val = all;
    scanned.val = { sheets: sheets.length, cells: total, at: new Date() };
    limit.val = PAGE;
  } catch (e) {
    jolt.notify(e instanceof Error ? e.message : String(e), "error");
  } finally {
    progress.val = "";
  }
}

/** @param {Finding} f */
async function goTo(f) {
  try {
    await jolt.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(f.sheet);
      sheet.activate();
      sheet.getRange(f.address).select();
      await context.sync();
    });
  } catch (e) {
    jolt.notify(e instanceof Error ? e.message : String(e), "error");
  }
}

/** Write the findings now shown to a new sheet, with links back to each cell. */
async function exportReport() {
  const list = visible();
  if (!list.length) return jolt.notify("There's nothing to export.", "info");
  try {
    const rows = reportRows(list);
    const name = await jolt.run(async (context) => {
      const ws = context.workbook.worksheets;
      ws.load("items/name");
      await context.sync();
      const taken = new Set(ws.items.map((s) => s.name.toLowerCase()));
      let sheetName = "Formula Audit";
      for (let i = 2; taken.has(sheetName.toLowerCase()); i++) sheetName = `Formula Audit (${i})`;
      const sheet = ws.add(sheetName);
      const range = sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length);
      range.formulas = rows;
      range.getRow(0).format.font.bold = true;
      sheet.freezePanes.freezeRows(1);
      range.format.autofitColumns();
      sheet.getRange("E:E").format.columnWidth = 300;
      sheet.activate();
      await context.sync();
      return sheetName;
    });
    jolt.notify(`Wrote ${list.length.toLocaleString()} findings to "${name}".`, "success");
  } catch (e) {
    jolt.notify(e instanceof Error ? e.message : String(e), "error");
  }
}

/** The findings that pass the category and sheet filters. */
const visible = () => findings.val.filter((f) => shown.val[f.category] && (!sheetFilter.val || f.sheet === sheetFilter.val));

// ---- The page ----------------------------------------------------------------------------------------

/** @param {Category} cat */
const CategoryToggle = (cat) => {
  const n = findings.val.filter((f) => f.category === cat && (!sheetFilter.val || f.sheet === sheetFilter.val)).length;
  return label(
    { class: `chip${n ? "" : " zero"}`, title: CATEGORIES[cat].why },
    input({
      type: "checkbox",
      checked: shown.val[cat],
      onchange: (/** @type {Event} */ e) => (shown.val = { ...shown.val, [cat]: /** @type {HTMLInputElement} */ (e.target).checked }),
    }),
    ` ${CATEGORIES[cat].title} `,
    span({ class: "n" }, n.toLocaleString()),
  );
};

/** @param {Finding} f */
const FindingItem = (f) =>
  li(
    button(
      { class: "finding", onclick: () => goTo(f) },
      span({ class: `tag ${f.category}` }, CATEGORIES[f.category].title),
      span({ class: "where" }, `${f.sheet}!${f.address}`),
      span({ class: "what" }, f.detail),
      f.formula ? code(f.formula.length > 120 ? `${f.formula.slice(0, 120)}…` : f.formula) : null,
    ),
  );

const Results = () => {
  const s = scanned.val;
  if (!s) return div({ class: "intro" }, ...Object.values(CATEGORIES).map((c) => p(span({ class: "strong" }, c.title), `: ${c.why}`)));
  const list = visible();
  const sheetNames = [...new Set(findings.val.map((f) => f.sheet))];
  return div(
    p({ class: "muted" }, `Scanned ${s.cells.toLocaleString()} cells on ${s.sheets} sheet${s.sheets === 1 ? "" : "s"} at ${s.at.toLocaleTimeString()}.`, findings.val.length ? "" : " Nothing found."),
    findings.val.length
      ? div(
          div({ class: "chips" }, ...(/** @type {Category[]} */ (Object.keys(CATEGORIES))).map(CategoryToggle)),
          div(
            { class: "toolbar" },
            sheetNames.length > 1
              ? select(
                  { onchange: (/** @type {Event} */ e) => (sheetFilter.val = /** @type {HTMLSelectElement} */ (e.target).value) },
                  option({ value: "" }, "All sheets"),
                  ...sheetNames.map((n) => option({ value: n, selected: n === sheetFilter.val }, n)),
                )
              : null,
            button({ onclick: exportReport }, "Write report to a sheet"),
          ),
          ul({ class: "findings" }, ...list.slice(0, limit.val).map(FindingItem)),
          list.length > limit.val ? button({ onclick: () => (limit.val += PAGE) }, `Show more (${(list.length - limit.val).toLocaleString()} left)`) : null,
        )
      : null,
  );
};

const App = () =>
  main(
    { class: "app" },
    header(h1("Formula Audit"), p("Find the formulas most likely to be wrong, before someone else does.")),
    div(
      { class: "toolbar" },
      button({ class: "primary", onclick: () => scan(), disabled: () => !!progress.val }, "Scan workbook"),
      button({ onclick: () => scan({ activeOnly: true }), disabled: () => !!progress.val }, "Scan this sheet"),
    ),
    () => (progress.val ? p({ class: "muted" }, progress.val) : span()),
    Results,
    small({ class: "muted" }, "Read-only: this app changes nothing unless you write the report to a sheet."),
  );

export function autoRun() {
  van.add(/** @type {Element} */ (jolt.$("#app")), App());
}
