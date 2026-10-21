// @ts-check
import van from "van";
import { contentsRows, filterMap, freeSheetName, sizeLabel } from "./analyze.js";

/**
 * Workbook Map: everything in the workbook in one searchable list (sheets, tables, charts, PivotTables
 * and named ranges). Click anything to go to it, and build a Contents sheet of links in one click.
 */

const { main, header, h1, h2, p, button, input, ul, li, div, span, small } = van.tags;

/**
 * @typedef {import("./analyze.js").SheetInfo} SheetInfo
 * @typedef {import("./analyze.js").NameInfo} NameInfo
 */

const sheets = van.state(/** @type {SheetInfo[]} */ ([]));
const names = van.state(/** @type {NameInfo[]} */ ([]));
const active = van.state("");
const query = van.state("");
const loading = van.state(false);

/** Read the workbook's structure. Two round trips: the sheets first, then everything on them. */
async function refresh() {
  loading.val = true;
  try {
    await jolt.run(async (context) => {
      const wb = context.workbook;
      const ws = wb.worksheets;
      ws.load("items/name,items/visibility");
      const current = ws.getActiveWorksheet();
      current.load("name");
      wb.names.load("items/name,items/formula,items/visible");
      await context.sync();

      const details = ws.items.map((sheet) => {
        // valuesOnly: ignore cells that only have formatting.
        const used = sheet.getUsedRangeOrNullObject(true);
        used.load("address,rowCount,columnCount");
        sheet.tables.load("items/name");
        sheet.charts.load("items/name");
        sheet.pivotTables.load("items/name");
        sheet.names.load("items/name,items/formula,items/visible");
        return { sheet, used };
      });
      await context.sync();

      // Each table's and PivotTable's own range, for jumping to it.
      const ranges = details.map(({ sheet }) => ({
        tables: sheet.tables.items.map((t) => t.getRange().load("address")),
        pivots: sheet.pivotTables.items.map((pt) => pt.layout.getRange().load("address")),
      }));
      await context.sync();

      sheets.val = details.map(({ sheet, used }, i) => ({
        name: sheet.name,
        visibility: /** @type {SheetInfo["visibility"]} */ (sheet.visibility),
        used: used.isNullObject ? "" : localAddress(used.address),
        rows: used.isNullObject ? 0 : used.rowCount,
        cols: used.isNullObject ? 0 : used.columnCount,
        tables: sheet.tables.items.map((t, j) => ({ name: t.name, address: localAddress(ranges[i].tables[j].address) })),
        charts: sheet.charts.items.map((c) => ({ name: c.name, address: "" })),
        pivots: sheet.pivotTables.items.map((pt, j) => ({ name: pt.name, address: localAddress(ranges[i].pivots[j].address) })),
      }));
      names.val = [
        ...wb.names.items.map((n) => ({ name: n.name, formula: n.formula, scope: "Workbook", visible: n.visible })),
        ...details.flatMap(({ sheet }) => sheet.names.items.map((n) => ({ name: n.name, formula: n.formula, scope: sheet.name, visible: n.visible }))),
      ].filter((n) => n.visible);
      active.val = current.name;
    });
  } catch (e) {
    jolt.notify(e instanceof Error ? e.message : String(e), "error");
  } finally {
    loading.val = false;
  }
}

/**
 * Go to a place: activate its sheet, then select the range (or activate the chart).
 * @param {string} sheetName @param {{ address?: string, chart?: string, name?: NameInfo }} where
 */
async function go(sheetName, where = {}) {
  try {
    await jolt.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.activate();
      if (where.chart) sheet.charts.getItem(where.chart).activate();
      else if (where.address) sheet.getRange(where.address).select();
      await context.sync();
    });
    active.val = sheetName;
  } catch (e) {
    jolt.notify(e instanceof Error ? e.message : String(e), "error");
  }
}

/** Go to what a named range refers to. @param {NameInfo} n */
async function goToName(n) {
  try {
    await jolt.run(async (context) => {
      const item = n.scope === "Workbook" ? context.workbook.names.getItem(n.name) : context.workbook.worksheets.getItem(n.scope).names.getItem(n.name);
      const range = item.getRangeOrNullObject();
      range.load("isNullObject");
      range.worksheet.load("name");
      await context.sync();
      if (range.isNullObject) throw new Error(`${n.name} is a formula or constant (${n.formula}), not a range.`);
      range.worksheet.activate();
      range.select();
      await context.sync();
      active.val = range.worksheet.name;
    });
  } catch (e) {
    jolt.notify(e instanceof Error ? e.message : String(e), "error");
  }
}

/** @param {string} sheetName */
async function unhide(sheetName) {
  await jolt.run(async (context) => {
    context.workbook.worksheets.getItem(sheetName).visibility = Excel.SheetVisibility.visible;
    await context.sync();
  });
  await refresh();
  await go(sheetName);
}

/** Add a sheet at the front of the workbook with a link to every other sheet. */
async function createContents() {
  try {
    let name = "Contents";
    const existing = sheets.val.map((s) => s.name);
    let replace = false;
    if (existing.some((n) => n.toLowerCase() === "contents")) {
      replace = await jolt.confirm(`There is already a sheet called "Contents". Replace it with a new one? Choose Keep to add a second contents sheet instead.`, {
        title: "Contents sheet",
        ok: "Replace",
        cancel: "Keep",
      });
      if (!replace) name = freeSheetName(existing);
    }
    const rows = contentsRows(sheets.val, name);
    await jolt.run(async (context) => {
      const ws = context.workbook.worksheets;
      if (replace) ws.getItem(name).delete();
      const sheet = ws.add(name);
      sheet.position = 0;
      const range = sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length);
      range.formulas = rows;
      const head = range.getRow(0);
      head.format.font.bold = true;
      head.format.fill.color = "#E4F6E8";
      sheet.freezePanes.freezeRows(1);
      range.format.autofitColumns();
      sheet.activate();
      await context.sync();
    });
    await refresh();
    jolt.notify(`Added "${name}" with links to ${rows.length - 1} sheets.`, "success");
  } catch (e) {
    jolt.notify(e instanceof Error ? e.message : String(e), "error");
  }
}

/** "'My Sheet'!A1:B2" → "A1:B2". @param {string} address */
const localAddress = (address) => address.slice(address.lastIndexOf("!") + 1);

/** Refresh when sheets are added, removed or renamed (debounced: one refresh per burst of events). */
async function watchWorkbook() {
  /** @type {ReturnType<typeof setTimeout> | undefined} */
  let timer;
  const later = async () => {
    clearTimeout(timer);
    timer = setTimeout(refresh, 300);
  };
  try {
    await jolt.run(async (context) => {
      const ws = context.workbook.worksheets;
      ws.onAdded.add(later);
      ws.onDeleted.add(later);
      // An event handler runs after this Excel.run has finished, so it needs a context of its own.
      ws.onActivated.add((e) =>
        jolt.run(async (ctx) => {
          const sheet = ctx.workbook.worksheets.getItem(e.worksheetId);
          sheet.load("name");
          await ctx.sync();
          active.val = sheet.name;
        }),
      );
      if (ws.onNameChanged) ws.onNameChanged.add(later);
      await context.sync();
    });
  } catch {
    // Older Excel: no events. The Refresh button still works.
  }
}

// ---- The page ----------------------------------------------------------------------------------------

/** @param {string} kind @param {import("./analyze.js").Part} part @param {() => void} onclick */
const PartItem = (kind, part, onclick) => li({ class: "part" }, button({ class: "link", onclick }, span({ class: "kind" }, kind), part.name));

/** @param {SheetInfo} s */
const SheetItem = (s) =>
  li(
    { class: () => `sheet${s.name === active.val ? " active" : ""}${s.visibility !== "Visible" ? " hidden" : ""}` },
    div(
      { class: "sheet-row" },
      button({ class: "link sheet-name", onclick: () => go(s.name, { address: s.used ? s.used.split(":")[0] : "A1" }), disabled: s.visibility !== "Visible" }, s.name),
      s.visibility !== "Visible" ? button({ class: "small", onclick: () => unhide(s.name) }, s.visibility === "Hidden" ? "Unhide" : "Unhide (very hidden)") : null,
    ),
    small(sizeLabel(s)),
    s.tables.length || s.charts.length || s.pivots.length
      ? ul(
          ...s.tables.map((t) => PartItem("Table", t, () => go(s.name, { address: t.address }))),
          ...s.pivots.map((pt) => PartItem("PivotTable", pt, () => go(s.name, { address: pt.address }))),
          ...s.charts.map((c) => PartItem("Chart", c, () => go(s.name, { chart: c.name }))),
        )
      : null,
  );

const Summary = () => {
  const all = sheets.val;
  /** @param {number} n @param {string} w */
  const count = (n, w) => `${n} ${w}${n === 1 ? "" : "s"}`;
  const hidden = all.filter((s) => s.visibility !== "Visible").length;
  return p(
    { class: "summary" },
    [
      count(all.length, "sheet") + (hidden ? ` (${hidden} hidden)` : ""),
      count(all.reduce((n, s) => n + s.tables.length, 0), "table"),
      count(all.reduce((n, s) => n + s.pivots.length, 0), "PivotTable"),
      count(all.reduce((n, s) => n + s.charts.length, 0), "chart"),
      count(names.val.length, "named range"),
    ].join(" · "),
  );
};

const App = () =>
  main(
    { class: "app" },
    header(h1("Workbook Map"), p("Everything in this workbook. Click anything to go to it.")),
    div(
      { class: "toolbar" },
      button({ class: "primary", onclick: createContents }, "Create Contents sheet"),
      button({ onclick: refresh, disabled: loading }, () => (loading.val ? "Reading…" : "Refresh")),
    ),
    input({ class: "search", type: "search", placeholder: "Search sheets, tables, charts and names", value: query, oninput: (/** @type {Event} */ e) => (query.val = /** @type {HTMLInputElement} */ (e.target).value) }),
    Summary,
    () => {
      const shown = filterMap(sheets.val, names.val, query.val);
      return div(
        h2("Sheets"),
        shown.sheets.length ? ul({ class: "tree" }, ...shown.sheets.map(SheetItem)) : p({ class: "empty" }, "No matching sheets."),
        h2("Named ranges"),
        shown.names.length
          ? ul(
              { class: "tree" },
              ...shown.names.map((n) =>
                li({ class: "part" }, button({ class: "link", onclick: () => goToName(n) }, n.name), small(` ${n.formula}${n.scope === "Workbook" ? "" : ` (${n.scope} only)`}`)),
              ),
            )
          : p({ class: "empty" }, query.val ? "No matching names." : "This workbook has no named ranges."),
      );
    },
  );

export async function autoRun() {
  van.add(/** @type {Element} */ (jolt.$("#app")), App());
  await refresh();
  await watchWorkbook();
}
