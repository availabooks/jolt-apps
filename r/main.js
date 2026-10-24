// @ts-check
import van from "van";
import {
  analysisComment,
  analysisFileName,
  blocksToGrid,
  freeSheetName,
  looksLikeHeaders,
  parseAnalysis,
  rName,
  safeFileName,
  serializeAnalysis,
  sheetName,
  splitAddress,
  toFrame,
  uniqueNames,
} from "./data.js";
import { EXAMPLES, pickColumns } from "./examples.js";
import { markdownToHtml } from "./markdown.js";
import { buildReport, reportHtml } from "./report.js";
import { buildFixPrompt, buildInterpretPrompt, buildPrompt, extractCode, splitReply } from "./prompt.js";
import { bindFrames, ensurePackages, restartR, runCode, startR } from "./rsession.js";

/**
 * R for Excel: run R code on data from the workbook and put the results back in Excel.
 *
 * The idea is that the user's own AI assistant writes the R code. This app reads the data, builds a
 * prompt that describes it, runs the code that comes back, and writes the tables and plots to sheets.
 *
 * How it talks to Excel: every Office call goes through `jolt.run(async (context) => …)`, which is
 * Excel.run. R runs in a web worker (webR); see rsession.js.
 */

const { main, header, h1, h2, h3, p, button, label, input, textarea, select, option, ul, li, div, span, small, pre, img, table, tr, td, th, details, summary } =
  van.tags;

/** Larger blocks are slow to read and to send to R. */
const MAX_CELLS = 500_000;
/** Plots are inserted into Excel at this width, in points. */
const PLOT_WIDTH = 480;

/**
 * @typedef {import("./data.js").Frame} Frame
 * @typedef {import("./data.js").Output} Output
 * @typedef {import("./data.js").Block} Block
 * @typedef {import("./data.js").Grid} Grid
 * @typedef {import("./rsession.js").ConsoleLine} ConsoleLine
 * @typedef {{ id: string, name: string, kind: "table" | "range", table?: string, sheet?: string, address?: string, headers: boolean }} Source
 *   Where a data frame comes from: an Excel Table (which can grow) or a fixed range.
 * @typedef {{ console: ConsoleLine[], plots: string[], outputs: Output[], error: string | null, ms: number, code: string, ranAt: Date }} Result
 *   plots are PNG data URLs
 * @typedef {{ id: string, name: string, description: string, sources: Source[], question: string, includeSample: boolean, code: string, explanation: string, interpretation: string }} Analysis
 *   One tab: everything the user has set up for one analysis. Saved in the workbook.
 */

/**
 * Tabs. The states below (description, sources, question, includeSample, explanation, interpretation, promptText, result) and the code box always
 * hold the active tab. Switching tabs copies them into the tab's record and loads the other tab's.
 */
const analyses = van.state(/** @type {Analysis[]} */ ([]));
const activeId = van.state("");
/** Each tab's last results and prompt, for this session only: they aren't saved in the workbook. */
const session = new Map(/** @type {[string, { result: Result | null, promptText: string }][]} */ ([]));

const rStatus = van.state(/** @type {{ state: "starting" | "ready" | "failed", text: string }} */ ({ state: "starting", text: "" }));
const sources = van.state(/** @type {Source[]} */ ([]));
const frames = van.state(/** @type {Record<string, Frame>} */ ({}));
const tables = van.state(/** @type {string[]} */ ([]));
const description = van.state("");
const question = van.state("");
/** The AI's Markdown explanation of why its code fits the question. */
const explanation = van.state("");
/** The AI's Markdown interpretation of the results. */
const interpretation = van.state("");
const includeSample = van.state(true);
const promptText = van.state("");
const result = van.state(/** @type {Result | null} */ (null));
const busy = van.state("");
let rVersion = "";
let helpers = "";

// ---- R ---------------------------------------------------------------------------------------------------

async function bootR() {
  rStatus.val = { state: "starting", text: "Starting R… The first time, this downloads about 30 MB." };
  try {
    helpers ||= await jolt.files.read("jolt.R");
    rVersion = await startR(helpers);
    rStatus.val = { state: "ready", text: `R ${rVersion} is ready.` };
  } catch (e) {
    rStatus.val = { state: "failed", text: `R didn't start: ${message(e)}` };
  }
}

/** Wait for R to finish starting, starting it again if it failed. */
async function readyR() {
  if (rStatus.val.state === "failed") await bootR();
  busy.val = "Waiting for R to start…";
  rVersion = await startR(helpers);
}

async function restart() {
  await work("Restarting R…", async () => {
    rStatus.val = { state: "starting", text: "Restarting R…" };
    rVersion = await restartR(helpers);
    rStatus.val = { state: "ready", text: `R ${rVersion} restarted. Every variable was cleared.` };
  });
}

// ---- Data: Excel → R -------------------------------------------------------------------------------------

/** The block to use: the selection, or the data around the active cell when only one cell is selected. */
async function selectedBlock() {
  /** @param {(sel: Excel.Range) => Excel.Range} pick */
  const read = (pick) =>
    jolt.run(async (context) => {
      const range = pick(context.workbook.getSelectedRange());
      range.load("address,cellCount,isNullObject");
      await context.sync();
      if (range.isNullObject) throw new Error("The selection has no data in it.");
      return { address: range.address, cells: range.cellCount };
    });
  // A whole selected column or row is mostly empty: keep only the part that has data.
  const selection = await read((sel) => sel.getIntersectionOrNullObject(sel.worksheet.getUsedRange()));
  if (selection.cells > 1) return selection;
  try {
    // One cell: use the "current region", the block of filled cells around it (Ctrl+A in Excel).
    return await read((sel) => sel.getSurroundingRegion());
  } catch {
    return read((sel) => sel.worksheet.getUsedRange());
  }
}

/** Names and full-range addresses of every Excel Table in the workbook. */
async function readTables() {
  return jolt.run(async (context) => {
    const items = context.workbook.tables;
    items.load("items/name");
    await context.sync();
    const ranges = items.items.map((t) => t.getRange().load("address"));
    await context.sync();
    return items.items.map((t, i) => ({ name: t.name, address: ranges[i].address }));
  });
}

/** Add the selected data (or the Table it's in) as a data frame. */
async function addSelection() {
  await work("Reading the selection…", async () => {
    const block = await selectedBlock();
    const inTable = (await readTables()).find((t) => t.address === block.address);
    if (inTable) return addSource({ kind: "table", table: inTable.name, headers: true }, inTable.name);
    const [sheet, address] = splitAddress(block.address);
    const { values } = await readRange({ id: "", name: "", kind: "range", sheet, address, headers: false });
    await addSource({ kind: "range", sheet, address, headers: looksLikeHeaders(values) }, sheet);
  });
}

/** @param {string} name */
async function addTable(name) {
  await work("Reading the table…", () => addSource({ kind: "table", table: name, headers: true }, name));
}

/** @param {Omit<Source, "id" | "name">} source @param {string} suggested */
async function addSource(source, suggested) {
  const taken = sources.val.map((s) => s.name);
  const name = uniqueNames([...taken, rName(suggested, "data")]).at(-1) ?? "data";
  const s = { ...source, id: crypto.randomUUID(), name };
  await readFrame(s);
  sources.val = [...sources.val, s];
  await save();
  void refreshTables();
}

/** @param {string} id @param {Partial<Source>} change */
async function updateSource(id, change) {
  sources.val = sources.val.map((s) => (s.id === id ? { ...s, ...change } : s));
  await save();
  const s = sources.val.find((x) => x.id === id);
  if (s) await work("Reading…", async () => void (await readFrame(s)));
}

/** @param {string} id */
async function removeSource(id) {
  sources.val = sources.val.filter((s) => s.id !== id);
  const { [id]: _, ...rest } = frames.val;
  frames.val = rest;
  await save();
}

/** @param {Source} s */
async function readRange(s) {
  return jolt.run(async (context) => {
    let range;
    let headers = s.headers;
    if (s.kind === "table") {
      const t = context.workbook.tables.getItemOrNullObject(s.table ?? "");
      t.load("isNullObject,showHeaders");
      await context.sync();
      if (t.isNullObject) throw new Error(`There's no table called "${s.table}" any more. Remove ${s.name} from the data list.`);
      headers = t.showHeaders;
      range = t.getRange();
    } else {
      const sheet = context.workbook.worksheets.getItemOrNullObject(s.sheet ?? "");
      await context.sync();
      if (sheet.isNullObject) throw new Error(`There's no sheet called "${s.sheet}" any more. Remove ${s.name} from the data list.`);
      range = sheet.getRange(s.address);
    }
    range.load("cellCount");
    await context.sync();
    if (range.cellCount > MAX_CELLS) throw new Error(`${s.name} has ${range.cellCount.toLocaleString()} cells. The limit is ${MAX_CELLS.toLocaleString()}.`);
    range.load("values,numberFormat,address");
    await context.sync();
    return { values: range.values, formats: /** @type {string[][]} */ (range.numberFormat), headers, address: range.address };
  });
}

/** Read one source from the workbook and keep its frame. @param {Source} s */
async function readFrame(s) {
  const { values, formats, headers } = await readRange(s);
  const frame = toFrame(values, formats, { name: s.name, headers });
  if (!frame.rows) throw new Error(`${s.name} has no rows of data.`);
  frames.val = { ...frames.val, [s.id]: frame };
  return frame;
}

/** Read every source again, so edits in the workbook are included. */
async function readFrames() {
  const out = [];
  for (const s of sources.val) out.push(await readFrame(s));
  return out;
}

// ---- Prompts ---------------------------------------------------------------------------------------------

async function copyPrompt() {
  await work("Reading the data…", async () => {
    const text = buildPrompt({ frames: await readFrames(), question: askedQuestion(), includeSample: includeSample.val, rVersion });
    await offerText(text, "Prompt copied. Paste it into your AI assistant, then paste its code into the box below.");
  });
}

async function copyFixPrompt() {
  const r = result.val;
  if (!r?.error) return;
  await work("Reading the data…", async () => {
    const consoleTail = r.console.filter((l) => l.kind !== "error").map((l) => l.text).slice(-15).join("\n");
    const text = buildFixPrompt({ frames: await readFrames(), code: r.code, error: r.error ?? "", consoleTail, includeSample: includeSample.val, rVersion });
    await offerText(text, "Copied. Paste it into your AI assistant and run the corrected code it sends back.");
  });
}

// ---- Report ----------------------------------------------------------------------------------------------

/** Everything the report needs, from the active tab and its last run. */
function reportInput() {
  const r = /** @type {Result} */ (result.val);
  return {
    title: activeAnalysis()?.name ?? "Analysis",
    ranAt: r.ranAt,
    rVersion,
    ms: r.ms,
    question: askedQuestion(),
    description: description.val,
    sources: sources.val.map((s) => ({ name: s.name, where: sourceWhere(s), frame: frames.val[s.id] })),
    code: r.code,
    explanation: explanation.val,
    error: r.error,
    consoleLines: consoleText(r.console),
    outputs: r.outputs,
    plots: r.plots.length,
    interpretation: interpretation.val,
  };
}

async function copyReport() {
  if (!result.val) return;
  await offerText(buildReport(reportInput()), "Copied the report. Paste it into a document or a Markdown editor.");
}

/** Download the report as a web page with the plots in it, which opens in Word and prints to PDF. */
function downloadReport() {
  const r = result.val;
  if (!r) return;
  // The analysis goes in too, as a comment, so importing the report rebuilds the analysis.
  const html = reportHtml(reportInput(), r.plots, analysisComment(activeJSON().json));
  downloadFile(`${safeFileName(activeAnalysis()?.name ?? "analysis")} report.html`, html, "text/html", "Copy report");
}

/** Copy a prompt that asks the AI to interpret the last run's results. */
async function copyInterpretPrompt() {
  const r = result.val;
  if (!r) return;
  const text = buildInterpretPrompt({
    question: askedQuestion(),
    consoleLines: consoleText(r.console),
    results: r.outputs.map((o) => ({ name: o.name, rows: outputGrid(o).values })),
    plots: r.plots.length,
    error: r.error,
  });
  await offerText(text, "Copied. Paste it into the same AI conversation that wrote the R code, then paste its answer below.");
}

/** Copy text to the clipboard. If the task pane won't allow it, show the text to copy by hand. @param {string} text @param {string} done */
async function offerText(text, done) {
  promptText.val = text;
  if (await copyText(text)) jolt.notify(done, "success");
  else {
    const box = /** @type {HTMLTextAreaElement | null} */ (jolt.$(".prompt-text"));
    box?.focus();
    box?.select();
    jolt.notify("Couldn't reach the clipboard. The prompt is selected below: press Ctrl+C (⌘C on a Mac).", "info");
  }
}

/** @param {string} text */
async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // Some Office task panes block the Clipboard API: fall back to the older copy command.
    const area = document.createElement("textarea");
    area.value = text;
    area.style.position = "fixed";
    area.style.opacity = "0";
    /** @type {Node} */ (jolt.root ?? document.body).appendChild(area);
    area.select();
    const ok = document.execCommand("copy");
    area.remove();
    return ok;
  }
}

// ---- Running code ----------------------------------------------------------------------------------------

/** @type {HTMLTextAreaElement} */
let editor;

async function run() {
  const code = extractCode(editor.value);
  if (!code.trim()) return jolt.notify("Paste or type some R code first.", "info");
  if (code !== editor.value) setCode(code);
  await work("Reading the data…", async () => {
    await readyR();
    busy.val = "Reading the data…";
    const data = await readFrames();
    await bindFrames(data);
    await ensurePackages(code, (m) => (busy.val = m));
    busy.val = "Running…";
    const r = await runCode(code);
    result.val = { console: r.console, plots: r.images.map(toPNG), outputs: r.outputs, error: r.error, ms: r.ms, code, ranAt: new Date() };
    void refreshTables();
    const auto = r.outputs.filter((o) => o.sheet || o.cell);
    if (auto.length) {
      busy.val = "Writing to Excel…";
      await writeAuto(auto);
    }
  });
}

/** @param {ImageBitmap} image */
function toPNG(image) {
  const canvas = document.createElement("canvas");
  canvas.width = image.width;
  canvas.height = image.height;
  canvas.getContext("2d")?.drawImage(image, 0, 0);
  image.close();
  return canvas.toDataURL("image/png");
}

/** @param {string} code */
function setCode(code) {
  editor.value = code;
  void save();
}

/** @param {number} i */
async function loadExample(i) {
  const ex = EXAMPLES[i];
  if (!ex) return;
  if (editor.value.trim() && !(await jolt.confirm("Replace the code in the box with this example?", { title: ex.title, ok: "Replace" }))) return;
  const first = sources.val[0] && frames.val[sources.val[0].id];
  setCode(ex.code(pickColumns(first)));
}

// ---- Results: R → Excel ----------------------------------------------------------------------------------

/** Border colour for data cells, close to Excel's own gridlines. */
const GRIDLINE = "#D9D9D9";

/**
 * Write a grid of cells with its styling, starting at `anchor`. Queues the changes; the caller syncs.
 * @param {Excel.Range} anchor top-left cell @param {Grid} grid @param {{ table?: boolean }} [options]
 */
function putGrid(anchor, grid, options = {}) {
  if (!grid.values.length) return;
  const range = anchor.getResizedRange(grid.values.length - 1, grid.width - 1);
  range.values = grid.values;
  /** @param {import("./data.js").Mark} m */
  const area = (m) => anchor.getOffsetRange(m.row, m.col).getResizedRange(m.rows - 1, m.cols - 1);
  /** @param {Excel.Range} r @param {"EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | "InsideHorizontal" | "InsideVertical"} edge @param {string} color @param {"Thin" | "Medium"} [weight] */
  const line = (r, edge, color, weight = "Thin") => {
    const b = r.format.borders.getItem(edge);
    b.style = "Continuous";
    b.color = color;
    b.weight = weight;
  };
  // Light grey lines stand in for the gridlines our sheets hide. They go first so the header and end lines draw over them.
  for (const m of grid.marks) {
    if (m.kind !== "table") continue;
    const r = area(m);
    for (const edge of /** @type {const} */ (["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"])) line(r, edge, GRIDLINE);
  }
  for (const m of grid.marks) {
    const r = area(m);
    if (m.kind === "title") r.format.font.bold = true;
    if (m.kind === "header") {
      r.format.font.italic = true;
      line(r, "EdgeTop", "#000000", "Medium");
      line(r, "EdgeBottom", "#000000");
    }
    if (m.kind === "end") line(r, "EdgeBottom", "#000000");
    if (m.kind === "text") r.format.font.name = "Consolas";
    if (m.kind === "format" && m.format) r.numberFormat = Array.from({ length: m.rows }, () => [m.format]);
    if (m.kind === "table") r.format.autofitColumns();
  }
  if (options.table) {
    const t = grid.marks.find((m) => m.kind === "table");
    if (t) anchor.worksheet.tables.add(area(t), true);
  }
}

/** Sheets this app has made, so it can clear them before writing to them again. */
const madeSheets = new Set(/** @type {string[]} */ ([]));

/** A new, empty sheet with a free name based on `base`, activated. @param {Excel.RequestContext} context @param {string} base */
async function newSheet(context, base) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  const name = freeSheetName(base, sheets.items.map((s) => s.name));
  const sheet = sheets.add(name);
  sheet.showGridlines = false;
  sheet.activate();
  madeSheets.add(name);
  void jolt.storage.set("madeSheets", [...madeSheets]);
  return sheet;
}

/** Turn an output into a grid, putting its name above it if the first block has no title of its own. @param {Output} o */
function outputGrid(o) {
  const blocks = o.blocks[0]?.title || !o.name ? o.blocks : [{ title: o.name }, ...o.blocks];
  return blocksToGrid(o.table ? o.blocks : blocks);
}

/** @param {Output} o @param {"sheet" | "selection"} where */
async function writeOutput(o, where) {
  await work("Writing to Excel…", async () => {
    const grid = outputGrid(o);
    await jolt.run(async (context) => {
      const anchor = where === "sheet" ? (await newSheet(context, o.name || "R Output")).getRange("A1") : context.workbook.getSelectedRange().getCell(0, 0);
      putGrid(anchor, grid, { table: o.table });
      await context.sync();
    });
    jolt.notify(where === "sheet" ? `Wrote "${o.name}" to a new sheet.` : `Wrote "${o.name}" at the selected cell.`, "success");
  });
}

/**
 * Outputs that named their own sheet or cell in excel_out(). Outputs for the same sheet without a cell are
 * stacked from A1, after clearing the sheet if this app made it.
 * @param {Output[]} outputs
 */
async function writeAuto(outputs) {
  await jolt.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const names = new Map(sheets.items.map((s) => [s.name.toLowerCase(), s]));
    /** @param {string} name */
    const sheetFor = (name) => {
      const n = sheetName(name);
      let s = names.get(n.toLowerCase());
      if (!s) {
        s = sheets.add(n);
        s.showGridlines = false;
        names.set(n.toLowerCase(), s);
        madeSheets.add(n);
      }
      return s;
    };
    const active = context.workbook.worksheets.getActiveWorksheet();
    /** @type {Map<string, Output[]>} */
    const stacks = new Map();
    for (const o of outputs) {
      if (o.cell) {
        const sheet = o.sheet ? sheetFor(o.sheet) : active;
        putGrid(sheet.getRange(o.cell).getCell(0, 0), outputGrid(o), { table: o.table });
      } else if (o.sheet) {
        const key = sheetName(o.sheet);
        stacks.set(key, [...(stacks.get(key) ?? []), o]);
      }
    }
    for (const [name, list] of stacks) {
      const sheet = sheetFor(name);
      if (madeSheets.has(name)) {
        // Clear the last run's results. Tables go first: clearing their cells would leave empty tables behind.
        sheet.tables.load("items");
        await context.sync();
        for (const t of sheet.tables.items) t.delete();
        sheet.getRange().clear();
      }
      let row = 0;
      for (const o of list) {
        const grid = outputGrid(o);
        putGrid(sheet.getRange("A1").getOffsetRange(row, 0), grid, { table: o.table });
        row += grid.values.length + 1;
      }
    }
    await context.sync();
  });
  void jolt.storage.set("madeSheets", [...madeSheets]);
  jolt.notify(`Wrote ${outputs.length} result${outputs.length === 1 ? "" : "s"} to the sheets named in the code.`, "success");
}

/** Insert plots as pictures, stacked from `anchor` down. @param {Excel.RequestContext} context @param {Excel.Range} anchor @param {string[]} plots */
async function putPlots(context, anchor, plots) {
  anchor.load("top,left");
  await context.sync();
  let top = anchor.top;
  for (const url of plots) {
    const shape = anchor.worksheet.shapes.addImage(url.slice(url.indexOf(",") + 1));
    shape.lockAspectRatio = true;
    shape.width = PLOT_WIDTH;
    shape.left = anchor.left;
    shape.top = top;
    shape.load("height");
    await context.sync();
    top += shape.height + 12;
  }
}

/** @param {string} plot */
async function insertPlot(plot) {
  await work("Inserting the plot…", async () => {
    await jolt.run(async (context) => {
      await putPlots(context, context.workbook.getSelectedRange().getCell(0, 0), [plot]);
    });
    jolt.notify("Inserted the plot at the selected cell.", "success");
  });
}

/** @param {ConsoleLine[]} lines */
const consoleText = (lines) =>
  lines.map((l) => (l.kind === "error" ? `Error: ${l.text}` : l.kind === "warning" ? `Warning: ${l.text}` : l.text)).join("\n").split("\n");

async function consoleToSheet() {
  const r = result.val;
  if (!r) return;
  await work("Writing to Excel…", async () => {
    await jolt.run(async (context) => {
      const sheet = await newSheet(context, "R Console");
      putGrid(sheet.getRange("A1"), blocksToGrid([{ title: "R console output", text: consoleText(r.console) }]));
      await context.sync();
    });
    jolt.notify("Wrote the console output to a new sheet.", "success");
  });
}

/** Everything from the last run on one new sheet: the outputs, the console and the plots. */
async function sendAll() {
  const r = result.val;
  if (!r) return;
  await work("Writing to Excel…", async () => {
    /** @type {Block[]} */
    const blocks = [{ title: `R output, ${new Date().toLocaleString()}` }];
    for (const o of r.outputs) blocks.push(...(o.blocks[0]?.title || !o.name ? o.blocks : [{ title: o.name }, ...o.blocks]));
    const text = consoleText(r.console);
    if (text.some((t) => t.trim())) blocks.push({ title: "Console", text });
    const grid = blocksToGrid(blocks);
    await jolt.run(async (context) => {
      const sheet = await newSheet(context, "R Output");
      putGrid(sheet.getRange("A1"), grid);
      await context.sync();
      if (r.plots.length) await putPlots(context, sheet.getRange("A1").getOffsetRange(grid.values.length + 1, 0), r.plots);
    });
    jolt.notify("Wrote the results to a new sheet.", "success");
  });
}

// ---- Tabs ------------------------------------------------------------------------------------------------

const activeAnalysis = () => analyses.val.find((a) => a.id === activeId.val);

/** The question for the AI: what's in "Ask your AI", or the description when that's empty. */
const askedQuestion = () => question.val.trim() || description.val.trim();

/** "Analysis 1", "Analysis 2"… whichever isn't taken. */
function nextName() {
  const taken = new Set(analyses.val.map((a) => a.name.toLowerCase()));
  for (let i = 1; ; i++) if (!taken.has(`analysis ${i}`)) return `Analysis ${i}`;
}

/** @returns {Analysis} */
const blankAnalysis = (name = nextName()) => ({ id: crypto.randomUUID(), name, description: "", sources: [], question: "", includeSample: true, code: "", explanation: "", interpretation: "" });

/** Make `a` the active tab: load its record into the states and the code box. @param {Analysis} a */
function load(a) {
  activeId.val = a.id;
  sources.val = a.sources;
  description.val = a.description ?? "";
  question.val = a.question;
  includeSample.val = a.includeSample;
  explanation.val = a.explanation ?? "";
  interpretation.val = a.interpretation ?? "";
  editingExplanation.val = false;
  editingInterpretation.val = false;
  if (editor) editor.value = a.code;
  const kept = session.get(a.id);
  result.val = kept?.result ?? null;
  promptText.val = kept?.promptText ?? "";
}

/** Keep the active tab's states before showing another tab. */
function leave() {
  capture();
  session.set(activeId.val, { result: result.val, promptText: promptText.val });
}

/** Read the active tab's sources that haven't been read yet, to show their columns. Problems are reported when the data is next used. */
async function readSources() {
  for (const s of sources.val) if (!frames.val[s.id]) await readFrame(s).catch(() => {});
}

/** @param {string} id */
async function switchTab(id) {
  const a = analyses.val.find((x) => x.id === id);
  if (!a || id === activeId.val || busy.val) return;
  leave();
  load(a);
  await save();
  await readSources();
}

/** Add a tab and show it. @param {Analysis} a */
async function openTab(a) {
  leave();
  analyses.val = [...analyses.val, a];
  load(a);
  await save();
  await readSources();
}

/** @param {string} id */
async function closeTab(id) {
  const a = analyses.val.find((x) => x.id === id);
  if (!a || analyses.val.length < 2 || busy.val) return;
  const ok = await jolt.confirm(
    `Close "${a.name}"? Its description, data choices, question and R code will be removed from this workbook. To keep a copy, click Cancel and download it first.`,
    { title: "Close analysis", ok: "Close" },
  );
  if (!ok) return;
  capture();
  const i = analyses.val.findIndex((x) => x.id === id);
  if (i < 0) return;
  const gone = new Set(analyses.val[i].sources.map((s) => s.id));
  analyses.val = analyses.val.filter((x) => x.id !== id);
  session.delete(id);
  frames.val = Object.fromEntries(Object.entries(frames.val).filter(([k]) => !gone.has(k)));
  if (id === activeId.val) load(analyses.val[Math.min(i, analyses.val.length - 1)]);
  await save();
  await readSources();
}

async function renameTab(id = activeId.val) {
  const a = analyses.val.find((x) => x.id === id);
  if (!a) return;
  const name = (await jolt.prompt("Name this analysis.", a.name, { title: "Rename analysis", ok: "Rename" }))?.trim();
  if (!name || name === a.name) return;
  analyses.val = analyses.val.map((x) => (x === a ? { ...x, name: name.slice(0, 60) } : x));
  await save();
}

/** The active tab as an analysis file. */
function activeJSON() {
  capture();
  const a = /** @type {Analysis} */ (activeAnalysis());
  return { json: serializeAnalysis(a), file: analysisFileName(a.name) };
}

/**
 * Save text as a file on the user's computer. Some Office task panes ignore downloads without telling us,
 * so the message says what to do if nothing arrives.
 * @param {string} file @param {string} text @param {string} type @param {string} fallback the button to use instead
 */
function downloadFile(file, text, type, fallback) {
  const url = URL.createObjectURL(new Blob([text], { type: `${type};charset=utf-8` }));
  const link = document.createElement("a");
  link.href = url;
  link.download = file;
  link.style.display = "none";
  /** @type {Node} */ (jolt.root ?? document.body).appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
  jolt.notify(`Downloading ${file}. If no file appears, use ${fallback} instead.`, "info");
}

function downloadTab() {
  const { json, file } = activeJSON();
  downloadFile(file, json, "application/json", "Copy as JSON");
}

async function copyTab() {
  await offerText(activeJSON().json, "Copied as JSON. Save it in a text file, or paste it into Import in another workbook.");
}

/** Open an analysis file in a new tab. @param {string} text */
async function importText(text) {
  await work("Importing…", async () => {
    const saved = parseAnalysis(text);
    const taken = new Set(analyses.val.map((a) => a.name.toLowerCase()));
    let name = saved.name;
    for (let i = 2; taken.has(name.toLowerCase()); i++) name = `${saved.name} (${i})`;
    await openTab({ ...saved, id: crypto.randomUUID(), name, sources: saved.sources.map((s) => ({ ...s, id: crypto.randomUUID() })) });
    jolt.notify(`Opened "${name}" in a new tab.`, "success");
  });
}

async function importPasted() {
  const text = await jolt.prompt("Paste the analysis JSON (from Copy as JSON or a downloaded file), or a whole HTML report.", "", { title: "Import analysis", ok: "Import" });
  if (text?.trim()) await importText(text);
}

// ---- Helpers ---------------------------------------------------------------------------------------------

/** Copy the active tab's states into its record. In place: the tab strip only shows names, which don't change here. */
function capture() {
  const a = activeAnalysis();
  if (!a) return;
  a.sources = sources.val.map(({ id, name, kind, table, sheet, address, headers }) => ({ id, name, kind, table, sheet, address, headers }));
  a.description = description.val;
  a.question = question.val;
  a.includeSample = includeSample.val;
  a.explanation = explanation.val;
  a.interpretation = interpretation.val;
  if (editor) a.code = editor.value;
}

/** Save every tab in the workbook. */
async function save() {
  capture();
  await jolt.storage.set("analyses", { tabs: analyses.val, active: activeId.val });
}

/** @param {unknown} e */
const message = (e) => (e instanceof Error ? e.message : String(e));

/**
 * Run a step with a status message, and show any error to the user.
 * @param {string} status @param {() => Promise<void>} fn
 */
async function work(status, fn) {
  if (busy.val) return;
  busy.val = status || "…";
  try {
    await fn();
  } catch (e) {
    jolt.notify(message(e), "error");
  } finally {
    busy.val = "";
  }
}

/** A cell value for the small previews. @param {unknown} v */
function show(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return Number.isInteger(v) ? v.toLocaleString() : String(Number(v.toPrecision(5)));
  return String(v);
}

// ---- The page --------------------------------------------------------------------------------------------

const off = () => !!busy.val;

/** Where a source's data is: "Table Sales" or "Sheet1!A1:D51". @param {Source} s */
const sourceWhere = (s) => (s.kind === "table" ? `Table ${s.table}` : `${s.sheet}!${s.address}`);

/** @param {Source} s */
const SourceRow = (s) => {
  const f = frames.val[s.id];
  const where = sourceWhere(s);
  return li(
    { class: "source" },
    div(
      { class: "source-head" },
      input({
        class: "rname",
        value: s.name,
        title: "The name of this data frame in R",
        spellcheck: false,
        onchange: (/** @type {Event} */ e) => {
          const el = /** @type {HTMLInputElement} */ (e.target);
          const others = sources.val.filter((x) => x.id !== s.id).map((x) => x.name);
          const name = uniqueNames([...others, rName(el.value, "data")]).at(-1) ?? s.name;
          el.value = name;
          void updateSource(s.id, { name });
        },
      }),
      button({ class: "link", title: "Remove", onclick: () => removeSource(s.id), disabled: off }, "✕"),
    ),
    small(
      { class: "muted" },
      where,
      f ? ` · ${f.rows.toLocaleString()} rows × ${f.columns.length} columns` : "",
      s.kind === "range"
        ? label(
            { class: "inline" },
            " · ",
            input({ type: "checkbox", checked: s.headers, onchange: (/** @type {Event} */ e) => updateSource(s.id, { headers: /** @type {HTMLInputElement} */ (e.target).checked }) }),
            " headers",
          )
        : "",
    ),
    f ? small({ class: "columns" }, f.columns.map((c) => `${c.name} (${c.type})`).join(", ")) : null,
  );
};

const DataPanel = () =>
  details(
    { class: "panel", open: true },
    summary(h2("1. Data")),
    p({ class: "muted" }, "Choose the data R should see. Each becomes a data frame in R, read again from the workbook each time you run."),
    () => (sources.val.length ? ul({ class: "sources" }, ...sources.val.map(SourceRow)) : p({ class: "muted" }, "No data yet.")),
    div(
      { class: "toolbar" },
      button({ onclick: addSelection, disabled: off }, "Add selected data"),
      () =>
        tables.val.length
          ? select(
              {
                disabled: off,
                onchange: (/** @type {Event} */ e) => {
                  const el = /** @type {HTMLSelectElement} */ (e.target);
                  if (el.value) void addTable(el.value);
                  el.value = "";
                },
              },
              option({ value: "" }, "Add a table…"),
              ...tables.val.map((t) => option({ value: t }, t)),
            )
          : span(),
    ),
  );

/** Always visible above the numbered sections: what this analysis is about. */
const DescriptionBox = () =>
  textarea({
    class: "description",
    rows: 2,
    "aria-label": "Description of this analysis",
    placeholder: "Description of this analysis: the question you're investigating. For example: Do sales differ between regions? (Used as the question for your AI if you leave that box blank.)",
    value: description,
    oninput: (/** @type {Event} */ e) => {
      description.val = /** @type {HTMLTextAreaElement} */ (e.target).value;
      void save();
    },
  });

const PromptPanel = () =>
  details(
    { class: "panel", open: true },
    summary(h2("2. Ask your AI")),
    p({ class: "muted" }, "Describe what you want to find out. The prompt tells your AI assistant about the data and how this R works."),
    textarea({
      class: "question",
      rows: 3,
      placeholder: () =>
        description.val.trim()
          ? `Leave blank to ask the description: ${description.val.trim()}`
          : "For example: Does advertising spend predict sales? Is the difference between regions significant?",
      value: question,
      oninput: (/** @type {Event} */ e) => {
        question.val = /** @type {HTMLTextAreaElement} */ (e.target).value;
        void save();
      },
    }),
    div(
      { class: "toolbar" },
      button({ class: "primary", onclick: copyPrompt, disabled: off }, "Copy prompt"),
      label(
        input({
          type: "checkbox",
          checked: includeSample,
          onchange: (/** @type {Event} */ e) => {
            includeSample.val = /** @type {HTMLInputElement} */ (e.target).checked;
            void save();
          },
        }),
        " Include 5 sample rows",
      ),
    ),
    () =>
      promptText.val
        ? details({ class: "prompt-preview" }, summary(small("Show the prompt")), textarea({ class: "prompt-text", readonly: true, rows: 10, value: promptText.val }))
        : span(),
  );

/** True while the explanation or interpretation is being typed or edited; otherwise it's shown formatted. */
const editingExplanation = van.state(false);
const editingInterpretation = van.state(false);

/**
 * A box for Markdown from the AI: a text box while it's empty or being edited, otherwise the formatted text.
 * @param {{ text: { val: string }, editing: { val: boolean }, title: string, placeholder: string, extract: (pasted: string) => string }} o
 *   extract picks the part of a pasted reply to keep
 */
const MarkdownBox = ({ text, editing, title, placeholder, extract }) => {
  const showBox = van.derive(() => editing.val || !text.val.trim());
  const box = textarea({
    class: "markdown-text",
    rows: 6,
    hidden: () => !showBox.val,
    placeholder,
    value: text,
    onfocus: () => (editing.val = true),
    onblur: () => (editing.val = false),
    onpaste: (/** @type {ClipboardEvent} */ e) => {
      e.preventDefault();
      box.setRangeText(extract(e.clipboardData?.getData("text") ?? ""), box.selectionStart, box.selectionEnd, "end");
      text.val = box.value;
      editing.val = false;
      void save();
    },
    oninput: () => {
      text.val = box.value;
      void save();
    },
  });
  const edit = () => {
    editing.val = true;
    setTimeout(() => box.focus());
  };
  return div(
    { class: "markdown-box" },
    div(
      { class: "markdown-head" },
      h3(title),
      () => (showBox.val ? span() : button({ class: "link", onclick: edit, title: "Edit the Markdown" }, "Edit")),
    ),
    box,
    div({ class: "markdown", hidden: showBox, innerHTML: () => markdownToHtml(text.val), ondblclick: edit }),
  );
};

/** The AI's explanation, under the code. */
const ExplanationBox = () =>
  MarkdownBox({
    text: explanation,
    editing: editingExplanation,
    title: "Why this analysis",
    placeholder: `Paste the AI's "Why this analysis" part here. It's written in Markdown and is shown formatted once you paste it or click away.`,
    // Pasted the whole reply: keep only the explanation part.
    extract: (t) => splitReply(t).explanation || t,
  });

const CodePanel = () => {
  editor = textarea({
    class: "code",
    rows: 14,
    spellcheck: false,
    placeholder: "Paste the R code from your AI here, then click Run (or press Ctrl+Enter).",
    onkeydown: (/** @type {KeyboardEvent} */ e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        void run();
      } else if (e.key === "Tab" && !e.shiftKey) {
        e.preventDefault();
        editor.setRangeText("  ", editor.selectionStart, editor.selectionEnd, "end");
      }
    },
    onpaste: (/** @type {ClipboardEvent} */ e) => {
      // A whole AI reply pasted in: keep just the code, and put its explanation in the box below.
      const text = e.clipboardData?.getData("text") ?? "";
      const reply = splitReply(text);
      if (!text.includes("```") && !reply.explanation) return;
      e.preventDefault();
      editor.setRangeText(reply.code, editor.selectionStart, editor.selectionEnd, "end");
      if (reply.explanation) {
        explanation.val = reply.explanation;
        editingExplanation.val = false;
        jolt.notify("Put the code in the box, and the AI's explanation below it.", "success");
      }
      void save();
    },
    oninput: () => void save(),
  });
  editor.value = activeAnalysis()?.code ?? "";
  return details(
    { class: "panel", open: true },
    summary(h2("3. Run R code")),
    editor,
    div(
      { class: "toolbar" },
      button({ class: "primary", onclick: run, disabled: off }, "Run"),
      select(
        {
          disabled: off,
          onchange: (/** @type {Event} */ e) => {
            const el = /** @type {HTMLSelectElement} */ (e.target);
            void loadExample(Number(el.value));
            el.value = "";
          },
        },
        option({ value: "" }, "Examples…"),
        ...EXAMPLES.map((ex, i) => option({ value: String(i) }, ex.title)),
      ),
      button({ class: "link", onclick: restart, disabled: off, title: "Start a fresh R session, clearing every variable" }, "Restart R"),
    ),
    ExplanationBox(),
  );
};

/** @param {Output} o */
const OutputCard = (o) => {
  const grid = outputGrid(o);
  const rows = grid.values.slice(0, 12);
  return li(
    { class: "output" },
    div({ class: "output-head" }, span({ class: "output-name" }, o.name || "Result"), o.sheet ? small({ class: "muted" }, ` → ${o.sheet}${o.cell ? "!" + o.cell : ""}`) : ""),
    div(
      { class: "preview" },
      table(
        ...rows.map((r, i) => {
          const isHeader = grid.marks.some((m) => m.kind === "header" && m.row === i);
          return tr(...r.map((v) => (isHeader ? th(show(v)) : td({ class: typeof v === "number" ? "num" : "" }, show(v)))));
        }),
      ),
      grid.values.length > rows.length ? small({ class: "muted" }, `…${(grid.values.length - rows.length).toLocaleString()} more rows`) : null,
    ),
    div(
      { class: "toolbar" },
      button({ onclick: () => writeOutput(o, "sheet"), disabled: off }, "New sheet"),
      button({ onclick: () => writeOutput(o, "selection"), disabled: off }, "At selected cell"),
    ),
  );
};

const Results = () => {
  const r = result.val;
  if (!r) return div();
  const shown = r.console.filter((l) => l.text !== "" || l.kind === "out");
  return details(
    { class: "panel results", open: true },
    summary(h2("4. Results"), small({ class: "muted" }, ` · ${(r.ms / 1000).toFixed(1)} s`)),
    r.error
      ? div(
          { class: "error-box" },
          p(span({ class: "error-label" }, "Error: "), r.error),
          button({ class: "primary", onclick: copyFixPrompt, disabled: off }, "Copy a prompt to fix it"),
        )
      : null,
    shown.length ? pre({ class: "console" }, ...shown.map((l) => span({ class: `line ${l.kind}` }, l.kind === "warning" ? `Warning: ${l.text}\n` : l.kind === "error" ? `Error: ${l.text}\n` : `${l.text}\n`))) : null,
    r.plots.length
      ? div({ class: "plots" }, ...r.plots.map((url) => div({ class: "plot" }, img({ src: url, alt: "R plot" }), button({ onclick: () => insertPlot(url), disabled: off }, "Insert into Excel"))))
      : null,
    r.outputs.length ? ul({ class: "outputs" }, ...r.outputs.map(OutputCard)) : null,
    div(
      { class: "toolbar" },
      button({ class: "primary", onclick: sendAll, disabled: off }, "Send everything to a new sheet"),
      shown.length ? button({ onclick: consoleToSheet, disabled: off }, "Console to sheet") : null,
    ),
    !r.outputs.length && !r.error
      ? small({ class: "muted" }, "Tip: call excel_out(x, \"Title\") in your code to send a table or model result to Excel.")
      : null,
  );
};

/** The report section shows once there are results and an interpretation. Changes only when that flips, not on every keystroke. */
const showReport = van.derive(() => !!result.val && !!interpretation.val.trim());

/** The whole analysis as one document. */
const ReportPanel = () =>
  details(
    { class: "panel", open: true },
    summary(h2("6. Report")),
    p(
      { class: "muted" },
      "The whole analysis in one document: the question, the data, the R code and why it fits, the results and the interpretation.",
    ),
    div(
      { class: "toolbar" },
      button({ class: "primary", onclick: copyReport, disabled: off }, "Copy report"),
      button({ onclick: downloadReport, disabled: off, title: "A web page with the plots in it" }, "Download .html"),
    ),
    small(
      { class: "muted" },
      "For Word: open the .html file in Word, then Save As a Word document. For PDF: open it in a web browser, then Print and choose Save as PDF. Copy report gives the text without the plots.",
    ),
  );

/** Shown once there are results: a prompt for the AI to interpret them, and a place for its answer. */
const InterpretationPanel = () =>
  details(
    { class: "panel", open: true },
    summary(h2("5. Interpretation")),
    p(
      { class: "muted" },
      "Copy a prompt containing R's output, and paste it into the same AI conversation that wrote the R code, so the AI already knows your data and question. Then paste its answer below.",
    ),
    div({ class: "toolbar" }, button({ class: "primary", onclick: copyInterpretPrompt, disabled: off }, "Copy interpretation prompt")),
    MarkdownBox({
      text: interpretation,
      editing: editingInterpretation,
      title: "The AI's interpretation",
      placeholder: "Paste the AI's interpretation here. It's shown formatted once you paste it or click away.",
      extract: (t) => t,
    }),
  );

const TabStrip = () => {
  const picker = input({
    type: "file",
    accept: ".json,.html,.htm,application/json,text/html,text/plain",
    hidden: true,
    onchange: async (/** @type {Event} */ e) => {
      const el = /** @type {HTMLInputElement} */ (e.target);
      const file = el.files?.[0];
      el.value = "";
      if (file) await importText(await file.text());
    },
  });
  return div(
    { class: "tabs" },
    () =>
      div(
        { class: "tab-list", role: "tablist" },
        ...analyses.val.map((a) =>
          div(
            { class: () => (a.id === activeId.val ? "tab active" : "tab") },
            button(
              {
                class: "tab-name",
                role: "tab",
                "aria-selected": () => String(a.id === activeId.val),
                title: "Double-click to rename",
                disabled: off,
                onclick: () => switchTab(a.id),
                ondblclick: () => renameTab(a.id),
              },
              a.name,
            ),
            analyses.val.length > 1 ? button({ class: "tab-close", title: `Close ${a.name}`, disabled: off, onclick: () => closeTab(a.id) }, "✕") : null,
          ),
        ),
        button({ class: "tab-add", title: "New analysis", disabled: off, onclick: () => openTab(blankAnalysis()) }, "+"),
      ),
    div(
      { class: "toolbar tab-tools" },
      button({ class: "link", onclick: () => renameTab(), disabled: off }, "Rename"),
      button({ class: "link", onclick: downloadTab, disabled: off, title: "Save this analysis's setup as a .json file" }, "Download"),
      button({ class: "link", onclick: copyTab, disabled: off, title: "Copy this analysis's setup as JSON text" }, "Copy as JSON"),
      select(
        {
          disabled: off,
          onchange: (/** @type {Event} */ e) => {
            const el = /** @type {HTMLSelectElement} */ (e.target);
            if (el.value === "file") picker.click();
            if (el.value === "paste") void importPasted();
            el.value = "";
          },
        },
        option({ value: "" }, "Import…"),
        option({ value: "file" }, "From a file…"),
        option({ value: "paste" }, "From pasted JSON…"),
      ),
      picker,
    ),
  );
};

const App = () =>
  main(
    { class: "app" },
    header(
      h1("R for Excel"),
      p({ class: "muted" }, "Run R on your workbook's data. Your AI assistant writes the code, and the results go back into Excel."),
      () => p({ class: `r-status ${rStatus.val.state}` }, rStatus.val.text, rStatus.val.state === "failed" ? button({ class: "link", onclick: bootR }, " Try again") : ""),
    ),
    TabStrip(),
    DescriptionBox(),
    DataPanel(),
    PromptPanel(),
    CodePanel(),
    () => (busy.val ? p({ class: "status" }, busy.val) : span()),
    Results,
    () => (result.val ? InterpretationPanel() : span()),
    () => (showReport.val ? ReportPanel() : span()),
  );

async function refreshTables() {
  const names = await readTables().then((ts) => ts.map((t) => t.name), () => []);
  if (names.join("\n") !== tables.val.join("\n")) tables.val = names;
}

/** Workbooks saved before tabs kept their one analysis under these keys. */
const OLD_KEYS = ["sources", "question", "includeSample", "code"];

export async function autoRun() {
  const stored = /** @type {{ tabs?: Analysis[], active?: string } | undefined} */ (jolt.storage.get("analyses"));
  const migrate = !stored?.tabs?.length;
  if (stored?.tabs?.length) {
    analyses.val = stored.tabs;
  } else {
    const old = Object.fromEntries(OLD_KEYS.map((k) => [k, jolt.storage.get(k)]));
    const a = blankAnalysis("Analysis 1");
    if (Array.isArray(old.sources)) a.sources = old.sources;
    if (typeof old.question === "string") a.question = old.question;
    if (typeof old.includeSample === "boolean") a.includeSample = old.includeSample;
    if (typeof old.code === "string") a.code = old.code;
    analyses.val = [a];
  }
  load(analyses.val.find((a) => a.id === stored?.active) ?? analyses.val[0]);
  const made = jolt.storage.get("madeSheets");
  if (Array.isArray(made)) for (const m of made) madeSheets.add(String(m));
  van.add(/** @type {Element} */ (jolt.$("#app")), App());
  if (migrate) {
    await save();
    for (const k of OLD_KEYS) if (jolt.storage.get(k) !== undefined) await jolt.storage.remove(k);
  }
  // Start R straight away: it takes a while the first time.
  void bootR();
  await refreshTables();
  await readSources();
}
