// @ts-check
import { serialToISO } from "./data.js";
import { markdownToHtml } from "./markdown.js";
import { fmt, rType, summarize } from "./prompt.js";

/**
 * The full report of one analysis: question, data, code, why this analysis, results and interpretation.
 * Written as Markdown (for Copy report) and as a self-contained HTML page with the plots in it
 * (which opens in Word, and prints to PDF from a browser). Pure: no Office or webR calls.
 *
 * @typedef {import("./data.js").Frame} Frame
 * @typedef {import("./data.js").Output} Output
 * @typedef {import("./data.js").Block} Block
 * @typedef {{
 *   title: string, ranAt: Date, rVersion: string, ms: number, question: string, description: string,
 *   sources: { name: string, where: string, frame?: Frame }[], code: string, explanation: string,
 *   error: string | null, consoleLines: string[], outputs: Output[], plots: number, interpretation: string
 * }} ReportInput
 */

/** Marks where plot n goes, so reportHtml can put the image there. @param {number} n */
const plotMarker = (n) => `@@PLOT ${n}@@`;

/**
 * The report as Markdown.
 * @param {ReportInput} r
 * @param {{ plot?: (n: number) => string }} [options] the line written for plot n
 */
export function buildReport(r, { plot = (n) => `*Plot ${n} is included in the HTML version of this report.*` } = {}) {
  const out = [`# ${r.title}`, ""];
  const ran = `${r.ranAt.toLocaleString()} · R ${r.rVersion || "4.x"} · ran in ${(r.ms / 1000).toFixed(1)} s`;
  out.push(`*${ran}*`, "");

  out.push("## Question", "", r.question.trim() || "*(No question was given.)*", "");
  if (r.description.trim() && r.description.trim() !== r.question.trim()) out.push("**Description:** " + r.description.trim(), "");

  out.push("## Data", "");
  if (!r.sources.length) out.push("No workbook data was used.", "");
  for (const s of r.sources) {
    const size = s.frame ? ` · ${s.frame.rows.toLocaleString()} rows × ${s.frame.columns.length} columns` : "";
    out.push(`### \`${s.name}\``, "", `${s.where}${size}`, "");
    if (s.frame?.columns.length) {
      out.push("| Column | Excel header | Type | Summary |", "| --- | --- | --- | --- |");
      for (const c of s.frame.columns) out.push(row([`\`${c.name}\``, c.label, rType(c), summarize(c)]));
      out.push("");
    }
  }

  out.push("## R code", "", fence(r.code.trim(), "r"), "");
  if (r.explanation.trim()) out.push("## Why this analysis", "", shiftHeadings(r.explanation.trim(), 2), "");

  out.push("## Results", "");
  if (r.error) out.push(`**The run stopped with an error:** ${r.error}`, "");
  const printed = r.consoleLines.join("\n").trim();
  if (printed) out.push("### R console output", "", fence(printed), "");
  for (const o of r.outputs) {
    out.push(`### ${o.name || "Result"}`, "");
    const blocks = o.blocks[0]?.title === o.name ? [{ ...o.blocks[0], title: undefined }, ...o.blocks.slice(1)] : o.blocks;
    out.push(blocksToMarkdown(blocks), "");
  }
  if (r.plots) {
    out.push("### Plots", "");
    for (let n = 1; n <= r.plots; n++) out.push(plot(n), "");
  }

  if (r.interpretation.trim()) out.push("## Interpretation", "", shiftHeadings(r.interpretation.trim(), 2), "");
  out.push("---", "", "*Made with R for Excel.*", "");
  return out.join("\n");
}

/**
 * excel_out() blocks as Markdown: titles in bold, tables as Markdown tables, text as code blocks.
 * @param {Block[]} blocks
 */
export function blocksToMarkdown(blocks) {
  const parts = [];
  for (const b of blocks) {
    if (b.title) parts.push(`**${b.title}**`);
    if (b.text) {
      parts.push(fence(b.text.join("\n")));
      continue;
    }
    const cols = b.cols ?? [];
    if (!cols.length) continue;
    const width = Math.max(cols.length, b.header?.length ?? 0);
    const n = Math.max(...cols.map((c) => c.length));
    /** @param {unknown} v @param {number} c */
    const cell = (v, c) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "number") return /yyyy/.test(b.fmt?.[c] ?? "") ? serialToISO(v, /hh/.test(b.fmt?.[c] ?? "")) : fmt(v);
      return String(v);
    };
    const lines = [row(Array.from({ length: width }, (_, c) => b.header?.[c] ?? "")), row(Array(width).fill("---"))];
    for (let i = 0; i < n; i++) lines.push(row(Array.from({ length: width }, (_, c) => cell(cols[c]?.[i], c))));
    parts.push(lines.join("\n"));
  }
  return parts.join("\n\n");
}

/** A Markdown table row, with | in cells escaped and line breaks flattened. @param {string[]} cells */
const row = (cells) => `| ${cells.map((c) => c.replace(/\|/g, "\\|").replace(/\s*\n\s*/g, " ")).join(" | ")} |`;

/** A fenced code block, with a longer fence if the text has ``` in it. @param {string} text @param {string} [lang] */
function fence(text, lang = "") {
  const ticks = text.includes("```") ? "````" : "```";
  return `${ticks}${lang}\n${text}\n${ticks}`;
}

/** Move Markdown headings down `by` levels (at most ######), leaving code blocks alone. @param {string} md @param {number} by */
export function shiftHeadings(md, by) {
  let inFence = "";
  return md
    .split("\n")
    .map((line) => {
      const f = line.match(/^\s*(```+|~~~+)/);
      if (f) {
        if (!inFence) inFence = f[1];
        else if (line.trim().startsWith(inFence)) inFence = "";
        return line;
      }
      if (inFence) return line;
      return line.replace(/^(\s{0,3})(#{1,6})(?=\s)/, (_, sp, h) => sp + "#".repeat(Math.min(6, h.length + by)));
    })
    .join("\n");
}

/** @param {string} s */
const escapeHtml = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * The report as a complete HTML page, with the plots embedded as images.
 * @param {ReportInput} r @param {string[]} plots PNG data URLs, in order
 * @param {string} [setup] the analysis as an HTML comment (analysisComment in data.js), so the report can be imported
 */
export function reportHtml(r, plots, setup = "") {
  const body = markdownToHtml(buildReport(r, { plot: plotMarker })).replace(/<p>@@PLOT (\d+)@@<\/p>/g, (m, n) => {
    const url = plots[Number(n) - 1] ?? "";
    if (!/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(url)) return "";
    return `<figure><img src="${url}" alt="Plot ${n}"><figcaption>Plot ${n}</figcaption></figure>`;
  });
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(r.title)}</title>
${setup}
<style>
  body { max-width: 50rem; margin: 2rem auto; padding: 0 1rem; color: #1f1f1f; font: 11pt/1.5 Georgia, "Times New Roman", serif; }
  h1, h2, h3, h4, h5, h6 { font-family: "Segoe UI", Calibri, Arial, sans-serif; line-height: 1.25; }
  h1 { font-size: 20pt; margin-bottom: 0.2rem; }
  h2 { font-size: 15pt; margin-top: 2rem; padding-bottom: 0.2rem; border-bottom: 1px solid #ccc; }
  h3 { font-size: 12pt; margin-top: 1.4rem; }
  h4, h5, h6 { font-size: 11pt; }
  code { font: 9.5pt Consolas, "Courier New", monospace; background: #f3f3f3; padding: 0 0.2em; }
  pre { background: #f3f3f3; padding: 0.6rem 0.8rem; overflow-x: auto; white-space: pre-wrap; }
  pre code { padding: 0; }
  table { border-collapse: collapse; margin: 0.5rem 0 1rem; font-size: 10pt; }
  th, td { border: 1px solid #bbb; padding: 0.2rem 0.5rem; text-align: left; vertical-align: top; }
  th { background: #f3f3f3; }
  blockquote { margin-left: 0; padding-left: 0.8rem; border-left: 3px solid #ccc; color: #555; }
  figure { margin: 1rem 0; }
  img { max-width: 100%; height: auto; }
  figcaption { font-size: 9pt; color: #555; }
  hr { border: none; border-top: 1px solid #ccc; margin-top: 2rem; }
  @media print {
    body { margin: 0; max-width: none; }
    h2, h3 { break-after: avoid; }
    pre, table, figure { break-inside: avoid; }
  }
</style>
</head>
<body>
${body}
</body>
</html>
`;
}
