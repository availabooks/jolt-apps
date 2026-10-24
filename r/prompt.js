// @ts-check
import { serialToISO } from "./data.js";

/**
 * Prompts for the user's own AI assistant. They describe the data (names, types and a summary of each
 * column, plus a few rows if the user allows it) and the rules of this R environment, so the code that comes
 * back runs here without changes.
 *
 * @typedef {import("./data.js").Frame} Frame
 * @typedef {import("./data.js").Column} Column
 */

const SAMPLE_ROWS = 5;
/** The two headings the AI is asked to reply under, so its reply can be split into code and explanation. */
const CODE_HEADING = "1. R code";
const WHY_HEADING = "2. Why this analysis";
const MAX_LEVELS = 12;

/**
 * The prompt for a new analysis.
 * @param {{ frames: Frame[], question: string, includeSample: boolean, rVersion?: string }} p
 */
export function buildPrompt({ frames, question, includeSample, rVersion }) {
  return [
    "Write R code for the analysis below. It will run in \"R for Excel\", an Excel add-in that runs R in the browser, and the results go back into Excel.",
    "",
    "## What I want to find out",
    question.trim() || "(Describe the analysis you want here.)",
    "",
    describeData(frames, includeSample),
    environmentRules(rVersion),
  ].join("\n");
}

/**
 * The prompt for fixing code that failed.
 * @param {{ frames: Frame[], code: string, error: string, consoleTail: string, includeSample: boolean, rVersion?: string }} p
 */
export function buildFixPrompt({ frames, code, error, consoleTail, includeSample, rVersion }) {
  return [
    "This R code failed in \"R for Excel\", an Excel add-in that runs R in the browser. Please fix it and return the complete corrected code.",
    "",
    "## The code",
    "```r",
    code.trim(),
    "```",
    "",
    "## The error",
    "```",
    error.trim(),
    "```",
    ...(consoleTail.trim() ? ["", "## Output before the error (last lines)", "```", consoleTail.trim(), "```"] : []),
    "",
    describeData(frames, includeSample),
    environmentRules(rVersion),
  ].join("\n");
}

/** Most of the console to put in an interpretation prompt, and most rows of each result. */
const MAX_CONSOLE_LINES = 300;
const MAX_RESULT_ROWS = 60;

/**
 * The prompt that asks the AI to interpret a run's results. It's meant for the same conversation that wrote
 * the code, so it doesn't repeat the data or the code, only the question and what R produced.
 * @param {{ question: string, consoleLines: string[], results: { name: string, rows: unknown[][] }[], plots: number, error: string | null }} p
 */
export function buildInterpretPrompt({ question, consoleLines, results, plots, error }) {
  const lines = [
    "Here is the output from running the R code you wrote above. Please write a plain-English interpretation of the results, focused on answering my original question:",
    "",
    `> ${question.trim().replace(/\n/g, "\n> ") || "(the question from earlier in this conversation)"}`,
    "",
    "In your interpretation:",
    "- Start with a direct answer to the question, then explain the numbers that support it (estimates, test statistics, p-values, confidence intervals, R² and so on) in everyday language.",
    "- Include warnings about how the results should and shouldn't be interpreted. For example: whether the method's assumptions hold, the sample size, outliers or missing data, statistical versus practical significance, and correlation versus causation.",
    "- Write for a business-statistics student, in Markdown, without repeating the R code.",
    "",
  ];
  if (error) lines.push("Note: the run stopped with this error, so the output may be incomplete:", "```", error.trim(), "```", "");
  const shown = consoleLines.length > MAX_CONSOLE_LINES ? consoleLines.slice(-MAX_CONSOLE_LINES) : consoleLines;
  lines.push("## R console output");
  if (shown.length < consoleLines.length) lines.push(`(The last ${MAX_CONSOLE_LINES} of ${consoleLines.length} lines.)`);
  lines.push("```", shown.join("\n").trim() || "(nothing was printed)", "```", "");
  for (const r of results) {
    lines.push(`## Result sent to Excel: ${r.name || "Result"}`, "```");
    for (const row of r.rows.slice(0, MAX_RESULT_ROWS)) lines.push(row.map(cellText).join("\t").trimEnd());
    if (r.rows.length > MAX_RESULT_ROWS) lines.push(`… ${r.rows.length - MAX_RESULT_ROWS} more rows`);
    lines.push("```", "");
  }
  if (plots) lines.push(`R also drew ${plots} plot${plots === 1 ? "" : "s"}, which ${plots === 1 ? "isn't" : "aren't"} included here.`);
  return lines.join("\n").trimEnd() + "\n";
}

/** @param {unknown} v */
const cellText = (v) => (v === null || v === undefined ? "" : typeof v === "number" ? fmt(v) : String(v));

/** @param {Frame[]} frames @param {boolean} includeSample */
export function describeData(frames, includeSample) {
  if (!frames.length) {
    return "## The data\nNo Excel data has been loaded yet, so use one of R's built-in datasets (such as mtcars) for now.\n";
  }
  const lines = ["## The data", "These data frames already exist in R, read from the Excel workbook. Use them directly; don't read files.", ""];
  for (const f of frames) {
    lines.push(`### \`${f.name}\`: ${f.rows.toLocaleString("en-US")} rows × ${f.columns.length} columns`);
    lines.push("| Column | Type | Summary |", "| --- | --- | --- |");
    for (const c of f.columns) {
      const name = c.label && c.label !== c.name ? `\`${c.name}\` (Excel header "${c.label}")` : `\`${c.name}\``;
      lines.push(`| ${name} | ${rType(c)} | ${summarize(c).replace(/\|/g, "\\|")} |`);
    }
    if (includeSample && f.rows) {
      const n = Math.min(SAMPLE_ROWS, f.rows);
      lines.push("", `First ${n} rows:`, "```", sampleRows(f, n), "```");
    }
    lines.push("");
  }
  return lines.join("\n");
}

/** @param {string} [rVersion] */
function environmentRules(rVersion) {
  return [
    "## The R environment",
    `- R ${rVersion ?? "4.x"} compiled to WebAssembly (webR). Base R, stats, graphics and the built-in datasets are available.`,
    "- Other CRAN packages are installed automatically when the code uses library(pkg) or pkg::f, but prefer base R where it does the job, because each package is a download. Packages with system dependencies may not be available.",
    "- There is no file system, internet access or interactive input: no read.csv(), setwd(), readline() or View().",
    "- Text columns are character, not factor: use factor() where the analysis needs it.",
    "- Dates are Date columns. Missing or error cells are NA, so pass na.rm = TRUE or use complete cases where needed.",
    "",
    "## How results get back to Excel",
    "- Everything printed appears in the app's console, as in the R console. Top-level expressions are printed automatically.",
    "- Plots (base graphics or ggplot2) are captured as images that the user can insert into Excel.",
    "- To put a result into Excel cells, call `excel_out(x, \"Title\")` with the object to send. It's already defined. It handles data frames, matrices, tables, named vectors, test results (t.test, chisq.test, cor.test, etc.), lm (laid out like Excel's Analysis ToolPak regression output), glm, aov, anova and TukeyHSD. Anything else is written as printed text.",
    "- Optional arguments: `excel_out(x, \"Title\", sheet = \"Results\", cell = \"A1\")` writes it to that sheet straight after the run, and `table = TRUE` makes a data frame into an Excel Table.",
    "- Call excel_out() for each key result the user will want to keep (usually 1 to 4), not for every intermediate value.",
    "",
    "## What to send back",
    "Reply in exactly two parts, under these two headings, written exactly like this:",
    "",
    `### ${CODE_HEADING}`,
    "One complete ```r code block that runs from top to bottom, with short comments that explain each step for a business-statistics student. Put nothing else in this part.",
    "",
    `### ${WHY_HEADING}`,
    "In Markdown (not inside a code block), for a business-statistics student: why this analysis is appropriate for the question and this data, the assumptions it makes and how the code checks them, and what to look for in the output to interpret the results.",
  ].join("\n");
}

/** @param {Column} c */
export function rType(c) {
  return { number: "numeric", date: "Date", datetime: "POSIXct", logical: "logical", text: "character" }[c.type];
}

/** One line about a column: range and mean for numbers, levels for text with few values, NA count. @param {Column} c */
export function summarize(c) {
  const present = c.values.filter((v) => v !== null);
  const na = c.values.length - present.length;
  const missing = na ? `; ${na} NA` : "";
  if (!present.length) return "all NA";
  if (c.type === "number") {
    const nums = /** @type {number[]} */ (present);
    const min = Math.min(...nums);
    const max = Math.max(...nums);
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    const distinct = new Set(nums).size;
    const few = distinct <= 6 ? `; ${distinct} distinct values` : "";
    return `min ${fmt(min)}, mean ${fmt(mean)}, max ${fmt(max)}${few}${missing}`;
  }
  if (c.type === "date" || c.type === "datetime") {
    const nums = /** @type {number[]} */ (present);
    const withTime = c.type === "datetime";
    return `${serialToISO(Math.min(...nums), withTime)} to ${serialToISO(Math.max(...nums), withTime)}${missing}`;
  }
  if (c.type === "logical") {
    const t = present.filter((v) => v === true).length;
    return `${t} TRUE, ${present.length - t} FALSE${missing}`;
  }
  /** @type {Map<string, number>} */
  const counts = new Map();
  for (const v of present) counts.set(String(v), (counts.get(String(v)) ?? 0) + 1);
  if (counts.size <= MAX_LEVELS) {
    const levels = [...counts].sort((a, b) => b[1] - a[1]).map(([k, n]) => `"${k}" (${n})`);
    return `${counts.size} values: ${levels.join(", ")}${missing}`;
  }
  return `${counts.size.toLocaleString("en-US")} distinct values${missing}`;
}

/** @param {Frame} f @param {number} n */
function sampleRows(f, n) {
  const header = f.columns.map((c) => c.name);
  const rows = Array.from({ length: n }, (_, r) =>
    f.columns.map((c) => {
      const v = c.values[r];
      if (v === null || v === undefined) return "NA";
      if (c.type === "date" || c.type === "datetime") return serialToISO(/** @type {number} */ (v), c.type === "datetime");
      if (c.type === "number") return fmt(/** @type {number} */ (v));
      return String(v);
    }),
  );
  const widths = header.map((h, i) => Math.min(24, Math.max(h.length, ...rows.map((r) => r[i].length))));
  /** @param {string[]} cells */
  const line = (cells) => cells.map((s, i) => (s.length > widths[i] ? s.slice(0, widths[i] - 1) + "…" : s.padEnd(widths[i]))).join("  ").trimEnd();
  return [line(header), ...rows.map(line)].join("\n");
}

/** Numbers to at most 6 significant digits, without exponent noise. @param {number} n */
export function fmt(n) {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  return String(Number(n.toPrecision(6)));
}

/**
 * Split an AI reply into its code and its explanation, at the "Why this analysis" heading the prompt asks
 * for. The code is taken only from the part before that heading, so R snippets in the explanation stay
 * there. Without the heading, the explanation is "" and the code is extracted from the whole reply.
 * @param {string} text
 */
export function splitReply(text) {
  const heading = /^[ \t]{0,3}#{1,6}[ \t]*(?:\*\*)?(?:\d+[.)][ \t]*)?(?:\*\*)?Why this analysis\b.*$/im.exec(text);
  if (!heading) return { code: extractCode(text), explanation: "" };
  let explanation = text.slice(heading.index + heading[0].length).trim();
  // An explanation sent as one ```markdown block: use what's inside it.
  const wrapped = /^```[ \t]*(?:markdown|md)?[ \t]*\n([\s\S]*?)\n```\s*$/i.exec(explanation);
  if (wrapped) explanation = wrapped[1].trim();
  const before = text.slice(0, heading.index);
  return { code: before.includes("```") ? extractCode(before) : "", explanation };
}

/**
 * Take the R code out of an AI reply: the ```r blocks if there are any, otherwise everything.
 * @param {string} text
 */
export function extractCode(text) {
  const blocks = [...text.matchAll(/```[ \t]*([A-Za-z]*)[^\n]*\n([\s\S]*?)```/g)];
  if (!blocks.length) return text;
  const r = blocks.filter((m) => /^r$/i.test(m[1]));
  return (r.length ? r : blocks).map((m) => m[2].trimEnd()).join("\n\n") + "\n";
}
