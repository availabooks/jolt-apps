// @ts-check
import { ChannelType, WebR } from "./vendor/webr.mjs";
import { normalizeOutputs } from "./data.js";

/**
 * The R session, run by webR (R compiled to WebAssembly) in a web worker.
 *
 * webR's loader (vendor/webr.mjs) is copied into the project because installed JOLT apps can't import code
 * from other sites. The loader then downloads R itself (about 30 MB the first time, cached after that)
 * and any packages from webr.r-wasm.org and repo.r-wasm.org.
 *
 * The PostMessage channel is used because the default one needs SharedArrayBuffer, which needs
 * cross-origin isolation that an Office task pane doesn't have. It means a running script can't be
 * interrupted, so "Restart R" is the way out of an endless loop.
 */

export const WEBR_VERSION = "0.6.0";
const BASE_URL = `https://webr.r-wasm.org/v${WEBR_VERSION}/`;

/** Packages that come with R and never need installing. */
const BASE_PACKAGES = new Set("base compiler datasets graphics grDevices grid methods parallel splines stats stats4 tcltk tools utils webr".split(" "));

/**
 * @typedef {import("./data.js").Frame} Frame
 * @typedef {import("./data.js").Output} Output
 * @typedef {{ kind: "out" | "err" | "message" | "warning" | "error", text: string }} ConsoleLine
 * @typedef {{ console: ConsoleLine[], images: ImageBitmap[], outputs: Output[], error: string | null, ms: number }} RunResult
 */

/** @type {WebR | null} */
let webR = null;
/** @type {Promise<string> | null} */
let starting = null;

/**
 * Start R (once) and load the helpers in jolt.R. Resolves to R's version, like "4.5.1".
 * @param {string} helpers the text of jolt.R
 * @param {{ baseUrl?: string }} [options] baseUrl is only for running outside JOLT, such as in tests
 */
export function startR(helpers, options = {}) {
  starting ??= (async () => {
    const r = new WebR({ baseUrl: options.baseUrl ?? BASE_URL, channelType: ChannelType.PostMessage, interactive: false });
    await r.init();
    await r.evalRVoid(helpers);
    // Let install.packages() work, for code written with it.
    await r.evalRVoid("try(webr::shim_install(), silent = TRUE)");
    webR = r;
    return r.versionR;
  })();
  starting.catch(() => (starting = null));
  return starting;
}

/** Stop R and start a fresh session, clearing every variable. @param {string} helpers */
export async function restartR(helpers) {
  const old = webR;
  webR = null;
  starting = null;
  old?.close();
  return startR(helpers);
}

/** @returns {WebR} */
function session() {
  if (!webR) throw new Error("R hasn't started yet.");
  return webR;
}

/**
 * Put data frames in R's global environment, replacing any of the same name.
 * Each column goes across as a plain array (null is NA) in its own variable, because webR turns a JS
 * object into a data frame only when it's already shaped like one. jolt.R's .jolt_frame() puts them together.
 * @param {Frame[]} frames
 */
export async function bindFrames(frames) {
  const r = session();
  const env = r.objs.globalEnv;
  for (const f of frames) {
    await env.bind(".jolt_name", f.name);
    await env.bind(".jolt_names", f.columns.map((c) => c.name));
    await env.bind(".jolt_types", f.columns.map((c) => c.type));
    for (const [i, c] of f.columns.entries()) await env.bind(`.jolt_c${i + 1}`, c.values);
    await r.evalRVoid(".jolt_frame()");
  }
}

/**
 * Packages the code loads that aren't installed yet: library(x), require(x), requireNamespace("x") and x::f.
 * @param {string} code
 */
export function packagesUsed(code) {
  const names = new Set();
  const noComments = code.replace(/#.*$/gm, "");
  for (const m of noComments.matchAll(/\b(?:library|require|requireNamespace)\s*\(\s*["']?([A-Za-z][A-Za-z0-9.]*)/g)) names.add(m[1]);
  for (const m of noComments.matchAll(/\b([A-Za-z][A-Za-z0-9.]*):::?[A-Za-z.]/g)) names.add(m[1]);
  return [...names].filter((n) => !BASE_PACKAGES.has(n));
}

/**
 * Install any packages the code uses that aren't installed yet.
 * @param {string} code @param {(message: string) => void} [status]
 * @returns {Promise<string[]>} the packages installed
 */
export async function ensurePackages(code, status) {
  const wanted = packagesUsed(code);
  if (!wanted.length) return [];
  const r = session();
  const quoted = wanted.map((w) => JSON.stringify(w)).join(",");
  const missing = (await r.evalRString(`paste(setdiff(c(${quoted}), rownames(installed.packages())), collapse = ",")`)).split(",").filter(Boolean);
  if (!missing.length) return [];
  status?.(`Installing ${missing.join(", ")}…`);
  await r.installPackages(missing, { quiet: true });
  return missing;
}

/**
 * Run R code the way the R console would: each top-level result is printed, and plots are captured
 * as images. Errors stop the run but keep the output so far.
 * @param {string} code
 * @param {{ width?: number, height?: number }} [plotSize]
 * @returns {Promise<RunResult>}
 */
export async function runCode(code, plotSize = {}) {
  const r = session();
  const started = performance.now();
  await r.evalRVoid(".jolt_reset()");
  const shelter = await new r.Shelter();
  try {
    const capture = await shelter.captureR(code, {
      withAutoprint: true,
      captureStreams: true,
      captureConditions: true,
      throwJsException: false,
      captureGraphics: typeof OffscreenCanvas === "undefined" ? false : { width: plotSize.width ?? 640, height: plotSize.height ?? 440 },
    });
    /** @type {ConsoleLine[]} */
    const lines = [];
    let error = null;
    for (const item of capture.output) {
      if (item.type === "stdout") lines.push({ kind: "out", text: item.data });
      else if (item.type === "stderr") lines.push({ kind: "err", text: item.data });
      else {
        const text = await conditionMessage(item.data);
        lines.push({ kind: /** @type {ConsoleLine["kind"]} */ (item.type), text });
        if (item.type === "error") error = text;
      }
    }
    const outputs = normalizeOutputs(JSON.parse(await r.evalRString(".jolt_outputs_json()")));
    return { console: lines, images: capture.images ?? [], outputs, error, ms: Math.round(performance.now() - started) };
  } finally {
    await shelter.purge();
  }
}

/**
 * The message of an R condition (an error, warning or message), with the call it came from for errors.
 * @param {any} condition an RObject proxy for the condition
 */
async function conditionMessage(condition) {
  try {
    const message = await (await condition.get("message")).toString();
    return String(message).replace(/\n$/, "");
  } catch {
    return String(await condition.toString?.() ?? "Unknown error");
  }
}

/** Is R running? */
export const isReady = () => webR !== null;
