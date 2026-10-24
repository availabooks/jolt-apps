# Handoff: R-powered "Analysis ToolPak" JOLT app

This is a summary of an earlier planning conversation. Nothing has been built or tested yet. This folder (`jolt-apps/r`) is empty apart from this file.

## Goal

Build a JOLT app for Excel that runs real R in the browser, using **webR**, and writes the results onto a sheet in the same layout as Excel's **Analysis ToolPak** (Data > Data Analysis). The intended users are students in a stats or business course. They'll compare its output with the real ToolPak, so the numbers must match Excel's exactly.

## Context

- **JOLT** (https://jolt.availabooks.com/) apps are static sites: `jolt.json`, `index.html`, `styles.css` and ES-module JavaScript, with no build step. Follow the sibling apps in this repo (`../formula-audit`, `../data-cleanup`, `../workbook-map`):
  - `main.js` holds the UI, built with VanJS (bundled by JOLT), and makes every Excel call through `jolt.run`.
  - Pure logic goes in a separate module, like `analyze.js` in the sibling apps.
  - `styles.css` uses JET's theme colours so dark mode works.
  - The repo `README.md` describes the file roles.
- **webR** is Posit's WebAssembly build of R. Import it from `https://webr.r-wasm.org/latest/webr.mjs`; pin a specific version rather than `latest` for a release. Prebuilt CRAN packages are available through `webR.installPackages()`.

## Key technical decisions and risks

1. **Use `ChannelType.PostMessage`.** webR's default channel needs `SharedArrayBuffer`, which needs COOP/COEP headers. An app stored in a document, or running in an Office task pane, probably can't set those headers. PostMessage mode works without them. It loses interrupting running code and interactive `readline()`, neither of which this app needs. **Check this first, inside a real Excel task pane** (WebView2 on Windows, WKWebView on Mac, and Excel on the web).
2. **First load is about 20–30 MB.** Call `webR.init()` as soon as the pane opens, show a "Starting R…" status, and rely on the browser cache after that.
3. **Data flow:** read the range with Office.js, then `webR.objs.globalEnv.bind('x', values)`, then `evalR(...)`, then `.toJs()` or `.toArray()`, then write the values back with Office.js. R has no access to the workbook itself.

## Output design

Each R analysis function should return a **list of display blocks**, not raw R objects. One generic JavaScript renderer then lays out any tool's output:

```r
list(
  list(title = "SUMMARY OUTPUT"),
  list(title = "Regression Statistics", rows = list(c("Multiple R", ...), ...)),
  list(title = "ANOVA", header = c("", "df", "SS", "MS", "F", "Significance F"), rows = ...),
  list(header = c("", "Coefficients", "Standard Error", "t Stat", "P-value", "Lower 95%", "Upper 95%"), rows = ...)
)
```

The renderer writes each block to cells, leaves a blank row between blocks, and copies the ToolPak's styling: italic column headers with a bottom border, a heavier top border on tables, and sensible number formats. Copy the ToolPak dialog options too:
- input range
- "Labels in first row"
- confidence level, where the tool has one
- output to a range, a new worksheet or a new workbook

## Tools to cover

These are the ToolPak tools: Descriptive Statistics, t-Test (paired; two-sample, equal variances; two-sample, unequal variances), z-Test, F-Test, ANOVA (single factor; two-factor with and without replication), Correlation, Covariance, Regression, Histogram, Moving Average, Exponential Smoothing, Rank and Percentile, Sampling, and Random Number Generation.

**Suggested first slice:** Regression and Descriptive Statistics. Between them they exercise most of the layout patterns.

## Exact-match pitfalls (Excel vs R defaults)

- **Descriptive Statistics:** Excel's Kurtosis and Skewness are the bias-corrected `KURT()` and `SKEW()`, not R's or e1071's defaults. Mode shows `#N/A` when no value repeats. Standard deviation and variance are sample versions. Confidence Level is the t-based half-width.
- **t-Tests:** Excel shows `P(T<=t)` and `t Critical` for both one tail and two tails. In the unequal-variance test, the df Excel displays and the df it computes with may differ by rounding. Confirm this against Excel.
- **Histogram:** Excel's bins are "≤ upper bound" plus a final "More" bin, and it chooses its own default bins when none are given.
- **Regression:** with "Constant is Zero", Excel's R² definition changes. `lm(y ~ 0 + x)` should match, but verify it.

**Testing:** for each tool, keep a fixture dataset and the real ToolPak's output for it. Assert that the R version matches to about 1e-9.

## Nice-to-haves beyond the ToolPak

- Build the ToolPak's charts (residual plots, line fit, histogram) as **native Excel charts** through Office.js, not R images, so users can edit them.
- An option to write live formulas (for example `LINEST`) instead of static values, where Excel has an equivalent.
- Extra analyses in the same layout: logistic regression, Tukey HSD, chi-square tests and robust standard errors.

---

## Status (2026-10-22): first build of "R for Excel"

The direction changed from this plan: the user's own AI writes the R code, so the app's main jobs are to expose the data, build the prompt, run the code and write the results back. ToolPak-style dialogs are deferred. ToolPak-style *layout* survives in `excel_out()` (for example, `lm` is written like the ToolPak's regression output).

| File | Role |
| --- | --- |
| `main.js` | UI and every Excel call: analysis tabs, data sources (Tables or ranges), prompt copy, run, write outputs and plots, download/import of analyses |
| `rsession.js` | webR start-up (PostMessage channel), binding data frames, auto-installing packages, `captureR` |
| `data.js` | Pure: Excel values → typed columns, R names, blocks → cell grid, sheet names, analysis files (`serializeAnalysis`/`parseAnalysis`) |
| `report.js` | Pure: the full analysis report as Markdown and as self-contained HTML |
| `markdown.js` | Pure: a small, escaping Markdown → HTML converter for the AI's explanation, interpretation and the report |
| `prompt.js` | Pure: the AI prompt, the "fix this error" prompt, and extracting code from a pasted AI reply |
| `examples.js` | Starter analyses that adapt to the first data frame's columns, or use mtcars |
| `jolt.R` | `excel_out()`, object → blocks (htest, lm, glm, aov, TukeyHSD, tables…), a JSON writer |
| `vendor/webr.mjs` | webR 0.6.0 loader (67 KB). R 4.6.0 itself downloads from webr.r-wasm.org |

**Verified in headless Edge** (served from localhost, so cross-origin to the webR CDN, and not cross-origin isolated):
- The vendored loader starts R in PostMessage mode.
- Data frames bind with NA and dates correct.
- All 9 examples run on sample data and on mtcars.
- Plots are captured.
- `library(x)` auto-installs, and `install.packages()` works through webR's shim.
- Errors keep the earlier output.

The UI was also smoke-tested against a fake Excel context.

**Not yet verified, so do this next:**
1. In a real JET task pane: Windows desktop (WebView2), Mac (WKWebView, which needs OffscreenCanvas for plots, in Safari 16.4 and later) and Excel on the web.
2. In the Office web view: the Clipboard API with its fallback, and `shapes.addImage`.
3. Writing to real sheets: styling, Excel Tables, and clearing app-made sheets on rerun.
4. Exact-match checks against the ToolPak for regression and t-tests.
5. The JAM review/fingerprint flow with a vendored file and runtime fetches from the webR CDN.

## Update (2026-10-23): analysis tabs and files

- Each analysis (description, data sources, question, sample option, code) is a tab. The description shows above the sections and is the AI question when "Ask your AI" is empty. All tabs are saved in one `jolt.storage` key, `analyses` (`{ tabs, active }`). Workbooks saved earlier are moved over on first open, and the old keys (`sources`, `question`, `includeSample`, `code`) are removed.
- The prompt asks the AI for two parts, headed "1. R code" and "2. Why this analysis" (Markdown). Pasting a reply into the code box splits it (`splitReply` in prompt.js), and the explanation shows formatted under the code (`markdown.js`, a small escaping Markdown renderer, because installed apps can't import libraries).
- Results are section "4. Results". Section "5. Interpretation" appears once there are results. It copies a prompt with R's output (`buildInterpretPrompt`) for the same AI conversation, and has a Markdown box for the answer, which is saved per tab but not in analysis files.
- Section "6. Report" appears once there are results and an interpretation. `report.js` builds the report as Markdown (Copy report) and as self-contained HTML with the plots embedded (Download .html, the only download). The HTML also carries the analysis file in a comment (`analysisComment` in data.js), and Import accepts a report: `parseAnalysis` reads the comment. Word and PDF come from the HTML: open it in Word and Save As, or print it to PDF from a browser.
- All tabs share one R session. Results are kept per tab for the session only.
- Download writes `<name>.r-analysis.json` (`format: "r-for-excel-analysis"`, `version: 1`). It holds the setup only, with no results and no source ids. Import opens the file in a new tab. "Copy as JSON" and "From pasted JSON" are fallbacks for task panes that block downloads or file pickers.
- Sheets the app creates have gridlines off. Data cells get light grey borders in place of gridlines.
- Not yet verified in a real task pane: tab switching, migration, and whether downloads work in WebView2, WKWebView and Excel on the web.
