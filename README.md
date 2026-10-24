# JOLT example apps

Three small, complete Excel apps built with [JOLT](https://jolt.availabooks.com/). Each one is a few hundred lines of plain JavaScript and HTML, with no build step. They're published on [JAM](https://jolt.availabooks.com/jam/), the JOLT app market.

| App | Install code | What it does |
| --- | --- | --- |
| [Data Cleanup](data-cleanup/) | `jam:jolt/data-cleanup` | Finds problems in a block of data and fixes each with one click, with Undo: extra spaces, numbers stored as text, blank and duplicate rows, and values that don't fit their column. |
| [Workbook Map](workbook-map/) | `jam:jolt/workbook-map` | Lists every sheet (hidden ones too), table, PivotTable, chart and named range, so you can click to any of them. It can also add a Contents sheet of links. |
| [Formula Audit](formula-audit/) | `jam:jolt/formula-audit` | Finds error values, formulas that break the pattern around them, numbers typed into formulas, links to other files and volatile functions, and can write a report sheet. |
| [R for Excel](r/) (preview) | not yet published | Runs R ([webR](https://docs.r-wasm.org/webr/)) on data from the workbook. It builds a prompt that tells your AI assistant about your data, runs the R code the AI writes, and writes the tables and plots back to sheets. Each analysis has its own tab, and you can download one as a file and import it into another workbook. |

## Try one

1. Install JOLT or JET in Excel. See [Install](https://jolt.availabooks.com/docs/install.html).
2. Open JET with the **Run** button on the Home tab.
3. Paste an install code into **Enter a code** and click the play button.
4. Optionally, click **Review with your AI** to have your own AI assistant check the code first.
5. Click **Approve & run**.

## Make your own

Each folder is a complete JOLT project:

| File | What's in it |
| --- | --- |
| `jolt.json` | Name, version, description, permissions, and where to start |
| `index.html` | The page: one element that the app fills in |
| `styles.css` | The look, using JET's theme colours so dark mode works |
| `main.js` | The UI (built with [VanJS](https://vanjs.org/), which JOLT bundles) and every call to Excel, through `jolt.run` |
| `analyze.js` | The logic, as pure functions that are easy to read and test |

R for Excel uses `data.js` and `prompt.js` for its pure logic instead of `analyze.js`. It also has:
- `rsession.js`, which starts and runs R;
- `jolt.R`, the R helpers, including `excel_out()`;
- `vendor/webr.mjs`, webR's loader, copied in because installed apps can't import code from other sites.

To build on one:
1. Open JADE in Excel (**Build** on the Home tab) and start a project.
2. Copy the folder's files in, then change them.
3. With VS Code and JUICE, open the folder and run "JUICE: Connect to Office".

To publish your version on JAM, change `id` in `jolt.json` to `<your GitHub name>/<app>` and follow [Publish your own app](https://jolt.availabooks.com/jam/#publish-your-own-app).

These apps are maintained in the JOLT repository and copied here for each release. Please report issues here.
