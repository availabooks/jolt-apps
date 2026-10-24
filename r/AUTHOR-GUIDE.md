# R for Excel: how the tool works (reference for writing reader instructions)

This document describes **R for Excel**, an Excel add-in that runs the R statistics language inside Excel. It is written for an author, human or AI, who is writing a statistics textbook for college students that uses the tool. Use it to write accurate step-by-step instructions for readers.

**Conventions:**
- Button labels, menu items and messages appear in **bold** or in "quotes", exactly as they appear on screen. Use them word for word in reader instructions.
- Don't describe buttons or features that aren't listed here.
- Where this document says something hasn't been confirmed, keep the instructions general or check with the book's human author.

---

## 1. What the tool is, in one paragraph

R for Excel is a panel that opens beside an Excel workbook. Its numbered sections follow one analysis from question to report:
1. **Data:** the student chooses data in the workbook (an Excel Table or a block of cells).
2. **Ask your AI:** the student writes a question in plain English and copies a ready-made **prompt** that describes the data to an AI assistant, such as ChatGPT, Claude, Copilot or Gemini.
3. **Run R code:** the student pastes the AI's reply back into the panel. The tool separates the R code from the AI's explanation of **why this analysis** fits, and the student clicks **Run**.
4. **Results:** R's output, tables and plots appear, and can be sent to Excel sheets.
5. **Interpretation:** the student copies a second prompt with R's output into the same AI conversation, and pastes back a plain-English interpretation.
6. **Report:** the whole analysis is copied or downloaded as one document, which can be made into a Word or PDF file.

A student can keep several analyses at once, each on its own tab, and save each one in the workbook or as a file.

The student doesn't need to install R or know R syntax, although they can type their own R code at any time. R runs inside the add-in itself, using **webR**. The data never leaves the computer to be analyzed. The only data shared with the AI is what the copied prompt contains (see §5.3).

**Another way to work, useful for a textbook:** the student can skip the AI completely and paste R code printed in the book into the code box. See §10.

---

## 2. Getting the tool

- R for Excel is a **JOLT** app. JOLT is a platform for small Excel add-ins (https://jolt.availabooks.com/). Students first install JOLT or **JET**, its app runner, in Excel. The instructions are at https://jolt.availabooks.com/docs/install.html.
- JOLT apps are opened in JET with the **Run** button on Excel's **Home** tab. A student then pastes an **install code** into **Enter a code**, clicks the play button, and clicks **Approve & run**.
- **Status:** R for Excel is a preview and **doesn't have a public install code yet**. Leave a placeholder for the install code, and confirm it with the book's human author before publication.
- The tool works in Excel for Windows, Excel for Mac and Excel on the web. Plots on a Mac need a recent version of macOS and Safari (Safari 16.4 or later). Without one, R code runs but plots aren't captured.
- **First start:** R downloads about **30 MB** the first time, so it can take a minute. The panel shows "Starting R… The first time, this downloads about 30 MB." When it's ready, it shows "R 4.x.x is ready." in green. Later starts are much faster, because the download is cached. If start-up fails, the panel says "R didn't start: …" and shows a **Try again** link.
- An internet connection is needed to start R, and the first time a package is used.

---

## 3. The layout of the panel

From top to bottom:

1. **Header:** the title "R for Excel", a one-line description, and R's status line ("Starting R…" / "R … is ready." / "R didn't start…").
2. **Tab strip:** one tab per *analysis* (see §8). The **+** button adds a tab. Under the tabs is a small toolbar: **Rename**, **Download**, **Copy as JSON** and an **Import…** menu.
3. **Description:** a box for what this analysis is about, usually the question being investigated. It stays visible above the numbered sections. Its placeholder text reads "Description of this analysis: the question you're investigating. …".
4. **1. Data:** choose the data R will see.
5. **2. Ask your AI:** type a question and copy the prompt.
6. **3. Run R code:** the code box, **Run**, an **Examples…** menu and **Restart R**.
7. **Status line:** a grey line that appears while something is happening, such as "Reading the data…", "Installing ggplot2…", "Running…" or "Writing to Excel…". Buttons are greyed out until it finishes.
8. **4. Results:** appears after a run (see §7).
9. **5. Interpretation:** appears once there are results (see §7.2).
10. **6. Report:** appears once there are results and an interpretation (see §7.3).

Each numbered section can be collapsed by clicking its heading.

---

## 4. Section 1: Data

### 4.1 Adding data

There are two ways:

- **Add selected data:** select cells in the workbook first, then click this button.
  - **One cell selected:** the tool uses the whole block of filled cells around it, like pressing Ctrl+A in Excel. Students can click any cell inside their data.
  - **A range selected:** the tool uses the selection. If whole columns are selected, the empty part below the data is ignored.
  - **The selection is exactly an Excel Table:** the Table is added *as a Table* (see 4.2).
- **Add a table…:** a drop-down list of every Excel Table in the workbook. It appears only when the workbook has at least one Table. Choosing one adds it.

Each piece of data becomes a **data frame** in R. A data frame is R's name for a table of rows and columns.

### 4.2 Tables versus ranges

- **Table:** the source is linked to the Excel Table by its name. If the student adds rows to the Table later, the next run includes them. **Recommend Excel Tables in the book.** They grow automatically and always have headers. In Excel, a block of data becomes a Table with **Insert > Table** or **Ctrl+T**.
- **Range:** a fixed block of cells, such as `Sheet1!A1:D51`. New rows typed below it are *not* included.

### 4.3 The data list

Each added source shows:
- **A name box:** the data frame's name in R, which can be edited. It's taken from the Table name or sheet name and made into a valid R name. Spaces and symbols become `_`, a name starting with a digit gets an `X` in front, and names are made unique by adding `_2`, `_3` and so on.
  - **Tip for the book:** tell students to rename the data frame to the name the book's code uses, such as `sales`. Click the name box, type the new name and press Enter or click away.
- **Where it comes from** ("Table Sales" or "Sheet1!A1:D51") and its size, for example "· 50 rows × 4 columns".
- **A headers checkbox** (ranges only). Tick it if the first row holds column names. The tool guesses: it ticks the box when the first row is all text and some row below contains a number.
- **The list of columns** with the type R will use for each, for example `Price (number), Region (text), Date (date)`.
- **✕** removes the source. It removes the source from the tool only, not the data in the workbook.

### 4.4 How Excel values become R values

The book may need to explain these rules when students see unexpected results.

| In Excel | In R |
| --- | --- |
| A column where every filled cell is a number | numeric |
| Numbers stored as text, like `"1,234"`, `"$5.00"`, `"(12)"`, `"45%"` | numeric: 1234, 5, −12, 0.45 |
| A numeric column where most cells have a date format | `Date`, or date-time (`POSIXct`) if there are times |
| A column of only TRUE/FALSE | logical |
| Anything else, including a column that mixes text and numbers | character (text). **Not** factor: the R code must use `factor()` where the analysis needs categories |
| Blank cells and Excel errors (`#N/A`, `#DIV/0!`, `#VALUE!`, …) | `NA` (missing) |

**Column names:** taken from the header row and changed into valid R names. For example, `Ad Spend ($)` becomes `Ad_Spend`, `2024 Sales` becomes `X2024_Sales`, and a blank header becomes `V1`, `V2`, …. The prompt tells the AI both names.

**Limits:**
- A single data source can have at most **500,000 cells**.
- A source with no data rows is rejected.

### 4.5 Data is read fresh

The data is **read from the workbook again every time the student clicks Run or Copy prompt**. If they fix a typo in a cell, the next run uses the corrected value. Nothing needs to be re-added.

### 4.6 Error messages about data

- "There's no table called "X" any more. Remove … from the data list." The Table was deleted or renamed.
- "There's no sheet called "X" any more. …" The sheet was deleted or renamed.
- "… has N cells. The limit is 500,000."
- "… has no rows of data."
- "The selection has no data in it."

---

## 5. Section 2: Ask your AI

### 5.1 Steps

1. Type the question in the box, in plain language. For example: *"Does advertising spend predict sales? Is the difference between regions significant?"*
   - **If this box is left empty, the analysis's description (§3) is used as the question.** The box's placeholder then shows "Leave blank to ask the description: …". A student who wrote a good description can skip this step.
2. Leave **Include 5 sample rows** ticked, unless the data is confidential (see 5.3).
3. Click **Copy prompt**. The message "Prompt copied. Paste it into your AI assistant, then paste its code into the box below." appears.
4. Paste the prompt into an AI assistant, such as a chat window, and send it.
5. Copy the AI's whole reply and paste it into the code box in Section 3. The tool splits it into the code and the explanation (§6.1).

If the add-in can't use the clipboard, it says "Couldn't reach the clipboard. The prompt is selected below: press Ctrl+C (⌘C on a Mac)." The prompt then appears, already selected, in a box under **Show the prompt**. The student can open **Show the prompt** at any time to read exactly what will be sent.

### 5.2 What the prompt contains

The prompt is written for the AI. It has these sections:
- **"What I want to find out":** the student's question, or the description if the question box is empty.
- **"The data":** for each data frame:
  - its R name and size;
  - a table of columns, with each column's R name, its original Excel header, its R type and a one-line summary:
    - numbers: min, mean and max, and the number of distinct values when there are 6 or fewer;
    - dates: the date range;
    - TRUE/FALSE: the counts;
    - text: up to 12 category values with their counts, or else the number of distinct values;
    - the number of NAs, if there are any;
  - the first 5 rows, if **Include 5 sample rows** is ticked.
- **"The R environment":** rules so the AI's code runs here first time: no files, no internet, no interactive input, text is character not factor, missing values are NA, and packages install automatically but base R is preferred.
- **"How results get back to Excel":** explains `excel_out()` (see §9).
- **"What to send back":** the AI is told to reply in **two clearly labelled parts**, under these exact headings:
  - **"1. R code":** one complete R code block, with short comments that explain each step **for a business-statistics student**, and nothing else.
  - **"2. Why this analysis":** in Markdown, why this analysis suits the question and the data, the assumptions it makes and how the code checks them, and what to look for in the output when interpreting it.

  The "Copy a prompt to fix it" prompt (§7) asks for the same two parts.

If no data has been added, the prompt tells the AI to use one of R's built-in datasets, such as `mtcars`.

### 5.3 Privacy

The prompt is the only thing that leaves the computer.
- **Unticking "Include 5 sample rows"** removes the actual rows.
- **The column summaries are always included,** even with the box unticked. That means column names, min, mean and max, and up to 12 category values with their counts. For confidential data, the student should rename sensitive columns or categories before copying, or not use an outside AI at all.

---

## 6. Section 3: Run R code

### 6.1 The code box

- **Pasting a whole AI reply:** paste the full reply, explanations and all. The tool **keeps only the R code** in the code box (the code fenced with ``` marks). If the reply has the "Why this analysis" part, that part goes into the **Why this analysis** box below (§6.1a), with the message "Put the code in the box, and the AI's explanation below it." Students don't need to select anything carefully.
- **Typing:** Tab inserts two spaces, and **Ctrl+Enter** (⌘+Enter on a Mac) runs the code.
- The code is saved with the workbook (see §8.3).

### 6.1a Why this analysis

Below the code box and its **Run** button is a box headed **Why this analysis**. It holds the AI's explanation of why the analysis is appropriate.
- **When it's empty,** it's a text box. Its placeholder reads "Paste the AI's "Why this analysis" part here. …". Students can paste the explanation, or the whole AI reply, here. Only the explanation part is kept.
- **When it's filled in,** it's shown as **formatted text**: headings, bold and italics, bullet and numbered lists, tables, and code samples. It's written in Markdown, a simple text format, but students don't need to know that.
- **To change it,** click **Edit** (or double-click the text). When the student clicks away, it's shown formatted again.
- **Mathematical notation:** formulas the AI writes in LaTeX, such as `$\bar{x}$`, appear as plain text, not as rendered maths.
- The explanation is saved with the workbook, and in downloaded analysis files.

**Use in a textbook:** the explanation is a natural prompt for reflection. For example, ask students to read it and check whether the assumptions it lists hold for their data.

### 6.2 Run

Click **Run**. The tool:
1. waits for R to finish starting, if it hasn't yet;
2. reads every data source from the workbook again and loads it into R as a data frame;
3. installs any packages the code uses that aren't installed yet: those named in `library(x)`, `require(x)` or `x::function`. The status shows "Installing x…". A download is needed the first time, and some CRAN packages aren't available in webR;
4. runs the code like the R console: each top-level expression's value is printed, and plots are captured;
5. shows the results (§7), and writes any results that the code sent to a named sheet (§9.2).

### 6.3 Examples…

A drop-down list of nine ready-made analyses. The book can refer to them by name:
1. **Descriptive statistics:** mean, median, SD, min, max and N for each numeric column, plus a box plot.
2. **Histogram:** a histogram and frequency table of the first numeric column.
3. **Confidence interval & one-sample t-test**
4. **Compare two groups (t-test):** a Welch two-sample t-test using the first text column with 2 to 12 categories, or a paired t-test if there is no such column. Includes a box plot.
5. **Correlation:** a correlation matrix and scatterplot matrix.
6. **Simple regression:** `lm` of the first numeric column on the second, with a fitted-line plot and a residual plot.
7. **Multiple regression:** the first numeric column predicted from all the other numeric columns, with diagnostic plots.
8. **One-way ANOVA:** with Tukey HSD and a box plot.
9. **Chi-square test of independence:** using two text columns.

Each example **adapts to the first data frame**: it picks that data's own numeric and category columns. If the data isn't suitable, for example it has fewer than two numeric columns, the example uses R's built-in `mtcars` data. Choosing an example **replaces the code in the box** after asking "Replace the code in the box with this example?" Every example includes comments and `excel_out()` calls.

### 6.4 Restart R

**Restart R** starts a fresh R session and clears every variable. The panel then shows "R … restarted. Every variable was cleared."

Use it when:
- **The code gets stuck,** for example in an endless loop. A running script **can't be stopped any other way**.
- **The student wants a clean slate.** R keeps variables between runs, and between tabs (§8).

R prints a technical message when it starts. The JOLT console may show "WebR is using `PostMessage` communication channel, nested R REPLs are not available." This is normal and can be ignored.

---

## 7. Results

The results appear after a run, in a section headed **4. Results** with the run time, such as "· 1.2 s". Like the other numbered sections, it can be collapsed by clicking its heading.

- **Error box:** if the code stopped with an error, a red box shows "Error: <message>" and a **Copy a prompt to fix it** button. That button copies a new prompt containing the code, the error, the last lines of output and the data description, ready to paste back into the AI. Pasting the AI's corrected code and clicking **Run** again is the normal way to fix errors. **Everything printed before the error is kept and shown.**
- **Console:** everything R printed, in a monospaced box. Warnings are prefixed "Warning:" and errors "Error:", in their own colours.
- **Plots:** each plot is shown as an image with an **Insert into Excel** button. That button inserts the plot as a picture at the selected cell, 480 points wide. Plots are pictures, not editable Excel charts.
- **Outputs:** one card for each `excel_out()` call (see §9). Each card shows the output's name, a preview of up to 12 rows, and two buttons:
  - **New sheet:** writes it to a new worksheet named after the output.
  - **At selected cell:** writes it starting at the selected cell. **This overwrites whatever is in those cells.**
- **Send everything to a new sheet:** writes all outputs, the console text and the plots to one new sheet, titled "R output, <date and time>".
- **Console to sheet:** writes the console text to a new sheet called "R Console".
- If the code made no outputs, a tip suggests calling `excel_out()`.

### 7.1 How results look in Excel

- **Sheets the tool creates have Excel's gridlines turned off.** Data cells get thin light-grey borders instead, so tables look clean.
- **Titles** are bold.
- **Column headers** are italic, with a heavier line above and a thin line below.
- **The last row of a table** has a thin line under it.
- Columns are widened to fit, and console text uses a monospaced font.
- This layout is modelled on Excel's **Analysis ToolPak** (Data > Data Analysis). In particular, a regression (`lm`) is written in the ToolPak's "SUMMARY OUTPUT" layout:
  - **Regression Statistics:** Multiple R, R Square, Adjusted R Square, Standard Error and Observations;
  - **ANOVA:** df, SS, MS, F and Significance F;
  - **coefficients:** Coefficients, Standard Error, t Stat, P-value, Lower 95% and Upper 95%.

  The book can compare it directly with the ToolPak.
- **Numbers are written at full precision, not rounded.** Students can use Excel's number formatting (Increase/Decrease Decimal) to choose how many decimals to show.
- Dates are written with a `yyyy-mm-dd` format.
- New sheet names avoid clashes: "R Output", "R Output 2", ….

### 7.2 Section 5: Interpretation

This section appears only once there are results. It helps students turn R's output into a plain-English answer.

1. Click **Copy interpretation prompt**. The message reads "Copied. Paste it into the same AI conversation that wrote the R code, then paste its answer below."
2. Paste the prompt into **the same AI conversation that wrote the R code**. That conversation already knows the data, the question and the code, so the prompt doesn't repeat them. Pasting it into a new conversation still works, but the AI then knows less about the data.
3. Copy the AI's answer and paste it into the box headed **The AI's interpretation**. Like the explanation box (§6.1a), it's shown as formatted text, and **Edit** changes it.

The prompt contains:
- the student's question, or the description if the question box is empty;
- R's console output, up to the last 300 lines;
- each result sent with `excel_out()`, up to 60 rows each;
- the number of plots, which aren't included because the AI can't see them from a prompt;
- the error message, if the run stopped with an error.

It asks the AI for a Markdown answer, written for a business-statistics student. The answer should:
- start with a direct answer to the question;
- explain the supporting numbers (estimates, test statistics, p-values, confidence intervals and R²) in everyday language;
- **include warnings about how the results should and shouldn't be interpreted**: assumptions, sample size, outliers or missing data, statistical versus practical significance, and correlation versus causation.

**Use in a textbook:** have students write their own interpretation first, then compare it with the AI's. The AI can be wrong, so the book should present its interpretation as something to check, not as the answer.

### 7.3 Section 6: Report

This section appears once there are results **and** the interpretation box has text. It produces the whole analysis as one document, ready to hand in or keep. The report contains:
1. **The tab's name as the title,** with the date and time of the run, R's version and the run time.
2. **Question:** the question, or the description if the question box is empty. The description is shown too, if it's different.
3. **Data:** for each data frame, its R name, where it comes from (Table or range), its size, and a table of its columns: R name, Excel header, type and summary.
4. **R code:** the code that actually ran in the last run, not any later edits in the code box.
5. **Why this analysis:** the AI's explanation.
6. **Results:**
   - the error, if the run stopped with one;
   - the console output;
   - each `excel_out()` result as a table, with numbers shown to 6 significant figures;
   - the plots.
7. **Interpretation:** the AI's interpretation.

The buttons are:
- **Copy report:** copies the report as Markdown text, to paste into a document or a Markdown editor. Plots aren't included; each one is listed with a note that it's in the HTML version.
- **Download .html:** a formatted web page **with the plots embedded**. The file also carries the analysis setup, hidden in the page, so the report can be imported later to rebuild the analysis (§8.2).

**Word and PDF:** the tool doesn't write Word or PDF files directly. Students make them from the .html file:
- **Word:** open the .html file in Word, then **File > Save As** and choose Word Document (.docx).
- **PDF:** open the .html file in a web browser (double-click it), then **Print** and choose **Save as PDF** as the printer.

The section shows this tip under the buttons. Downloading may be blocked in some versions of Excel (§8.2). If no file appears, use **Copy report**.

**Not yet verified:** whether Word keeps the embedded plots when it opens the .html file. If it doesn't, students can insert the plots from Excel (**Insert into Excel** in section 4) and copy them into the Word document.

---

## 8. Tabs: several analyses at once

### 8.1 What a tab holds

Each tab is one *analysis*. It holds:
- its description;
- its data sources;
- its question;
- the **Include 5 sample rows** setting;
- its R code;
- the AI's explanation (**Why this analysis**);
- the AI's interpretation (**The AI's interpretation**);
- its last results, while the panel stays open.

Clicking a tab switches to it. Tabs are useful for keeping, say, "Chapter 3 – Descriptive stats" and "Chapter 5 – Regression" side by side in one workbook.

- **Add:** click **+**. New tabs are named "Analysis 1", "Analysis 2", and so on.
- **Rename:** double-click a tab, or click **Rename** in the toolbar under the tabs.
- **Close:** click the **✕** on a tab. The tool asks first: *"Close "<name>"? Its description, data choices, question and R code will be removed from this workbook. To keep a copy, click Cancel and download it first."* The last remaining tab can't be closed.
- Tabs can't be switched while something is running.

**All tabs share one R session.** Each run loads its own tab's data, but variables created in another tab are still visible. To start clean, use **Restart R**.

### 8.2 Saving, sharing and importing an analysis

These are in the toolbar under the tabs:
- **Download:** saves the current tab's setup as a file named `<tab name>.r-analysis.json`. The file holds:
  - the name;
  - the description;
  - the data sources, saved by reference: the Table name, or the sheet and address, but not the data itself;
  - the question;
  - the sample-rows setting;
  - the R code;
  - the explanation.

  It does **not** hold results, data values or the interpretation, which depends on a run's results.
- **Copy as JSON:** copies the same content as text, to save or send when downloading isn't possible.
- **Import… → From a file…:** opens a downloaded `.r-analysis.json` file, **or a downloaded HTML report** (§7.3), **in a new tab**. It never replaces the current tab. A report rebuilds the same setup as the analysis file: description, data sources, question, sample setting, R code and explanation. It doesn't rebuild the results or the interpretation; click **Run** to get the results again. A report that has been re-saved from Word as .docx can't be imported.
- **Import… → From pasted JSON…:** the same, from pasted text: either the JSON, or the whole text of an HTML report file.

**Use in a textbook:** the author can provide a ready-made analysis file for each chapter, with the code and the expected Table names. A student imports it into the chapter's workbook and clicks **Run**. If the workbook doesn't have a Table the analysis expects, the data list shows "There's no table called …".

Downloading may not work in every version of Excel. If no file appears, use **Copy as JSON** instead. This hasn't been tested in every Excel version yet, so mention both options in the book.

### 8.3 What is kept when the workbook is saved

The following are stored **inside the workbook**, so they're kept when the student **saves** and reopens it:
- all tabs;
- each tab's description, data sources, question, sample setting, code, explanation and interpretation;
- which tab was open.

They travel with the file if it's emailed. The workbook must be saved, as with any other change.

The following are **not** kept:
- **The results (section 4) and R's variables.** After reopening, click **Run** again. The saved interpretation is kept, but section 5 only appears again once there are results.
- **Packages.** They are reinstalled automatically on first use.

**Results already written to sheets are ordinary Excel content and are kept.**

---

## 9. `excel_out()`: sending results to Excel from R code

`excel_out()` is a function the tool adds to R. Calling it in R code marks a result to send to Excel. The AI is told about it in the prompt, and book code should use it too.

### 9.1 Basic use

```r
model <- lm(Sales ~ AdSpend, data = sales)
excel_out(model, "Regression of Sales on AdSpend")
```

The output then appears as a card in **4. Results**, with **New sheet** and **At selected cell** buttons. The second argument is the title written above the result. If it's left out, the variable's name is used.

### 9.2 Writing straight to a sheet

```r
excel_out(model, "Regression", sheet = "Results")              # stacked from A1 on sheet "Results"
excel_out(stats, "Summary", sheet = "Results", cell = "H2")    # placed at H2
excel_out(clean_data, "Cleaned data", sheet = "Data2", table = TRUE)  # written as an Excel Table
```

With `sheet =`, the result is written **automatically after each run**, with no button to click.
- The sheet is created if it doesn't exist.
- Several outputs to the same sheet without `cell =` are stacked from A1, with a blank row between them.
- **If the tool created that sheet,** it's cleared before writing, so re-running replaces the old results instead of piling up.
- **If the sheet already existed** (the student's own sheet), it's *not* cleared, and results overwrite the cells they land on.
- `table = TRUE` turns a data frame into an Excel Table.

### 9.3 What it handles

| R object | How it appears in Excel |
| --- | --- |
| `lm(...)` (linear regression) | Analysis ToolPak "SUMMARY OUTPUT" layout (§7.1) |
| `summary(lm(...))` | Regression statistics plus the coefficient table |
| `t.test`, `chisq.test`, `cor.test`, `wilcox.test`, `shapiro.test`, `prop.test`, and other test results | A two-column table: data, estimates, hypothesized value, test statistic, df, **P-value**, alternative, and confidence interval bounds |
| `aov(...)` | The ANOVA table: Df, Sum Sq, Mean Sq, F value, Pr(>F) |
| `TukeyHSD(...)` | One table per factor: diff, lwr, upr, p adj |
| `glm(...)` | Model fit (observations, deviances, AIC) and the coefficient table |
| Data frames and matrices | A table. Meaningful row names become the first column |
| `table(...)` (counts, cross-tabs) | A one-way or two-way frequency table |
| Named vectors, such as `sapply(df, mean)` or `summary(x)` | A two-column label / value table |
| Single values | A label and its value |
| Lists | Each element in turn |
| Anything else | Its printed text, one line per row |

The prompt asks the AI to call `excel_out()` for the 1 to 4 key results only, not for every intermediate step.

---

## 10. Using the tool without an AI (recommended for book exercises)

The AI writes different code each time, so for exercises where every reader should see the same output, the book can print the R code and have students:
1. Put the chapter's data in an Excel Table (Ctrl+T) and give it the name the code expects (**Table Design > Table Name**).
2. Add it with **Add a table…**, and check that the data frame's name in the name box matches the code. If it doesn't, rename it.
3. Paste the book's code into the code box (section **3. Run R code**) and click **Run**.

This gives reproducible results the book can show. Use **Examples…** for quick starters, and the AI path (§5) for open-ended questions. An analysis file (§8.2) can bundle the code and setting for a chapter.

Without an AI, the **Why this analysis** box stays empty, and that's fine. Students can also **type their own interpretation** into section 5's box instead of pasting the AI's. Doing that unlocks section 6, so a report can be a student's own work: the book's code, R's results and the student's interpretation.

When writing code for the book, follow the same rules the prompt gives the AI:
- base R only where possible;
- no `read.csv()`, `setwd()`, `readline()` or `View()`;
- use `factor()` for category columns;
- handle `NA` with `na.rm = TRUE` or complete cases;
- use `excel_out()` for the results to keep.

`library(ggplot2)` and other CRAN packages work, but each one needs a download the first time it's used.

---

## 11. What the tool can't do (limits to mention to readers)

- **No files or internet from R:** R can't read CSV files from disk or the web. Data must come from the workbook through Section 1, or from R's built-in datasets, such as `mtcars`, `iris` and `ToothGrowth`.
- **No interactive input:** `readline()`, `menu()`, `browser()` and `View()` don't work.
- **A running script can't be interrupted.** Use **Restart R**.
- **Some CRAN packages aren't available** in webR, especially those that need system libraries.
- **Plots are pictures,** not native Excel charts.
- **Data limit:** 500,000 cells per data source.
- **Speed:** the first start is slow because of the 30 MB download. Later starts are quick. Large data or heavy models are slower than desktop R.
- **No formulas:** results are values, not live formulas. If the data changes, click **Run** again.

---

## 12. Quick troubleshooting (for a reader-facing appendix)

| What the student sees | What to do |
| --- | --- |
| "Starting R…" for a long time on first use | Wait: it's a one-time 30 MB download. Check the internet connection. |
| "R didn't start: …" | Click **Try again**. Check the internet connection. |
| A red **Error:** box after Run | Click **Copy a prompt to fix it**, paste it into the AI, then paste the corrected code and click **Run** again. |
| `object 'xyz' not found` | The data frame's name in Section 1 doesn't match the name in the code. Rename it in the name box, or change the code. |
| A number column shows as text in the column list | Some cells in that column contain text, such as "n/a" or "-". Clear or fix those cells. |
| "There's no table called …" | The Table was renamed or deleted. Remove the source (✕) and add it again. |
| Run seems frozen | Click **Restart R**, then run again. |
| "Couldn't reach the clipboard…" | Press Ctrl+C (⌘C) to copy the selected text shown under **Show the prompt** in section 2. This applies to every Copy button, including **Copy interpretation prompt** and **Copy report**. |
| The **Why this analysis** box stayed empty after pasting the AI's reply | The AI didn't use the "2. Why this analysis" heading. Copy its explanation and paste it into the box by hand. |
| Section 5 (Interpretation) isn't showing | It appears only after a run. Click **Run**. |
| Section 6 (Report) isn't showing | It appears only after a run *and* once the interpretation box has text. |
| Downloading an analysis didn't produce a file | Use **Copy as JSON** and paste it into a text file, or use **Import… → From pasted JSON…** in the other workbook. |
| Downloading the report didn't produce a file | Use **Copy report** and paste it into a document. Plots aren't included that way; insert them from section 4. |
| The Word version of the report has no plots | Insert the plots into Excel with **Insert into Excel** (section 4), then copy them into the Word document. |
| Results gone after reopening the workbook | Expected: click **Run** again. Results written to sheets are still there. |

---

## 13. Notes for the author

- **Preview software:** R for Excel is version 0.1.0 and doesn't have a public install code yet. Its behaviour in real Excel task panes on every platform hasn't been fully verified, particularly downloading and importing files, plots on a Mac, and whether Word keeps the plots when it opens a report. Keep version-specific claims modest, and confirm screenshots against the released version.
- **Interface text:** the labels in this guide match the current version. If screenshots differ, the screenshots win.
- **Output layout:** the regression layout matches the Analysis ToolPak's labels. Students comparing numbers with the ToolPak should see the same values. For other tests, R's conventions apply. For example, `t.test` uses the Welch test by default, unless the code sets `var.equal = TRUE`, while the ToolPak offers both. The book may want to point this out.
