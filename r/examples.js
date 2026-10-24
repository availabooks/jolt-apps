// @ts-check

/**
 * Starter code for common business-statistics analyses. Each example uses the columns of the first data
 * frame when it has suitable ones (numbers, or text with a few groups), and R's built-in mtcars otherwise.
 *
 * @typedef {import("./data.js").Frame} Frame
 * @typedef {{ df: string, num: string[], group: string | null, group2: string | null }} Picks
 * @typedef {{ title: string, code: (p: Picks) => string }} Example
 */

/** @param {Frame | undefined} frame @returns {Picks} */
export function pickColumns(frame) {
  const num = frame?.columns.filter((c) => c.type === "number").map((c) => c.name) ?? [];
  const groups =
    frame?.columns
      .filter((c) => {
        if (c.type !== "text" && c.type !== "logical") return false;
        const n = new Set(c.values.filter((v) => v !== null)).size;
        return n >= 2 && n <= 12;
      })
      .map((c) => c.name) ?? [];
  if (frame && num.length >= 2) return { df: frame.name, num, group: groups[0] ?? null, group2: groups[1] ?? null };
  return { df: "mtcars", num: ["mpg", "wt", "hp"], group: "cyl", group2: "gear" };
}

/** `df$col`, with backticks when the name needs them. @param {string} df @param {string} col */
const col = (df, col) => `${df}$${/^[A-Za-z.][A-Za-z0-9._]*$/.test(col) ? col : "`" + col + "`"}`;

/** @type {Example[]} */
export const EXAMPLES = [
  {
    title: "Descriptive statistics",
    code: ({ df, num }) => `# Mean, median, spread and shape of each numeric column
nums <- ${df}[sapply(${df}, is.numeric)]
stats <- data.frame(
  Mean = sapply(nums, mean, na.rm = TRUE),
  Median = sapply(nums, median, na.rm = TRUE),
  SD = sapply(nums, sd, na.rm = TRUE),
  Min = sapply(nums, min, na.rm = TRUE),
  Max = sapply(nums, max, na.rm = TRUE),
  N = sapply(nums, function(x) sum(!is.na(x)))
)
stats
excel_out(stats, "Descriptive statistics")
boxplot(nums, main = "Spread of each column")
`,
  },
  {
    title: "Histogram",
    code: ({ df, num }) => `# The shape of the distribution of ${num[0]}
x <- ${col(df, num[0])}
hist(x, main = "Histogram of ${num[0]}", xlab = "${num[0]}", col = "steelblue", border = "white")
bins <- table(cut(x, breaks = pretty(x)))
excel_out(bins, "Frequency table for ${num[0]}")
`,
  },
  {
    title: "Confidence interval & one-sample t-test",
    code: ({ df, num }) => `# Is the mean of ${num[0]} different from a target value?
target <- ${"mean(" + col(df, num[0]) + ", na.rm = TRUE)"}  # change to the value you want to test against
result <- t.test(${col(df, num[0])}, mu = target, conf.level = 0.95)
result
excel_out(result, "One-sample t-test for ${num[0]}")
`,
  },
  {
    title: "Compare two groups (t-test)",
    code: ({ df, num, group }) =>
      group
        ? `# Does ${num[0]} differ between two groups of ${group}?
groups <- unique(na.omit(${col(df, group)}))[1:2]
two <- ${df}[${col(df, group)} %in% groups, ]
result <- t.test(${num[0]} ~ ${group}, data = two)  # Welch test: doesn't assume equal variances
result
excel_out(result, "Two-sample t-test")
boxplot(${num[0]} ~ ${group}, data = two, main = "${num[0]} by ${group}")
`
        : `# Are ${num[0]} and ${num[1]} different on average? (paired: each row is one case)
result <- t.test(${col(df, num[0])}, ${col(df, num[1])}, paired = TRUE)
result
excel_out(result, "Paired t-test")
`,
  },
  {
    title: "Correlation",
    code: ({ df }) => `# How strongly do the numeric columns move together?
nums <- ${df}[sapply(${df}, is.numeric)]
r <- cor(nums, use = "pairwise.complete.obs")
round(r, 3)
excel_out(r, "Correlation matrix")
pairs(nums, main = "Scatterplot matrix")
`,
  },
  {
    title: "Simple regression",
    code: ({ df, num }) => `# Predict ${num[0]} from ${num[1]}. Swap the two names to predict the other way round.
model <- lm(${num[0]} ~ ${num[1]}, data = ${df})
summary(model)
excel_out(model, "Regression of ${num[0]} on ${num[1]}")
plot(${num[0]} ~ ${num[1]}, data = ${df}, main = "${num[0]} vs ${num[1]}", pch = 19, col = "steelblue")
abline(model, col = "firebrick", lwd = 2)
plot(fitted(model), resid(model), main = "Residuals vs fitted", xlab = "Fitted", ylab = "Residual")
abline(h = 0, lty = 2)
`,
  },
  {
    title: "Multiple regression",
    code: ({ df, num }) => `# Predict ${num[0]} from all the other numeric columns. Change ${num[0]} to predict a different one.
nums <- ${df}[sapply(${df}, is.numeric)]
model <- lm(${num[0]} ~ ., data = nums)
summary(model)
excel_out(model, "Multiple regression for ${num[0]}")
par(mfrow = c(2, 2))
plot(model)
`,
  },
  {
    title: "One-way ANOVA",
    code: ({ df, num, group }) =>
      group
        ? `# Does the mean of ${num[0]} differ across the groups of ${group}?
model <- aov(${num[0]} ~ factor(${group}), data = ${df})
summary(model)
excel_out(model, "ANOVA: ${num[0]} by ${group}")
excel_out(TukeyHSD(model), "Which groups differ")
boxplot(${num[0]} ~ ${group}, data = ${df}, main = "${num[0]} by ${group}")
`
        : `# ANOVA needs a column of group labels. Add one, or try it on R's built-in data:
model <- aov(mpg ~ factor(cyl), data = mtcars)
summary(model)
excel_out(model, "ANOVA: mpg by cylinders")
`,
  },
  {
    title: "Chi-square test of independence",
    code: ({ df, group, group2 }) => `# Are two categorical columns related?
${
  group && group2
    ? `counts <- table(${col(df, group)}, ${col(df, group2)})`
    : "# This needs two columns of categories, so it uses R's built-in mtcars data:\ncounts <- table(cylinders = mtcars$cyl, gears = mtcars$gear)"
}
counts
result <- chisq.test(counts)
result
excel_out(counts, "Counts")
excel_out(result, "Chi-square test")
`,
  },
];
