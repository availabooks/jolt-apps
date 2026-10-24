# Helpers that R for Excel loads into R at start-up.
#
# - .jolt_frame() turns the columns main.js sends into a data frame.
# - excel_out() is for the user's code: it marks an object to be written to Excel.
# - .jolt_blocks() turns an R object into "blocks" (a title, a header row and columns, or lines of text),
#   which main.js writes to cells. .jolt_json() sends them to JavaScript as JSON.

.jolt <- new.env()
.jolt$outputs <- list()

# Build a data frame from the variables rsession.js binds: .jolt_name, .jolt_names, .jolt_types
# and one .jolt_c<i> per column. Dates arrive as Excel serial numbers.
.jolt_frame <- function() {
  g <- globalenv()
  x <- list(name = g$.jolt_name, names = g$.jolt_names, types = g$.jolt_types)
  col_vars <- paste0(".jolt_c", seq_along(x$names))
  on.exit(rm(list = intersect(c(".jolt_name", ".jolt_names", ".jolt_types", col_vars), ls(g, all.names = TRUE)), envir = g))
  cols <- lapply(seq_along(x$names), function(i) {
    v <- get(col_vars[i], envir = g)
    switch(x$types[i],
      date = as.Date(v, origin = "1899-12-30"),
      datetime = as.POSIXct(round((v - 25569) * 86400), origin = "1970-01-01", tz = "UTC"),
      number = as.numeric(v),
      logical = as.logical(v),
      as.character(v)
    )
  })
  names(cols) <- x$names
  df <- as.data.frame(cols, stringsAsFactors = FALSE, check.names = FALSE)
  assign(x$name, df, envir = globalenv())
  invisible(df)
}

#' Send an R object to Excel.
#'
#' @param x a data frame, matrix, table, vector, test result (t.test, chisq.test...), model (lm, glm, aov) or anything printable
#' @param name the title shown above it in Excel
#' @param sheet write to this worksheet straight after the run (created if missing)
#' @param cell top-left cell for `sheet`, such as "A1"
#' @param table TRUE to make it an Excel Table (data frames only)
excel_out <- function(x, name = deparse(substitute(x)), sheet = NULL, cell = NULL, table = FALSE) {
  name <- paste(name, collapse = " ")
  .jolt$outputs[[length(.jolt$outputs) + 1L]] <- list(
    name = name,
    sheet = sheet,
    cell = cell,
    table = isTRUE(table),
    blocks = .jolt_blocks(x, name)
  )
  invisible(x)
}

.jolt_reset <- function() .jolt$outputs <- list()

.jolt_outputs_json <- function() .jolt_json(.jolt$outputs)

# ---- R objects to blocks --------------------------------------------------------------------------------

.jolt_blocks <- function(x, title = NULL) {
  if (inherits(x, "htest")) return(.jolt_htest(x))
  if (inherits(x, "lm") && !inherits(x, c("glm", "aov", "mlm"))) return(.jolt_lm(x))
  if (inherits(x, "glm")) return(.jolt_glm(x))
  if (inherits(x, "aov") && !inherits(x, "aovlist")) return(.jolt_table_block(summary(x)[[1]], title))
  if (inherits(x, "summary.lm")) return(.jolt_summary_lm(x))
  if (inherits(x, "TukeyHSD")) {
    return(unlist(lapply(names(x), function(n) .jolt_table_block(x[[n]], paste("Tukey HSD:", n))), recursive = FALSE))
  }
  if (is.data.frame(x) || is.matrix(x)) return(.jolt_table_block(x, title))
  if (is.table(x) || inherits(x, "summaryDefault")) {
    if (length(dim(x)) == 2) return(.jolt_table_block(unclass(x), title, corner = paste(names(dimnames(x)), collapse = " \\ ")))
    if (length(dim(x)) <= 1) return(.jolt_vector_block(setNames(as.vector(x), names(x)), title))
  }
  if (is.atomic(x) && is.null(dim(x)) && length(x) > 0) return(.jolt_vector_block(x, title))
  if (is.list(x) && !is.object(x) && length(x) > 0 && length(x) <= 20) {
    nm <- names(x)
    if (is.null(nm)) nm <- paste0("[[", seq_along(x), "]]")
    return(unlist(lapply(seq_along(x), function(i) .jolt_blocks(x[[i]], paste(c(title, nm[i]), collapse = ": "))), recursive = FALSE))
  }
  list(list(title = title, text = capture.output(print(x))))
}

# A data frame or matrix as a table. Row names become the first column unless they're just 1, 2, 3...
.jolt_table_block <- function(x, title = NULL, corner = "") {
  rn <- rownames(x)
  if (!is.null(rn)) rn <- trimws(rn)
  keep_rn <- !is.null(rn) && !(is.data.frame(x) && .row_names_info(x) <= 0) && !identical(rn, as.character(seq_len(nrow(x))))
  if (is.matrix(x)) {
    cn <- colnames(x)
    if (is.null(cn)) cn <- paste0("[,", seq_len(ncol(x)), "]")
    x <- lapply(seq_len(ncol(x)), function(j) x[, j])
    names(x) <- cn
  }
  cols <- lapply(x, .jolt_column)
  header <- names(x)
  if (is.null(header)) header <- rep("", length(cols))
  if (keep_rn) {
    cols <- c(list(list(values = rn, fmt = "")), cols)
    header <- c(corner, header)
  }
  list(list(
    title = title,
    header = as.character(header),
    cols = unname(lapply(cols, `[[`, "values")),
    fmt = unname(vapply(cols, `[[`, "", "fmt"))
  ))
}

# One column's values plus the Excel number format it needs.
.jolt_column <- function(v) {
  if (inherits(v, "Date")) return(list(values = as.numeric(v) + 25569, fmt = "yyyy-mm-dd"))
  if (inherits(v, "POSIXt")) return(list(values = as.numeric(v) / 86400 + 25569, fmt = "yyyy-mm-dd hh:mm:ss"))
  if (is.factor(v)) return(list(values = as.character(v), fmt = ""))
  if (is.list(v)) return(list(values = vapply(v, function(e) paste(format(e), collapse = ", "), ""), fmt = ""))
  if (is.numeric(v) || is.logical(v) || is.character(v)) return(list(values = as.vector(v), fmt = ""))
  list(values = format(v), fmt = "")
}

.jolt_vector_block <- function(x, title = NULL) {
  nm <- names(x)
  if (is.factor(x)) x <- as.character(x)
  col <- .jolt_column(x)
  if (is.null(nm)) {
    if (length(x) == 1) return(list(list(cols = list(title %||% "", col$values), fmt = c("", col$fmt))))
    return(list(list(title = title, cols = list(col$values), fmt = col$fmt)))
  }
  list(list(title = title, cols = list(nm, col$values), fmt = c("", col$fmt)))
}

`%||%` <- function(a, b) if (is.null(a)) b else a

# A two-column table of labels and values, for results like test statistics.
.jolt_pairs <- function(title, labels, values) {
  list(title = title, header = c("", "Value"), cols = list(labels, values), fmt = c("", ""))
}

.jolt_htest <- function(x) {
  labels <- character()
  values <- list()
  add <- function(l, v) {
    labels <<- c(labels, l)
    values[[length(values) + 1L]] <<- v
  }
  if (!is.null(x$data.name)) add("Data", x$data.name)
  if (!is.null(x$estimate)) for (i in seq_along(x$estimate)) add(names(x$estimate)[i] %||% "Estimate", unname(x$estimate[i]))
  if (!is.null(x$null.value)) for (i in seq_along(x$null.value)) add(paste("Hypothesized", names(x$null.value)[i] %||% "value"), unname(x$null.value[i]))
  if (!is.null(x$statistic)) add(names(x$statistic) %||% "Statistic", unname(x$statistic))
  if (!is.null(x$parameter)) for (i in seq_along(x$parameter)) add(names(x$parameter)[i] %||% "Parameter", unname(x$parameter[i]))
  if (!is.null(x$p.value)) add("P-value", x$p.value)
  if (!is.null(x$alternative)) add("Alternative", x$alternative)
  if (!is.null(x$conf.int)) {
    level <- attr(x$conf.int, "conf.level")
    pct <- if (is.null(level)) "" else paste0(" ", format(100 * level), "%")
    add(paste0("Lower", pct), x$conf.int[1])
    add(paste0("Upper", pct), x$conf.int[2])
  }
  list(.jolt_pairs(trimws(x$method), labels, values))
}

# Linear regression in the same layout as Excel's Analysis ToolPak.
.jolt_lm <- function(x, level = 0.95) {
  s <- summary(x)
  res <- residuals(x)
  y <- res + fitted(x)
  n <- length(res)
  intercept <- attr(terms(x), "intercept") == 1
  sse <- sum(res^2)
  sst <- if (intercept) sum((y - mean(y))^2) else sum(y^2)
  df_reg <- x$rank - as.integer(intercept)
  df_res <- x$df.residual
  f <- (sst - sse) / df_reg / (sse / df_res)
  ci <- confint(x, level = level)
  coefs <- s$coefficients
  pct <- paste0(format(100 * level), "%")
  list(
    list(title = "SUMMARY OUTPUT"),
    .jolt_pairs("Regression Statistics",
      c("Multiple R", "R Square", "Adjusted R Square", "Standard Error", "Observations"),
      list(sqrt(s$r.squared), s$r.squared, s$adj.r.squared, s$sigma, n)
    ),
    list(
      title = "ANOVA",
      header = c("", "df", "SS", "MS", "F", "Significance F"),
      cols = list(
        c("Regression", "Residual", "Total"),
        c(df_reg, df_res, df_reg + df_res),
        c(sst - sse, sse, sst),
        c((sst - sse) / df_reg, sse / df_res, NA),
        c(f, NA, NA),
        c(pf(f, df_reg, df_res, lower.tail = FALSE), NA, NA)
      ),
      fmt = rep("", 6)
    ),
    list(
      header = c("", "Coefficients", "Standard Error", "t Stat", "P-value", paste("Lower", pct), paste("Upper", pct)),
      cols = list(rownames(coefs), coefs[, 1], coefs[, 2], coefs[, 3], coefs[, 4], ci[rownames(coefs), 1], ci[rownames(coefs), 2]),
      fmt = rep("", 7)
    )
  )
}

.jolt_summary_lm <- function(s) {
  fs <- s$fstatistic
  stats <- .jolt_pairs("Regression Statistics",
    c("Multiple R", "R Square", "Adjusted R Square", "Standard Error", "Observations", "F", "Significance F"),
    list(
      sqrt(s$r.squared), s$r.squared, s$adj.r.squared, s$sigma, sum(s$df[1:2]),
      if (is.null(fs)) NA else unname(fs[1]),
      if (is.null(fs)) NA else unname(pf(fs[1], fs[2], fs[3], lower.tail = FALSE))
    )
  )
  coefs <- .jolt_table_block(s$coefficients, "Coefficients")[[1]]
  coefs$header[2:5] <- c("Coefficients", "Standard Error", "t Stat", "P-value")
  list(stats, coefs)
}

.jolt_glm <- function(x) {
  s <- summary(x)
  fit <- .jolt_pairs(
    paste0("Generalized linear model (", x$family$family, ", ", x$family$link, " link)"),
    c("Observations", "Null deviance", "Residual deviance", "AIC"),
    list(length(residuals(x)), x$null.deviance, x$deviance, x$aic)
  )
  list(fit, .jolt_table_block(s$coefficients, "Coefficients")[[1]])
}

# ---- JSON (base R has no JSON writer) -----------------------------------------------------------------
# Atomic vectors always become arrays, lists with names become objects, other lists become arrays.
# NA and NaN become null; Inf becomes the text "Inf".

.jolt_json <- function(x) {
  if (is.null(x)) return("null")
  if (is.list(x)) {
    items <- vapply(x, .jolt_json, "", USE.NAMES = FALSE)
    nm <- names(x)
    if (!is.null(nm) && length(x) && all(nzchar(nm))) {
      return(paste0("{", paste0(.jolt_str(nm), ":", items, collapse = ","), "}"))
    }
    return(paste0("[", paste(items, collapse = ","), "]"))
  }
  if (is.factor(x)) x <- as.character(x)
  vals <- if (is.logical(x)) {
    ifelse(is.na(x), "null", ifelse(x, "true", "false"))
  } else if (is.numeric(x)) {
    x <- as.double(x)
    out <- sprintf("%.17g", x)
    out[is.na(x)] <- "null"
    out[is.infinite(x)] <- ifelse(x[is.infinite(x)] > 0, "\"Inf\"", "\"-Inf\"")
    out
  } else {
    x <- as.character(x)
    ifelse(is.na(x), "null", .jolt_str(x))
  }
  paste0("[", paste(vals, collapse = ","), "]")
}

.jolt_str <- function(s) {
  s <- enc2utf8(as.character(s))
  s <- gsub("\\", "\\\\", s, fixed = TRUE)
  s <- gsub("\"", "\\\"", s, fixed = TRUE)
  s <- gsub("\n", "\\n", s, fixed = TRUE)
  s <- gsub("\r", "\\r", s, fixed = TRUE)
  s <- gsub("\t", "\\t", s, fixed = TRUE)
  s <- gsub("[\001-\037]", " ", s)
  paste0("\"", s, "\"")
}
