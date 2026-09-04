

## ===========================================================================
## 0.  Packages
## ===========================================================================
need <- c("readxl", "dplyr", "tidyr", "tibble", "stringr",
          "openxlsx", "ggplot2", "patchwork", "scales",
          "survival", "NADA", "EnvStats")   # survival: KM/Turnbull censored estimation
          # (Section 6C); NADA: optional ROS cross-check; EnvStats: censored mean
new  <- need[!need %in% rownames(installed.packages())]
if (length(new)) install.packages(new, repos = "https://cloud.r-project.org")
invisible(lapply(need, library, character.only = TRUE))

## ---------------------------------------------------------------------------
## CONSOLE POLICY.  The pipeline runs silently. 
## QUIET <- FALSE to restore the full diagnostic trace.
## ---------------------------------------------------------------------------
QUIET <- FALSE     # progress ON, so run_log.txt shows how far the run got
msg   <- function(...) if (!QUIET) base::message(...)
wrote <- function(f) base::message("Written: ", f)

## STAGE() prints regardless of QUIET, marking the start of each major section
## in run_log.txt for progress tracking.
.stage_n <- 0L
STAGE <- function(label) {
  .stage_n <<- .stage_n + 1L
  base::message(sprintf("[STAGE %02d] %s", .stage_n, label))
  utils::flush.console()
}

## Token written to spreadsheets in place of NA (used by every writeData call,
## including the Section 2H harmonisation constructor below, so it must be
## defined here rather than beside write_book far downstream).
NA_TOKEN <- "NA"

## ---------------------------------------------------------------------------
## WARNING POLICY.  Warnings are NOT suppressed.  
## ---------------------------------------------------------------------------
options(warn = 1)



## ===========================================================================
## 1.  Locate SM_2.xlsx and set the output folder
## ===========================================================================
## ---------------------------------------------------------------------------
## INPUT FOLDER.  The analysis folder is searched first; the working directory
## and the usual Desktop / Downloads / Documents locations are retained as
## fallbacks so the script remains portable if the folder is moved or the
## package is re-run by a reader on another machine.
## ---------------------------------------------------------------------------
in_dir <- getwd()   # analysis folder holding SM_2.xlsx (working directory by default)

up <- Sys.getenv("USERPROFILE")
od <- Sys.getenv(c("OneDrive", "OneDriveConsumer", "OneDriveCommercial"))
extra_od <- character(0)
if (nzchar(up) && dir.exists(up)) {
  dd <- list.dirs(up, recursive = FALSE)
  extra_od <- dd[grepl("onedrive", basename(dd), ignore.case = TRUE)]
}
search_dirs <- unique(c(
  in_dir, ".",
  file.path(od[nzchar(od)], "Desktop"),
  file.path(extra_od, "Desktop"),
  if (nzchar(up)) file.path(up, c("OneDrive/Desktop", "Desktop",
                                  "Downloads", "Documents")),
  path.expand("~"), getwd()))
search_dirs <- search_dirs[nzchar(search_dirs) & dir.exists(search_dirs)]

## ---------------------------------------------------------------------------
## ---------------------------------------------------------------------------
OUTPUT_FILES <- c("SM_2_encoding_audit.xlsx", "SM_3_properties.xlsx", "SM_4_fate.xlsx",
                  "Table_SM_2.1_a.xlsx", "Table_SM_2.4.xlsx",  "Table_SM_2.7.xlsx",
                  "Table_SM_2.8.xlsx",    "Table_SM_2.9.xlsx",  "Table_SM_2.9_a.xlsx",
                  "Table_SM_2.10.xlsx",   "Table_SM_2.11.xlsx")
## Table_SM_2.5 and Table_SM_2.6 are NOT listed here: they are DERIVED by
## Section 2H from Table_SM_2.2 and Table_SM_2.3.  Requiring them as inputs
## would reject a correctly prepared workbook.
REQUIRED_SHEETS <- c("Table_SM_2.1", "Table_SM_2.2", "Table_SM_2.3")

locate <- function(label, pattern, must_have = character(0)) {
  for (d in search_dirs) {
    xs  <- list.files(d, pattern = "\\.xlsx$", full.names = TRUE, ignore.case = TRUE)
    xs  <- xs[!grepl("^~\\$", basename(xs))]                        # Excel lock files
    xs  <- xs[!tolower(basename(xs)) %in% tolower(OUTPUT_FILES)]     # our own outputs
    hit <- xs[grepl(pattern, basename(xs), ignore.case = TRUE)]
    exact <- hit[tolower(basename(hit)) == "sm_2.xlsx"]              # exact name wins
    for (h in c(exact, setdiff(hit, exact))) {
      if (!length(must_have)) return(normalizePath(h, winslash = "/"))
      sh <- tryCatch(readxl::excel_sheets(h), error = function(e) character(0))
      ## Sheet names are matched leniently: "Table SM_2.1", "Table_SM_2.1" and
      ## "table-sm-2.1" are treated as the same sheet, so a stray space or dash
      ## in a tab name cannot stop the pipeline finding the workbook.
      key <- function(x) gsub("[^a-z0-9.]", "", tolower(x))
      if (all(key(must_have) %in% key(sh))) return(normalizePath(h, winslash = "/"))
      msg("NOTE: ", basename(h), " matches the name pattern but lacks the required ",
              "sheets (", paste(setdiff(must_have, sh), collapse = ", "), ") - skipped.")
    }
  }
  avail <- vapply(search_dirs, function(d) paste0(
    "   ", d, "  ->  ",
    paste(basename(list.files(d, "\\.xlsx$", ignore.case = TRUE)), collapse = ", ")),
    character(1))
  stop("Could not find the ", label, " file (name matching /", pattern, "/ AND containing ",
       "the sheets ", paste(must_have, collapse = ", "), ").\n",
       paste(avail, collapse = "\n"),
       "\n\nFix: set `in_dir` to the folder that holds SM_2.xlsx.", call. = FALSE)
}

sm2_path <- locate("SM_2", "^sm[ _-]?2", must_have = REQUIRED_SHEETS)   # the input workbook
in_dir   <- dirname(sm2_path)
out_dir  <- in_dir
fig_dir  <- file.path(out_dir, "figures")
dir.create(fig_dir, showWarnings = FALSE, recursive = TRUE)

msg("Input: ", sm2_path)

## ---------------------------------------------------------------------------
## ---------------------------------------------------------------------------
norm_names <- function(nm) gsub("\\s+", " ", trimws(nm))
find_col <- function(nms, pattern, label) {
  hit <- which(grepl(pattern, nms))
  if (length(hit) != 1L)
    stop("Column for '", label, "' (/", pattern, "/) matched ", length(hit),
         " columns:\nHeaders present: ", paste(nms, collapse = " | "), call. = FALSE)
  hit
}

## ===========================================================================
## ---------------------------------------------------------------------------
CHR_SPACE <- "[\u00A0\u2000-\u200B\u202F\u205F\u3000\uFEFF]" # invisible -> DELETED
CHR_DASH  <- "[\u2010\u2011\u2012\u2013\u2014\u2015\u2212\uFE63\uFF0D]" # dashes -> "-"

norm_text <- function(x) {
  x <- as.character(x)
  x <- gsub(CHR_DASH,  "-", x, perl = TRUE)   # U+2212 & friends -> ASCII "-"
  x <- gsub(CHR_SPACE, "",  x, perl = TRUE)   # NBSP & friends   -> deleted
  x
}
norm_header <- function(x) {
  x <- norm_text(x)
  x <- gsub("\u03BC", "\u00B5", x, fixed = TRUE)   # Greek mu -> micro sign
  gsub("\\s+", " ", trimws(x))
}

## reference-key aliases: one spelling of one key differs between the data sheets
## and the (revised) reference table Table_SM_2.1.  Canonicalise to the key used
## in Table_SM_2.1 so the reference join is exact in both directions.
REF_ALIASES <- c("Q. Li et al. (2023)" = "Li Q. et al. (2023)")
canon_ref <- function(x) {
  x <- trimws(as.character(x))
  hit <- match(x, names(REF_ALIASES))
  x[!is.na(hit)] <- unname(REF_ALIASES[hit[!is.na(hit)]])
  x
}

## ---------------------------------------------------------------------------
## REFERENCE-BLOCK TRIMMING  
## ---------------------------------------------------------------------------
## 
## ---------------------------------------------------------------------------
REF_KEY_PAT <- "\\((?:1[89]|20)[0-9]{2}[a-z]?\\)"      # "... (2018)" / "... (2018a)"

trim_ref_block <- function(df, key_col = 1L, quiet = FALSE) {
  if (is.null(df) || !nrow(df)) return(df)
  k  <- trimws(as.character(df[[key_col]]))
  ok <- !is.na(k) & nzchar(k)
  ## contiguous block from the top: stop at the first blank key
  last <- if (!ok[1]) 0L else {
    br <- which(!ok)
    if (length(br)) br[1] - 1L else length(ok)
  }
  keep <- rep(FALSE, nrow(df))
  if (last > 0L) keep[seq_len(last)] <- TRUE
  ## and it must actually look like a citation key
  keep <- keep & grepl(REF_KEY_PAT, k, perl = TRUE)
  n_drop <- nrow(df) - sum(keep)
  if (n_drop > 0L && !quiet)
    msg("  Table_SM_2.1: ", sum(keep), " reference key(s) retained; ", n_drop,
            " non-reference row(s) below the list ignored (appendix block pasted ",
            "into the sheet - see Section 10).")
  df[keep, , drop = FALSE]
}

## ---- substitution log ------------------------------------------------------
ENC <- new.env(parent = emptyenv())
ENC$cells <- list(); ENC$seen <- character(0); ENC$refs <- list()

xl_col <- function(i) {                       # 1 -> "A", 27 -> "AA"
  s <- ""
  while (i > 0L) { r <- (i - 1L) %% 26L; s <- paste0(LETTERS[r + 1L], s); i <- (i - 1L) %/% 26L }
  s
}
chr_class <- function(s) {
  h <- c(if (grepl("\u00A0", s, fixed = TRUE)) "U+00A0 NBSP -> deleted",
         if (grepl("\u2212", s, fixed = TRUE)) "U+2212 MINUS -> '-'",
         if (grepl("\u2013", s, fixed = TRUE)) "U+2013 EN DASH -> '-'",
         if (grepl("\u2014", s, fixed = TRUE)) "U+2014 EM DASH -> '-'",
         if (grepl("[\u2010\u2011\u2012\u2015\uFE63\uFF0D]", s)) "other Unicode hyphen -> '-'",
         if (grepl("[\u2000-\u200B\u202F\u205F\u3000\uFEFF]", s)) "other invisible space -> deleted")
  if (length(h)) paste(h, collapse = "; ") else ""
}

## sheets that FEED the analysis (the rest of SM_2 holds pipeline outputs)
SM2_INPUT_SHEETS <- c("Table_SM_2.1", "Table_SM_2.2", "Table_SM_2.3",
                      "Table_SM_2.5", "Table_SM_2.6")

## ---------------------------------------------------------------------------
## ---------------------------------------------------------------------------
MIN_HDR_CELLS <- 3L        # a header has >= 3 populated cells; a caption has 1
HDR_PROBE     <- 40L       # rows scanned when looking for the header

HDR <- new.env(parent = emptyenv())
HDR$skip <- list()

## Resolve a canonical sheet name (e.g. "Table_SM_2.1") to the actual tab name
## in the workbook, ignoring spaces, underscores, dashes and case.  This makes
## every read tolerant of cosmetic tab renaming in Excel.
SHEET_KEY <- function(x) gsub("[^a-z0-9.]", "", tolower(x))
resolve_sheet <- function(sheet) {
  actual <- tryCatch(readxl::excel_sheets(sm2_path), error = function(e) character(0))
  if (sheet %in% actual) return(sheet)
  hit <- actual[SHEET_KEY(actual) == SHEET_KEY(sheet)]
  if (length(hit)) return(hit[1])
  sheet
}

header_skip <- function(sheet) {
  sheet  <- resolve_sheet(sheet)
  cached <- HDR$skip[[sheet]]
  if (!is.null(cached)) return(cached)
  pk <- tryCatch(
    readxl::read_excel(sm2_path, sheet = sheet, col_types = "text",
                       col_names = FALSE, n_max = HDR_PROBE,
                       .name_repair = "minimal"),
    error = function(e) NULL)
  s <- 0L
  if (!is.null(pk) && nrow(pk)) {
    filled <- vapply(seq_len(nrow(pk)), function(i) {
      r <- as.character(unlist(pk[i, ], use.names = FALSE))
      sum(!is.na(r) & nzchar(trimws(r)))
    }, integer(1))
    hit <- which(filled >= MIN_HDR_CELLS)
    if (length(hit)) s <- as.integer(hit[1] - 1L)
  }
  HDR$skip[[sheet]] <- s
  s
}

## ---- THE reader: every read of SM_2.xlsx goes through this ------------------
read_sm2 <- function(sheet) {
  sheet <- resolve_sheet(sheet)                  # tolerate "Table SM_2.1" etc.
  sk  <- header_skip(sheet)                      # caption rows above the header
  raw <- readxl::read_excel(sm2_path, sheet = sheet, col_types = "text",
                            .name_repair = "minimal", skip = sk)
  hdr <- norm_header(names(raw))
  names(raw) <- hdr
  first_time <- !(sheet %in% ENC$seen)
  if (first_time && sk > 0L)
    msg("  header of sheet ", sheet, " found on Excel row ", sk + 1L,
            " (", sk, " caption/note row(s) skipped).")
  for (j in seq_along(raw)) {
    v0 <- as.character(raw[[j]])
    v1 <- norm_text(v0)
    if (first_time) {
      ch <- which(!is.na(v0) & v0 != v1)
      if (length(ch))
        ENC$cells[[length(ENC$cells) + 1L]] <- data.frame(
          sheet = sheet,
          sheet_role = if (sheet %in% SM2_INPUT_SHEETS) "input (feeds analysis)"
                       else "derived (pipeline output)",
          cell       = paste0(xl_col(j), ch + sk + 1L),   # TRUE Excel row
          column     = hdr[j],
          original   = v0[ch],
          normalised = v1[ch],
          substitution = vapply(v0[ch], chr_class, character(1), USE.NAMES = FALSE),
          stringsAsFactors = FALSE)
    }
    ## reference columns: canonicalise the key spelling
    if (grepl("^References", hdr[j])) {
      v2 <- canon_ref(v1)
      if (first_time) {
        rc <- which(!is.na(v1) & v1 != v2)
        if (length(rc))
          ENC$refs[[length(ENC$refs) + 1L]] <- data.frame(
            sheet = sheet, cell = paste0(xl_col(j), rc + sk + 1L),
            original = v1[rc], canonical = v2[rc], stringsAsFactors = FALSE)
      }
      v1 <- v2
    }
    raw[[j]] <- v1
  }
  if (first_time) ENC$seen <- c(ENC$seen, sheet)
  raw
}

## ---- pre-flight scan of the WHOLE workbook ---------------------------------
SM2_SHEETS <- readxl::excel_sheets(sm2_path)
invisible(lapply(SM2_SHEETS, function(s) try(read_sm2(s), silent = TRUE)))
enc_cells <- if (length(ENC$cells)) do.call(rbind, ENC$cells) else
  data.frame(sheet = character(0), sheet_role = character(0), cell = character(0),
             column = character(0), original = character(0), normalised = character(0),
             substitution = character(0), stringsAsFactors = FALSE)
enc_refs  <- if (length(ENC$refs)) do.call(rbind, ENC$refs) else
  data.frame(sheet = character(0), cell = character(0),
             original = character(0), canonical = character(0), stringsAsFactors = FALSE)

enc_by_col <- if (nrow(enc_cells))
  enc_cells %>% dplyr::count(sheet, sheet_role, column, substitution, name = "n_cells") %>%
    dplyr::arrange(sheet, column) else
  data.frame(sheet = character(0), sheet_role = character(0), column = character(0),
             substitution = character(0), n_cells = integer(0))

n_in <- sum(enc_cells$sheet %in% SM2_INPUT_SHEETS)
msg("Encoding pre-processing: ", nrow(enc_cells), " cell(s) normalised in SM_2.xlsx (",
        n_in, " in the five INPUT sheets; ", nrow(enc_cells) - n_in,
        " in derived sheets, which the pipeline regenerates).")
if (nrow(enc_refs))
  msg("Reference-key aliases applied: ", nrow(enc_refs), " cell(s)  [",
          paste(unique(paste0("'", enc_refs$original, "' -> '", enc_refs$canonical, "'")),
                collapse = "; "), "]")
if (nrow(enc_by_col))
  if (!QUIET) print(enc_by_col[enc_by_col$sheet %in% SM2_INPUT_SHEETS, ], row.names = FALSE)


## ===========================================================================
## 1c.  VALUE-LEVEL CORRECTIONS  --  NONE ARE APPLIED
## ---------------------------------------------------------------------------
## ---------------------------------------------------------------------------
VALUE_CORRECTIONS <- data.frame(
  sheet     = character(0),   # e.g. "Table_SM_2.6"
  compound  = character(0),   # e.g. "Tylosin"
  column    = character(0),   # start-anchored regex, e.g. "^t1/2 SED"
  old_value = character(0),   # exact text to match, e.g. "45145"
  new_value = character(0),   # replacement text
  reason    = character(0),
  stringsAsFactors = FALSE)

apply_corrections <- function(df, sheet) {
  if (!nrow(VALUE_CORRECTIONS)) return(df)          # <- current state: no-op
  nms <- norm_header(names(df))
  ci_cmp <- which(grepl("^Compound", nms, ignore.case = TRUE))[1]
  if (is.na(ci_cmp)) return(df)
  cmp <- as.character(df[[ci_cmp]])
  for (i in seq_along(cmp))
    if (is.na(cmp[i]) || cmp[i] == "") cmp[i] <- if (i > 1L) cmp[i - 1L] else NA_character_
  rules <- VALUE_CORRECTIONS[VALUE_CORRECTIONS$sheet == sheet, , drop = FALSE]
  for (k in seq_len(nrow(rules))) {
    ci <- which(grepl(rules$column[k], nms))[1]
    if (is.na(ci)) next
    hit <- which(trimws(tolower(cmp)) == tolower(rules$compound[k]) &
                 !is.na(df[[ci]]) & trimws(as.character(df[[ci]])) == rules$old_value[k])
    if (length(hit)) {
      df[[ci]][hit] <- rules$new_value[k]
      msg("  apply_corrections [", sheet, "]: ", length(hit), " cell(s) patched (",
              names(df)[ci], "): ", rules$old_value[k], " -> ", rules$new_value[k],
              "  [", rules$reason[k], "]")
    }
  }
  df
}
msg("Value-level corrections registered: ", nrow(VALUE_CORRECTIONS),
        "  (none - the script alters no compiled value; data corrections are made ",
        "at source in SM_2.xlsx and documented there).")



## present beside SM_2.xlsx, its corrected values override matching entries.
locate_cls <- function() {
  ## Locate Table_SM_2.1_corrected.xlsx (the Stage-1 abstract-audit classification).
  ## Returns the path, or NULL (never stops the session).
  for (d in search_dirs) {
    xs  <- list.files(d, pattern = "\\.xlsx$", full.names = TRUE, ignore.case = TRUE)
    hit <- xs[grepl("table_sm_2\\.?1_corrected", basename(xs), ignore.case = TRUE)]
    if (length(hit)) return(normalizePath(hit[1], winslash = "/"))
  }
  NULL
}
cls_path <- locate_cls()          # may be NULL; Section 11 checks and reports
if (!is.null(cls_path)) {
  msg("Study classification file found: ", cls_path)
} else {
  msg("NOTE: Table_SM_2.1_corrected.xlsx not found in the search folders. ",
          "Section 11 will look for a classification column inside SM_2.xlsx instead.")
}


## sheet names (edit here if your tab names differ) --------------------------
SH_RAW_PROP  <- "Table_SM_2.2"
SH_RAW_FATE  <- "Table_SM_2.3"
SH_HARM_PROP <- "Table_SM_2.5"
SH_HARM_FATE <- "Table_SM_2.6"


## ===========================================================================
## 2.  Helpers: numeric cleaner + strict recast + column finder
## ===========================================================================
## clean_num(): defensive.  Input reaching it has already been normalised by
## read_sm2() (Section 1b); the substitutions are repeated here so the helper is
## still correct if it is ever called on un-normalised text.
clean_num <- function(x) {
  x <- as.character(x)
  x <- gsub(CHR_DASH,  "-", x, perl = TRUE)   # U+2212 & Unicode hyphens -> "-"
  x <- gsub(CHR_SPACE, "",  x, perl = TRUE)   # NBSP & invisible spaces  -> deleted
  x <- gsub("[, ]",    "",  x)                # thousands separator / stray spaces
  suppressWarnings(as.numeric(x))
}

## ---------------------------------------------------------------------------
## ---------------------------------------------------------------------------
NUM <- new.env(parent = emptyenv())
NUM$report <- list(); NUM$bad <- list()

to_numeric_strict <- function(x, sheet, column) {
  s <- trimws(as.character(x))
  s[s == ""] <- NA_character_
  v <- clean_num(s)
  key <- paste(sheet, column, sep = "|")
  bad <- which(!is.na(s) & is.na(v))
  if (!key %in% names(NUM$report)) {
    NUM$report[[key]] <- data.frame(
      sheet = sheet, column = column,
      n_nonempty   = sum(!is.na(s)),
      n_numeric    = sum(!is.na(v)),
      n_unparsable = length(bad),
      status = if (length(bad)) "FAIL - see 'unparsable_cells'" else "OK - fully numeric",
      stringsAsFactors = FALSE)
    if (length(bad))
      NUM$bad[[key]] <- data.frame(sheet = sheet, column = column,
                                   row_in_sheet = bad + 1L, value = s[bad],
                                   stringsAsFactors = FALSE)
  }
  v
}

## fold accents to ASCII so reference keys join across platforms/encodings
fold_ascii <- function(x) {
  y <- iconv(x, from = "UTF-8", to = "ASCII//TRANSLIT")
  y[is.na(y)] <- x[is.na(y)]
  gsub("[^A-Za-z0-9 ().&,'/-]", "", y)
}

## normalise a header (collapse internal whitespace, trim)
norm_names <- function(nm) gsub("\\s+", " ", trimws(nm))

## return the single column INDEX whose normalised name matches `pattern`
find_col <- function(nms, pattern, label) {
  hit <- which(grepl(pattern, nms))
  if (length(hit) != 1L)
    stop("Column for '", label, "' (/", pattern, "/) matched ", length(hit),
         " columns: ", paste(nms[hit], collapse = " | "),
         "\nHeaders present: ", paste(nms, collapse = " | "), call. = FALSE)
  hit
}

## start-anchored patterns -> canonical analysis names (shared by A and B)
## Column-matching patterns are written to match BOTH the current uniform headers
## (MW ..., log KOW/KOC/DOW, C_SW/SED/HZ/GW, t½ ...) AND the earlier verbose
## headers (Molecular weight ..., log Kow/Koc/Dow, Concentration ... (µg/L),
## t1/2 ... (days)), so the pipeline is robust to either naming convention.
## Start-anchoring keeps the pH-of-measurement columns (pH log KOC, pH log DOW)
## out of the descriptor matches, and the half-life patterns "^t.* <compartment>"
## match the ½ glyph without embedding it in the source.
prop_patterns <- c(
  Molecular_weight = "^(MW|Molecular weight)", pKa = "^pKa",
  log_Kow = "^log +[Kk][Oo][Ww]", `log _Koc` = "^log +[Kk][Oo][Cc]",
  log_Dow = "^log +[Dd][Oo][Ww]", log_S = "^log +S")
conc_patterns <- c(
  Concentration_SW = "^(C_SW|Concent.* SW)", Concentration_SED = "^(C_SED|Concent.* SED)",
  Concentration_HZ = "^(C_HZ|Concent.* HZ)", Concentration_GW = "^(C_GW|Concent.* GW)")
hl_patterns <- c(
  half_life_SW = "^t.* SW", half_life_SED = "^t.* SED", half_life_HZ = "^t.* HZ",
  half_life_aerobic = "^t.* aerobic", half_life_anoxic = "^t.* anoxic",
  half_life_GW = "^t.* GW")
fate_patterns <- c(conc_patterns, hl_patterns)


## ===========================================================================
## SECTION 2H.  HARMONISATION CONSTRUCTOR
STAGE("2H  build harmonised SM_2.5 / SM_2.6 from raw tables")

##              Table_SM_2.2 -> Table_SM_2.5   and   Table_SM_2.3 -> Table_SM_2.6
## ---------------------------------------------------------------------------
##
##
## WHAT IT WRITES
##   Table_SM_2.5.xlsx      derived harmonised physicochemical properties, with
##                          its own README, cell-level rule audit and parameters
##   Table_SM_2.6.xlsx      derived harmonised environmental fate, likewise
##   SM_2_generated.xlsx    a working copy of the input workbook with the two
##                          derived tables inserted; this is the file every later
##                          section reads
## ===========================================================================

msg("\n== SECTION 2H: constructing Table_SM_2.5 and Table_SM_2.6 from the raw tables ==")

## ---------------------------------------------------------------------------
## 2H.1  HARMONISATION PARAMETERS.  Every choice that is not fully determined by
##       rules (i)-(vii) is exposed here, named, and reported in the audit.
## ---------------------------------------------------------------------------

## Rule (i).  A left-censored entry is substituted with HALF its limit.  Where
## the study states a limit x the substituted value is x/2.  Where no limit is
## stated (bare n.d., <LOQ, <LOD, <DL) an assumed limit is used in its place and
## halved in exactly the same way, so rule (i) is one rule with one arithmetic:
## the substituted value is always limit/2.
##   RETIRED: no fixed assumed limit is applied to concentrations. A no-limit
##   non-detect is excluded from the point table and enters the censored estimator
##   (Section 6C) at the lowest detected value of its own variable.
HARM_ND_LIMIT_WATER <- NA_real_  # retired (was 1e-4)
HARM_ND_LIMIT_SED   <- NA_real_  # retired (was 1e-4)

## Rule (i), half-lives.  No assumed limit is defined for a half-life: "not
## detected" is a statement about the analyte, not about its persistence, and
## substituting 5e-5 d would assert a half-life of four seconds.  A bare
## non-detect is therefore excluded from the half-life variables.  A STATED
## limit is still halved (<x -> x/2), because that is a genuine upper bound.
##   "exclude"  : drop the entry                                   [default]
##   "constant" : substitute HARM_ND_LIMIT_HL / 2
HARM_ND_HL_POLICY <- "exclude"
HARM_ND_LIMIT_HL  <- 1e-4        # only used when the policy is "constant"

## Rule (i), physicochemical properties.  Same reasoning: <x -> x/2 is applied,
## a bare non-detect is excluded.
HARM_ND_PROP_POLICY <- "exclude"

## Rule (vi).  A concentration reported as 0 carries the same information as a
## non-detect - the analyte was not measurably present - and is resolved under
## rule (i), taking the assumed limit and halving it.  This applies to a
## standalone 0 and to a 0 appearing as the lower bound of a range, so that the
## same entry is treated identically wherever it occurs.  Physicochemical
## properties are never affected: a logarithm of 0 is a ratio of 1 and is a
## genuine value.
##   "as_nondetect" : resolve under rule (i)                        [default]
##   "retain"       : keep 0 as a measured value
##   "exclude"      : drop it
HARM_ZERO_CONC_POLICY <- "as_nondetect"

## Rule (vii).  t1/2 = 0 d is physically implausible and is always excluded.
## (No switch: this is not a judgement call.)

## ---------------------------------------------------------------------------
## 2H.2  CELL PARSER.  Returns the rule applied and the value produced, or NA
##       with the reason for exclusion.  One function, used for every variable;
##       the variable KIND supplies the unit-dependent constant.
##       kind: "conc_water" | "conc_sed" | "half_life" | "property"
## ---------------------------------------------------------------------------

## number token, tolerant of E-notation and a leading sign
HARM_NUM <- "[-+]?[0-9]*\\.?[0-9]+(?:[eE][-+]?[0-9]+)?"

harm_prep <- function(s) {
  x <- norm_text(s)                       # NBSP stripped, dashes -> "-"
  x <- gsub("\u00A0", "", x, fixed = TRUE)
  ## a x 10^-n  /  a x 10-n  /  a X10-n   ->  aE-n   (removes the hyphen that
  ## would otherwise be read as a range separator)
  x <- gsub(paste0("(", HARM_NUM, ")\\s*[\u00D7xX*]\\s*10\\s*\\^?\\s*([-+]?[0-9]+)"),
            "\\1E\\2", x, perl = TRUE)
  ## parenthesised negatives:  (-1.97)  ->  -1.97
  x <- gsub("\\(\\s*(-\\s*[0-9.]+)\\s*\\)", "\\1", x, perl = TRUE)
  x <- gsub("\\s+", " ", trimws(x))
  x
}

harm_first_num <- function(x) {
  m <- regmatches(x, regexpr(HARM_NUM, x, perl = TRUE))
  if (!length(m)) return(NA_real_)
  suppressWarnings(as.numeric(m))
}

## the value substituted for a censored entry that states no limit
harm_nd_value <- function(kind) {
  ## Concentrations: no fabricated point value. A bare non-detect is excluded from
  ## the point table and handled by the censored estimator (Section 6C).
  switch(kind,
         conc_water = NA_real_,
         conc_sed   = NA_real_,
         half_life  = if (identical(HARM_ND_HL_POLICY, "constant")) HARM_ND_LIMIT_HL / 2 else NA_real_,
         property   = NA_real_)
}

## resolve ONE side of an entry (or a whole single-valued entry).
## Rule (vi) is applied here rather than downstream, so that a zero is treated
## identically whether it stands alone or forms the lower bound of a range.
harm_side <- function(x, kind) {
  x <- trimws(x)
  low <- tolower(x)
  nd_val <- harm_nd_value(kind)
  ## bare non-detect, no stated limit
  if (grepl("^(n\\.?\\s*d\\.?|nd|not\\s*detect(?:ed)?|<\\s*lo[dq]\\.?|<\\s*dl\\.?)$",
            low, perl = TRUE, ignore.case = TRUE))
    return(list(v = nd_val,
                rule = if (is.na(nd_val)) "rule (i) non-detect, no assumed limit defined -> excluded"
                       else "rule (i) non-detect -> assumed limit / 2"))
  ## < x        -> x/2
  if (grepl(paste0("^<\\s*=?\\s*", HARM_NUM, "$"), x, perl = TRUE))
    return(list(v = harm_first_num(x) / 2, rule = "rule (i) left-censored -> x/2"))
  ## > x , >= x -> x
  if (grepl(paste0("^[>\u2265]\\s*=?\\s*", HARM_NUM, "$"), x, perl = TRUE))
    return(list(v = harm_first_num(x), rule = "rule (iii) right-censored -> x"))
  ## ~ x , = x
  if (grepl(paste0("^[\u2248\u2245~\u224B]\\s*", HARM_NUM, "$"), x, perl = TRUE))
    return(list(v = harm_first_num(x), rule = "rule (iv) approximate -> x"))
  ## x +- s
  if (grepl(paste0("^", HARM_NUM, "\\s*(\u00B1|\\+/-|\\+-)\\s*", HARM_NUM, "$"), x, perl = TRUE))
    return(list(v = harm_first_num(x), rule = "rule (iv) mean +- SD/SE -> mean"))
  ## plain number
  if (grepl(paste0("^", HARM_NUM, "$"), x, perl = TRUE)) {
    v <- harm_first_num(x)
    ## ---- rule (vi): a zero concentration is a non-detect --------------------
    if (v == 0 && kind %in% c("conc_water", "conc_sed")) {
      if (identical(HARM_ZERO_CONC_POLICY, "as_nondetect"))
        return(list(v = nd_val, rule = "rule (vi) zero concentration -> rule (i), assumed limit / 2"))
      if (identical(HARM_ZERO_CONC_POLICY, "exclude"))
        return(list(v = NA_real_, rule = "rule (vi) zero concentration -> excluded"))
      return(list(v = 0, rule = "rule (vi) zero retained as analytically confirmed"))
    }
    return(list(v = v, rule = "standard value"))
  }
  list(v = NA_real_, rule = paste0("UNPARSED: ", x))
}

harmonise_cell <- function(s, kind) {
  if (is.na(s)) return(list(v = NA_real_, rule = "empty", excl = TRUE))
  x <- harm_prep(s)
  if (!nzchar(x) || x == "-")
    return(list(v = NA_real_, rule = "empty", excl = TRUE))
  low <- tolower(x)

  ## ---- rule (v): qualitative extremes and non-quantitative entries ---------
  if (grepl(">>|<<", x, fixed = FALSE))
    return(list(v = NA_real_, rule = "rule (v) qualitative extreme -> excluded", excl = TRUE))
  if (gsub("[ .]", "", low) %in% c("na", "n/a", "nd?", "notavailable", "notreported", "nr"))
    return(list(v = NA_real_, rule = "rule (v) not available -> excluded", excl = TRUE))

  ## ---- rule (ii): ranges ---------------------------------------------------
  ## split on a "-" that follows a digit or ")" and precedes a number, "<" or
  ## a sign; E-notation exponents are already protected by harm_prep().
  sep <- "(?<=[0-9)])\\s*-\\s*(?=[<\u2264]?\\s*[-+(]?[0-9.])"
  ## a censored lower bound (n.d.- , <LOQ- , <x-) does not match the lookbehind,
  ## so handle it explicitly first
  ## Censored-floor range: a lower bound that is a non-detect token (n.d., <LOQ,
  ## <LOD, <DL) or a censored numeric (<x), followed by "-" and a detected upper
  ## bound.  Matched case-insensitively and tolerant of internal spaces
  ## ("< LOQ-0.344"), so no censored-floor form is dropped.
  floor_lb <- "n\\.?\\s*d\\.?|nd|not\\s*detect(?:ed)?|<\\s*lo[dq]\\.?|<\\s*dl\\.?|<\\s*=?\\s*"
  m_nd <- regmatches(x, regexec(paste0("^(", floor_lb, HARM_NUM, ")\\s*-\\s*(.+)$"),
                                x, perl = TRUE, ignore.case = TRUE))[[1]]
  parts <- NULL
  if (length(m_nd) == 3L) {
    parts <- c(m_nd[2], m_nd[3])
  } else if (grepl(sep, x, perl = TRUE)) {
    p <- strsplit(x, sep, perl = TRUE)[[1]]
    if (length(p) == 2L) parts <- p
  }
  if (!is.null(parts)) {
    lo <- harm_side(trimws(parts[1]), kind)
    hi <- harm_side(trimws(parts[2]), kind)
    if (grepl("^UNPARSED", lo$rule) || grepl("^UNPARSED", hi$rule))
      return(list(v = NA_real_, rule = paste0("UNPARSED range: ", x), excl = TRUE))
    if (is.na(lo$v) || is.na(hi$v)) {
      ## A censored-floor range (n.d.-y, 0-y, <LOQ-y) resolves to NA because no-limit
      ## non-detects and zeros are no longer substituted. It is interval-censored data:
      ## excluded from the single-value point table and handled by the censored estimator
      ## (Section 6C), not a parser failure. Any other NA bound still trips the gate.
      if (grepl("non-detect|rule \\(vi\\)|no assumed limit", paste(lo$rule, hi$rule)))
        return(list(v = NA_real_,
                    rule = "rule (ii) censored-floor range -> excluded from point table; handled by Section 6C",
                    excl = TRUE))
      return(list(v = NA_real_,
                  rule = paste0("rule (ii) range with an unresolvable bound -> excluded (", x, ")"),
                  excl = TRUE))
    }
    a <- min(lo$v, hi$v); b <- max(lo$v, hi$v)   # tolerate descending ranges
    if (kind == "half_life" && (a + b) / 2 == 0)
      return(list(v = NA_real_, rule = "rule (vii) t1/2 = 0 d -> excluded", excl = TRUE))
    return(list(v = (a + b) / 2,
                rule = paste0("rule (ii) range midpoint [lower: ", lo$rule, "]"),
                excl = FALSE))
  }

  ## ---- single value --------------------------------------------------------
  r <- harm_side(x, kind)
  if (grepl("^UNPARSED", r$rule)) return(list(v = NA_real_, rule = r$rule, excl = TRUE))
  if (is.na(r$v))                 return(list(v = NA_real_, rule = r$rule, excl = TRUE))

  ## ---- rule (vii): a half-life of zero -----------------------------------
  if (kind == "half_life" && r$v == 0)
    return(list(v = NA_real_, rule = "rule (vii) t1/2 = 0 d -> excluded", excl = TRUE))

  list(v = r$v, rule = r$rule, excl = FALSE)
}

## ---------------------------------------------------------------------------
## 2H.3  BUILD ONE HARMONISED TABLE
## ---------------------------------------------------------------------------

HARM_KIND <- c(Concentration_SW = "conc_water", Concentration_SED = "conc_sed",
               Concentration_HZ = "conc_water", Concentration_GW = "conc_water",
               half_life_SW = "half_life", half_life_SED = "half_life",
               half_life_HZ = "half_life", half_life_aerobic = "half_life",
               half_life_anoxic = "half_life", half_life_GW = "half_life",
               Molecular_weight = "property", pKa = "property",
               log_Kow = "property", `log _Koc` = "property",
               log_Dow = "property", log_S = "property")

## header text written to the generated sheets (must match the patterns above)
HARM_OUT_HEADER <- c(Molecular_weight = "Molecular weight (g/mol)", pKa = "pKa",
                     log_Kow = "log Kow", `log _Koc` = "log Koc",
                     log_Dow = "log Dow", log_S = "log S",
                     Concentration_SW = "C_SW", Concentration_SED = "C_SED",
                     Concentration_HZ = "C_HZ", Concentration_GW = "C_GW",
                     half_life_SW = "t\u00BD SW", half_life_SED = "t\u00BD SED",
                     half_life_HZ = "t\u00BD HZ", half_life_aerobic = "t\u00BD aerobic",
                     half_life_anoxic = "t\u00BD anoxic", half_life_GW = "t\u00BD GW")

harm_ffill <- function(v) {
  v <- as.character(v)
  for (i in seq_along(v))
    if (i > 1L && (is.na(v[i]) || !nzchar(trimws(v[i])))) v[i] <- v[i - 1L]
  v
}

build_harmonised <- function(src_sheet, patterns, out_label) {
  raw <- read_sm2(src_sheet)
  nms <- norm_names(names(raw))

  ci_cls <- which(grepl("^(PCs|Clinical Application|Therapeutic|Class)", nms, ignore.case = TRUE))[1]
  ci_cmp <- find_col(nms, "^Compound", "Compound")
  ci_ref <- find_col(nms, "[Rr]eference", "References")

  cls <- if (is.na(ci_cls)) rep(NA_character_, nrow(raw)) else harm_ffill(raw[[ci_cls]])
  cmp <- harm_ffill(raw[[ci_cmp]])
  ref <- as.character(raw[[ci_ref]])

  out <- data.frame(`PCs class` = trimws(cls), Compound = trimws(cmp),
                    check.names = FALSE, stringsAsFactors = FALSE)
  audit <- list()

  for (canon in names(patterns)) {
    ci <- find_col(nms, patterns[[canon]], canon)
    src <- as.character(raw[[ci]])
    kind <- unname(HARM_KIND[canon])
    vals <- rep(NA_real_, length(src))
    for (i in seq_along(src)) {
      s <- src[i]
      if (is.na(s) || !nzchar(trimws(s))) next
      r <- harmonise_cell(s, kind)
      vals[i] <- r$v
      audit[[length(audit) + 1L]] <- data.frame(
        source_sheet = src_sheet, source_row = i + header_skip(src_sheet) + 1L,
        Compound = trimws(cmp[i]), Reference = trimws(ref[i]),
        Variable = unname(HARM_OUT_HEADER[canon]),
        raw_entry = trimws(as.character(s)),
        rule_applied = r$rule,
        harmonised_value = r$v,
        status = if (is.na(r$v)) "excluded" else "retained",
        stringsAsFactors = FALSE)
    }
    out[[unname(HARM_OUT_HEADER[canon])]] <- vals
  }

  ## Molecular weight is a compound-intrinsic constant reported once per
  ## compound in Table_SM_2.2; broadcast it to every row of that compound so the
  ## record-weighted (HAR) summaries treat it like the other descriptors.
  if ("Molecular weight (g/mol)" %in% names(out)) {
    mwc <- tapply(out[["Molecular weight (g/mol)"]], out$Compound,
                  function(z) { z <- z[is.finite(z)]; if (length(z)) z[1] else NA_real_ })
    out[["Molecular weight (g/mol)"]] <- unname(mwc[out$Compound])
  }

  out$References <- trimws(ref)
  attr(out, "audit") <- if (length(audit)) do.call(rbind, audit) else NULL
  attr(out, "label") <- out_label
  out
}

harm_prop_new <- build_harmonised(SH_RAW_PROP, prop_patterns, SH_HARM_PROP)
harm_fate_new <- build_harmonised(SH_RAW_FATE, fate_patterns, SH_HARM_FATE)

harm_audit <- rbind(attr(harm_prop_new, "audit"), attr(harm_fate_new, "audit"))

## ---------------------------------------------------------------------------
## 2H.4  AUDIT SUMMARY AND HARD STOPS
## ---------------------------------------------------------------------------
## A cell is a silent loss if the parser could not resolve it (rule begins
## "UNPARSED") OR if it yielded NA without an explicit rule-(v)/(vi)/(vii)
## exclusion.  Testing the OUTCOME as well as the label means a future parser
## gap cannot pass through even if it is not labelled "UNPARSED".
.expl_excl <- "rule \\(v\\)|rule \\(vi\\)|rule \\(vii\\)|empty|no assumed limit|censored-floor range"
harm_unparsed <- harm_audit[
  grepl("^UNPARSED", harm_audit$rule_applied) |
    (is.na(harm_audit$harmonised_value) &
       !grepl(.expl_excl, harm_audit$rule_applied)), , drop = FALSE]
if (nrow(harm_unparsed)) {
  base::message("SECTION 2H: ", nrow(harm_unparsed),
                " raw cell(s) could not be parsed by rules (i)-(vii):")
  print(utils::head(harm_unparsed[, c("source_sheet", "source_row", "Compound",
                                      "Variable", "raw_entry")], 20))
  stop("Section 2H: unparsable raw entries. Correct them in the raw tables, ",
       "or extend the parser, then re-run. The pipeline will not proceed on ",
       "silently dropped data.", call. = FALSE)
}

harm_summary <- harm_audit %>%
  dplyr::mutate(rule_group = sub(" \\[.*$", "", rule_applied)) %>%
  dplyr::group_by(Variable, rule_group) %>%
  dplyr::summarise(n = dplyr::n(), .groups = "drop") %>%
  tidyr::pivot_wider(names_from = rule_group, values_from = n, values_fill = 0) %>%
  as.data.frame()

harm_counts <- data.frame(
  Table = c(SH_HARM_PROP, SH_HARM_FATE),
  source = c(SH_RAW_PROP, SH_RAW_FATE),
  rows = c(nrow(harm_prop_new), nrow(harm_fate_new)),
  raw_cells_seen = c(sum(harm_audit$source_sheet == SH_RAW_PROP),
                     sum(harm_audit$source_sheet == SH_RAW_FATE)),
  values_retained = c(sum(harm_audit$source_sheet == SH_RAW_PROP & harm_audit$status == "retained"),
                      sum(harm_audit$source_sheet == SH_RAW_FATE & harm_audit$status == "retained")),
  values_excluded = c(sum(harm_audit$source_sheet == SH_RAW_PROP & harm_audit$status == "excluded"),
                      sum(harm_audit$source_sheet == SH_RAW_FATE & harm_audit$status == "excluded")),
  stringsAsFactors = FALSE)

harm_params <- data.frame(
  parameter = c("HARM_ND_LIMIT_WATER", "HARM_ND_LIMIT_SED", "substituted value",
                "HARM_ND_HL_POLICY", "HARM_ND_PROP_POLICY", "HARM_ZERO_CONC_POLICY"),
  value = c(format(HARM_ND_LIMIT_WATER, scientific = TRUE),
            format(HARM_ND_LIMIT_SED, scientific = TRUE),
            format(HARM_ND_LIMIT_WATER / 2, scientific = TRUE),
            HARM_ND_HL_POLICY, HARM_ND_PROP_POLICY, HARM_ZERO_CONC_POLICY),
  meaning = c("rule (i) assumed limit for a non-detect with no stated limit, ug L-1 (C_SW, C_HZ, C_GW)",
              "rule (i) assumed limit for a non-detect with no stated limit, ug kg-1 (C_SED)",
              "the value actually substituted: the assumed limit halved, as rule (i) halves a stated limit",
              "rule (i) treatment of a bare non-detect in a half-life column",
              "rule (i) treatment of a bare non-detect in a property column",
              "rule (vi) treatment of a concentration reported as zero"),
  stringsAsFactors = FALSE)

msg("  ", SH_HARM_PROP, ": ", nrow(harm_prop_new), " rows built from ", SH_RAW_PROP)
msg("  ", SH_HARM_FATE, ": ", nrow(harm_fate_new), " rows built from ", SH_RAW_FATE)
msg("  raw cells classified: ", nrow(harm_audit),
    "  retained: ", sum(harm_audit$status == "retained"),
    "  excluded: ", sum(harm_audit$status == "excluded"))

## ---------------------------------------------------------------------------
## 2H.5  WRITE THE GENERATED WORKBOOK AND RE-POINT sm2_path
##       Every other sheet of the input workbook is copied through unchanged, so
##       the generated file is a drop-in replacement for SM_2.xlsx.
## ---------------------------------------------------------------------------
harm_out_path <- file.path(out_dir, "SM_2_generated.xlsx")

## ---- 2H.5a  the two mandatory derived tables, written as their own files ----
## Table_SM_2.5 and Table_SM_2.6 are outputs of this pipeline in exactly the same
## sense as Table_SM_2.4 or Table_SM_2.7, and are written as standalone workbooks
## under their own names.  Each carries its own rule audit as a second sheet, so
## the provenance of every harmonised value travels with the table it belongs to.
harm_readme <- function(tbl, src, n_rows, aud) data.frame(Notes = c(
  paste0(tbl, " - DERIVED table, generated by analysis_pipeline_2026.R."),
  paste0("Constructed from ", src, " by the harmonisation rules (i)-(vii) of Section 2.3."),
  paste0("Rows: ", n_rows, ", one per row of ", src, ", in the same order, so every value",
         " traces to its raw cell by row number."),
  "",
  "Sheet 'harmonisation_audit' records, for every non-empty raw cell: the source row,",
  "  compound, reference, variable, the entry as written, the rule applied, the value",
  "  produced, and whether it was retained or excluded.",
  "Sheet 'parameters' records the settings that rules (i)-(vii) do not fully determine.",
  "",
  paste0("Raw cells classified: ", nrow(aud),
         " | retained: ", sum(aud$status == "retained"),
         " | excluded: ", sum(aud$status == "excluded")),
  "",
  "This table is regenerated on every run and should not be edited by hand.",
  "Corrections belong in the raw compilation, not here."),
  stringsAsFactors = FALSE)

write_derived <- function(sheet_name, df, src_sheet, aud, file) {
  wb <- openxlsx::createWorkbook()
  ## the data sheet carries the header on row 1 and no caption, so this file can
  ## itself be read back by read_sm2() without a skip; provenance lives on the
  ## README and harmonisation_audit sheets instead of in a caption cell
  openxlsx::addWorksheet(wb, sheet_name)
  openxlsx::writeData(wb, sheet_name, df, startRow = 1, keepNA = TRUE, na.string = NA_TOKEN)
  openxlsx::addWorksheet(wb, "README")
  openxlsx::writeData(wb, "README", harm_readme(sheet_name, src_sheet, nrow(df), aud))
  openxlsx::addWorksheet(wb, "harmonisation_audit")
  openxlsx::writeData(wb, "harmonisation_audit", aud)
  openxlsx::addWorksheet(wb, "parameters")
  openxlsx::writeData(wb, "parameters", harm_params)
  tryCatch(openxlsx::saveWorkbook(wb, file, overwrite = TRUE),
           error = function(e) stop("Section 2H: cannot write ", basename(file), ".\n",
                                    "  ", conditionMessage(e), "\n",
                                    "  If the file is open in Excel, close it and re-run.",
                                    call. = FALSE))
  wrote(basename(file))
}

aud_prop <- harm_audit[harm_audit$source_sheet == SH_RAW_PROP, , drop = FALSE]
aud_fate <- harm_audit[harm_audit$source_sheet == SH_RAW_FATE, , drop = FALSE]

write_derived(SH_HARM_PROP, harm_prop_new, SH_RAW_PROP, aud_prop,
              file.path(out_dir, "Table_SM_2.5.xlsx"))
write_derived(SH_HARM_FATE, harm_fate_new, SH_RAW_FATE, aud_fate,
              file.path(out_dir, "Table_SM_2.6.xlsx"))

## ---- 2H.5b  the working workbook the rest of the pipeline reads -------------
## A copy of the input workbook with the two derived tables inserted.  Every
## other sheet is carried through unchanged, so the file is a drop-in replacement
## for SM_2.xlsx.  The two derived sheets are ADDED UNCONDITIONALLY: the input
## workbook is no longer expected to contain them, and building the copy by
## replacing sheets that happen to be present would leave them absent.
wb_gen    <- openxlsx::createWorkbook()
sheets    <- readxl::excel_sheets(sm2_path)
harm_keys <- c(SHEET_KEY(SH_HARM_PROP), SHEET_KEY(SH_HARM_FATE))

for (sh in sheets) {
  if (SHEET_KEY(sh) %in% harm_keys) next          # superseded by the derived table
  dat <- tryCatch(readxl::read_excel(sm2_path, sheet = sh, col_types = "text",
                                     col_names = FALSE, .name_repair = "minimal"),
                  error = function(e) NULL)
  if (is.null(dat)) next
  openxlsx::addWorksheet(wb_gen, sh)
  openxlsx::writeData(wb_gen, sh, dat, colNames = FALSE)
}

for (spec in list(list(nm = SH_HARM_PROP, df = harm_prop_new),
                  list(nm = SH_HARM_FATE, df = harm_fate_new))) {
  openxlsx::addWorksheet(wb_gen, spec$nm)
  ## header on row 1 (no caption row): read_sm2()/header_skip() locate the header
  ## by the first fully-populated row, so a one-cell caption above it would be
  ## mistaken for that header.
  openxlsx::writeData(wb_gen, spec$nm, spec$df, startRow = 1,
                      keepNA = TRUE, na.string = NA_TOKEN)
}

tryCatch(openxlsx::saveWorkbook(wb_gen, harm_out_path, overwrite = TRUE),
         error = function(e) stop("Section 2H: cannot write ", basename(harm_out_path),
                                  ".\n  ", conditionMessage(e),
                                  "\n  If the file is open in Excel, close it and re-run.",
                                  call. = FALSE))
wrote(basename(harm_out_path))

## ---- THE REDIRECT ----------------------------------------------------------
## From here on every read of the workbook resolves to the generated file.  The
## raw sheets inside it are verbatim copies, so the censoring census is computed
## on the same input as before; everything downstream consumes the derived
## Table_SM_2.5 and Table_SM_2.6 that this section built.
sm2_path <- harm_out_path
HDR$skip <- list()          # clear the cached header offsets for the new file

## ---- 2H.6  PRE-CONDITION GATE ----------------------------------------------
## Nothing downstream may run unless both derived tables exist and are readable.
## Failing loudly here is deliberate: a missing harmonised table would otherwise
## surface much later as an "object not found" error.
gen_sheets <- tryCatch(readxl::excel_sheets(sm2_path), error = function(e) character(0))
missing_gen <- c(SH_HARM_PROP, SH_HARM_FATE)[
  !SHEET_KEY(c(SH_HARM_PROP, SH_HARM_FATE)) %in% SHEET_KEY(gen_sheets)]
if (length(missing_gen))
  stop("Section 2H: ", paste(missing_gen, collapse = " and "),
       " absent from ", basename(sm2_path), ".\n",
       "  The derived tables were not written, so no downstream analysis can run.",
       call. = FALSE)

for (sh in c(SH_HARM_PROP, SH_HARM_FATE)) {
  chk <- tryCatch(read_sm2(sh), error = function(e) NULL)
  if (is.null(chk) || !nrow(chk))
    stop("Section 2H: ", sh, " was written but cannot be read back.\n",
         "  Delete ", basename(harm_out_path), " and re-run.", call. = FALSE)
}

base::message("Section 2H: ", SH_HARM_PROP, " (", nrow(harm_prop_new), " rows) and ",
              SH_HARM_FATE, " (", nrow(harm_fate_new), " rows) generated and verified.")
base::message("Section 2H: sm2_path re-pointed to ", basename(harm_out_path),
              " - all downstream tables and figures now derive from the raw compilations.")





## ===========================================================================
## 3A. CENSORING REPORT  (from the RAW tables Table_SM_2.2 / Table_SM_2.3)
##     WHAT: classify every raw cell of the analysis variables and report the
##           left-censoring rate per variable and per group.
## ===========================================================================
## classify ONE raw cell per the harmonisation rules
## A cell is LEFT-censored ONLY if it carries a censor symbol (n.d./<LOQ/<x) AND has
## NO detected upper value.  A reported range with a detected upper bound -- including
## a censored-floor range such as "n.d.-y" or "<x-y" -- is a DETECTION (rule ii,
## midpoint), NOT a non-detect.  This is the key reconciliation: counting n.d.-y / <x-y
## as LEFT inflates the censoring rate and masks real detections.
classify_cell <- function(raw) {
  if (is.na(raw)) return("EXCL")
  s <- trimws(as.character(raw))
  if (s == "" || s == "-") return("EXCL")
  if (grepl(">>", s, fixed = TRUE) || grepl("<<", s, fixed = TRUE)) return("EXCL") # rule v: >>x,<<x
  if (tolower(gsub("\\s+", "", s)) %in% c("n.a.", "n.a", "na")) return("EXCL")      # rule v: N.A.
  t  <- gsub("[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]", "-", s, perl = TRUE)   # unify dashes
  t2 <- gsub("\\s*[\u00d7x\u00b7*]\\s*10\\s*\\^?\\s*(?=[+-]?[0-9])", "e",           # x10^-n -> e-n
             t, perl = TRUE, ignore.case = TRUE)
  nums <- suppressWarnings(as.numeric(regmatches(t2,
            gregexpr("[0-9]*\\.?[0-9]+(?:[eE][+-]?[0-9]+)?", t2, perl = TRUE))[[1]]))
  nums <- nums[is.finite(nums)]
  has_sep <- grepl("-", sub("^-", "", gsub("[eE][+-]?[0-9]+", "", t2, perl = TRUE)))
  censor  <- grepl("n\\.?d\\.?", t, ignore.case = TRUE) || grepl("<", t, fixed = TRUE)
  if (has_sep && length(nums) >= 1) return("NUMERIC")   # rule ii: a-b, n.d.-y, <x-y  -> DETECTED
  if (censor)                       return("LEFT")      # rule i : pure n.d., <LOQ, <x -> censored
  if (grepl(">", t, fixed = TRUE))  return("RIGHT")     # rule iii: >x
  if (grepl("[\u2245\u2248~]", s) || grepl("\u00B1", s)) return("APPROX")           # rule iv: ~x, x+/-SD
  "NUMERIC"                                                                         # plain value
}

## tally one set of columns (read as TEXT so qualifiers survive)
censor_tally <- function(path, sheet, patterns, group) {
  raw <- read_sm2(sheet)                    # encoding-normalised (Section 1b)
  raw <- apply_corrections(raw, sheet)      # currently a no-op (Section 1c)
  nms <- norm_names(names(raw))
  do.call(rbind, lapply(names(patterns), function(canon) {
    idx  <- find_col(nms, patterns[[canon]], canon)
    cats <- vapply(raw[[idx]], classify_cell, character(1))
    nex  <- sum(cats == "EXCL"); nL <- sum(cats == "LEFT"); nR <- sum(cats == "RIGHT")
    nan  <- length(cats) - nex                                  # analysed observations
    data.frame(group = group, variable = canon, n_analyzed = nan,
               n_left_censored = nL,
               left_censoring_pct = if (nan) round(100 * nL / nan, 2) else NA_real_,
               n_right_censored = nR, n_excluded = nex,
               stringsAsFactors = FALSE)
  }))
}

## Censoring is assessed only on genuinely measured, multi-source variables.
## Molecular weight is EXCLUDED from the censoring/census: it is a single,
## compound-intrinsic look-up value (one figure per compound, simply repeated
## down each compound block) and is never reported below a detection limit, so
## it carries no censoring information and would only inflate the denominator.
## The compound identifier, the canonical-SMILES and clinical-use text fields
## and the reference tag are likewise not analysis variables.  Concretely this
## assesses columns E:I of Table_SM_2.2 (pKa, log KOW, log KOC, log DOW, log S)
## and columns C:L of Table_SM_2.3 (the four concentrations and six half-lives)
## -> 5 physicochemical + 10 environmental-fate = 15 censoring variables.
cens_prop_patterns <- prop_patterns[setdiff(names(prop_patterns), "Molecular_weight")]
cens_by_var <- rbind(
  censor_tally(sm2_path, SH_RAW_PROP, cens_prop_patterns, "Properties"),
  censor_tally(sm2_path, SH_RAW_FATE, conc_patterns, "Concentrations"),
  censor_tally(sm2_path, SH_RAW_FATE, hl_patterns,   "Half-lives"))

## group + overall summaries
cens_by_group <- cens_by_var %>%
  dplyr::group_by(group) %>%
  dplyr::summarise(n_analyzed = sum(n_analyzed),
                   n_left_censored = sum(n_left_censored),
                   left_censoring_pct = round(100 * sum(n_left_censored) / sum(n_analyzed), 2),
                   .groups = "drop")
overall_pct <- round(100 * sum(cens_by_var$n_left_censored) / sum(cens_by_var$n_analyzed), 2)
cens_overall <- data.frame(scope = "Analysed variables (overall)",
                           n_analyzed = sum(cens_by_var$n_analyzed),
                           n_left_censored = sum(cens_by_var$n_left_censored),
                           left_censoring_pct = overall_pct)


## ===========================================================================
## 3B. DATASET CENSUS  -> Table_SM_2.4, sheets census_*
STAGE("3B  dataset census -> Table_SM_2.4")

## ===========================================================================
## WHAT: a complete, reproducible census of the ANALYSED dataset (the 15
##       censoring variables), so that every number quoted in Section 3.1 of the
##       manuscript is traceable to this pipeline.
## PARAMETERS: every column of Table_SM_2.2 / Table_SM_2.3 that is not an
##       identifier (Compound, SMILES, Clinical Application), not molecular
##       weight (a single non-representative look-up value), not the reference
##       tag, and not a DERIVED column (any header starting with "Computed").
##       -> 5 physicochemical + 10 environmental-fate = 15 analysed parameters.
## REPORTING CONVENTION (this is the one that must be stated in the paper):
##       a cell carrying a censor symbol is counted as LEFT-CENSORED even when it
##       is a censored-floor range (n.d.-y, <x-y).  Such ranges are nevertheless
##       treated as DETECTIONS by the harmonisation (rule ii) and therefore do
##       NOT receive DL/2 substitution; the census reports both sub-counts so the
##       two figures can never be confused:
##         left_censored              = pure non-detect + censored-floor range
##         of_which_non_detect        = the ones actually DL/2-substituted
##                                      (this is the 197/3370 figure of Table_SM_2.4)
##       "Quantitative non-standard" = ranges + approximate values.
CENSUS_DROP <- "SMILES|^(MW|Molecular weight|Compound|Clinical|References|Computed)"

census_cell <- function(raw) {
  if (is.na(raw)) return("EMPTY")
  s <- trimws(as.character(raw))
  if (s == "" || s == "-") return("EMPTY")
  if (grepl(">>", s, fixed = TRUE) || grepl("<<", s, fixed = TRUE)) return("EXCL_QUAL")
  if (tolower(gsub("\\s+", "", s)) %in% c("n.a.", "n.a", "na")) return("EXCL_QUAL")
  t  <- gsub("[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]", "-", s, perl = TRUE)
  t2 <- gsub("\\s*[\u00d7x\u00b7*]\\s*10\\s*\\^?\\s*(?=[+-]?[0-9])", "e",
             t, perl = TRUE, ignore.case = TRUE)
  has_sep <- grepl("-", sub("^-", "", gsub("[eE][+-]?[0-9]+", "", t2, perl = TRUE)))
  censor  <- grepl("n\\.?d\\.?", t, ignore.case = TRUE) || grepl("<", t, fixed = TRUE)
  if (censor && has_sep)           return("LEFT_RANGE")
  if (censor)                      return("LEFT_ND")
  if (grepl(">", t, fixed = TRUE)) return("RIGHT")
  if (has_sep)                     return("RANGE")
  if (grepl("[\u2245\u2248~]", s) || grepl("\u00B1", s)) return("APPROX")
  "STANDARD"
}

census_sheet <- function(sheet, group) {
  raw  <- apply_corrections(read_sm2(sheet), sheet)
  nms  <- norm_names(names(raw))
  keep <- nzchar(nms) & !grepl(CENSUS_DROP, nms, ignore.case = TRUE)
  do.call(rbind, lapply(which(keep), function(j) {
    cats <- vapply(raw[[j]], census_cell, character(1))
    data.frame(
      group = group, parameter = nms[j], n_rows = length(cats),
      entries                 = sum(cats != "EMPTY"),
      standard                = sum(cats == "STANDARD"),
      range                   = sum(cats == "RANGE"),
      approximate             = sum(cats == "APPROX"),
      left_censored           = sum(cats %in% c("LEFT_ND", "LEFT_RANGE")),
      of_which_non_detect     = sum(cats == "LEFT_ND"),
      of_which_censored_range = sum(cats == "LEFT_RANGE"),
      right_censored          = sum(cats == "RIGHT"),
      excluded_qualitative    = sum(cats == "EXCL_QUAL"),
      stringsAsFactors = FALSE)
  }))
}

census_by_param <- rbind(census_sheet(SH_RAW_PROP, "Physicochemical properties"),
                         census_sheet(SH_RAW_FATE, "Environmental fate"))
census_by_param$quantitative_non_standard <- census_by_param$range + census_by_param$approximate
census_by_param$non_standard <- with(census_by_param,
  quantitative_non_standard + left_censored + right_censored + excluded_qualitative)

census_by_group <- census_by_param %>%
  dplyr::group_by(group) %>%
  dplyr::summarise(
    n_parameters = dplyr::n(), n_rows = max(n_rows), n_cells = dplyr::n() * max(n_rows),
    entries = sum(entries), standard = sum(standard),
    quantitative_non_standard = sum(quantitative_non_standard),
    left_censored = sum(left_censored),
    of_which_non_detect = sum(of_which_non_detect),
    of_which_censored_range = sum(of_which_censored_range),
    right_censored = sum(right_censored),
    excluded_qualitative = sum(excluded_qualitative),
    non_standard = sum(non_standard),
    non_standard_pct = round(100 * sum(non_standard) / sum(entries), 2),
    .groups = "drop")

CENS_N <- sum(census_by_param$entries)
pc     <- function(k) round(100 * k / CENS_N, 2)
census_overall <- data.frame(
  metric = c("Compiled parameters",
             "Cells (parameters x records)",
             "Entries (populated cells)",
             "Standard single values",
             "Quantitative non-standard (ranges + approximate)",
             "   of which ranges (rule ii)",
             "   of which approximate / +-SD (rule iv)",
             "Left-censored (n.d. / <x, censored-floor ranges included)",
             "   of which pure non-detect -> DL/2 substituted (rule i)",
             "   of which censored-floor range -> treated as detection (rule ii)",
             "Right-censored (>x, rule iii)",
             "Qualitative extremes excluded (>>x, <<x, N.A.; rule v)",
             "NON-STANDARD ENTRIES (total)"),
  n = c(sum(census_by_group$n_parameters),
        sum(census_by_group$n_cells), CENS_N,
        sum(census_by_param$standard),
        sum(census_by_param$quantitative_non_standard),
        sum(census_by_param$range), sum(census_by_param$approximate),
        sum(census_by_param$left_censored),
        sum(census_by_param$of_which_non_detect),
        sum(census_by_param$of_which_censored_range),
        sum(census_by_param$right_censored),
        sum(census_by_param$excluded_qualitative),
        sum(census_by_param$non_standard)),
  stringsAsFactors = FALSE)
census_overall$pct_of_entries <- c(NA, NA, 100,
  pc(census_overall$n[4]), pc(census_overall$n[5]), pc(census_overall$n[6]),
  pc(census_overall$n[7]), pc(census_overall$n[8]), pc(census_overall$n[9]),
  pc(census_overall$n[10]), pc(census_overall$n[11]), pc(census_overall$n[12]),
  pc(census_overall$n[13]))

## ---------------------------------------------------------------------------
## COLUMN-COMPLETENESS AUDIT  (merged in from the standalone census script)
## Confirms that every one of the 15 analysed variables was actually located and
## carries data.  Its purpose is defensive: if a column were renamed in SM_2 and
## silently failed to match, the census would quietly shrink instead of erroring.
## Because it is derived from census_by_param it uses exactly the same
## classifier as the census itself, so the two can never disagree.
## ---------------------------------------------------------------------------
census_completeness <- data.frame(
  source_table  = ifelse(census_by_param$group == "Physicochemical properties",
                         SH_RAW_PROP, SH_RAW_FATE),
  variable      = census_by_param$parameter,
  n_rows        = census_by_param$n_rows,
  populated     = census_by_param$entries,
  pct_populated = round(100 * census_by_param$entries /
                          pmax(census_by_param$n_rows, 1), 2),
  is_empty      = as.integer(census_by_param$entries == 0),
  stringsAsFactors = FALSE)

if (nrow(census_completeness) != 15L)
  stop(sprintf(paste0("Census covers %d variables, expected 15 (5 physicochemical + ",
                      "10 environmental-fate). A column in Table_SM_2.2/2.3 has probably ",
                      "been renamed."), nrow(census_completeness)))
if (any(census_completeness$is_empty == 1L))
  stop("Analysed variable(s) with no data: ",
       paste(census_completeness$variable[census_completeness$is_empty == 1L],
             collapse = ", "))

wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, "by_variable"); openxlsx::writeData(wb, "by_variable", cens_by_var)
openxlsx::addWorksheet(wb, "by_group");    openxlsx::writeData(wb, "by_group",    cens_by_group)
openxlsx::addWorksheet(wb, "overall");     openxlsx::writeData(wb, "overall",     cens_overall)
openxlsx::addWorksheet(wb, "census_by_parameter"); openxlsx::writeData(wb, "census_by_parameter", census_by_param)
openxlsx::addWorksheet(wb, "census_by_group");     openxlsx::writeData(wb, "census_by_group",     census_by_group)
openxlsx::addWorksheet(wb, "census_overall");      openxlsx::writeData(wb, "census_overall",      census_overall)
openxlsx::addWorksheet(wb, "column_completeness"); openxlsx::writeData(wb, "column_completeness", census_completeness)
openxlsx::saveWorkbook(wb, file.path(out_dir, "Table_SM_2.4.xlsx"), overwrite = TRUE)
wrote("Table_SM_2.4.xlsx")


## ===========================================================================
## 3D. READ HARMONISED DATA (Table_SM_2.5 / Table_SM_2.6), keep only the analysis variables
## ===========================================================================
id_pat_class <- "^PCs"; id_pat_comp <- "^Compound"

read_harmonized <- function(path, sheet, patterns) {
  raw <- read_sm2(sheet)                    # encoding-normalised (Section 1b)
  raw <- apply_corrections(raw, sheet)      # currently a no-op (Section 1c)
  nms <- norm_names(names(raw))
  out <- tibble::tibble(
    PCs_class = trimws(as.character(raw[[ find_col(nms, id_pat_class, "PCs_class") ]])),
    Compound  = trimws(as.character(raw[[ find_col(nms, id_pat_comp,  "Compound")  ]])))
  ## STRICT NUMERIC RECAST of the analysis columns of part (B): every non-empty
  ## cell must parse to a number.  Any that does not is logged (Section 12) and
  ## the run reports it, so a silent NA can never enter a published statistic.
  for (canon in names(patterns))
    out[[canon]] <- to_numeric_strict(raw[[ find_col(nms, patterns[[canon]], canon) ]],
                                      sheet, canon)
  ## ---- drop rows that carry NO compound ------------------------------------
  ## Table_SM_2.5 contains one spacer row (row 237) holding a therapeutic class
  ## and nothing else: no compound, no reference, no value.  Such a row carries
  ## no datum and is excluded exactly as harmonisation rule (v) excludes
  ## non-quantitative entries, keeping the compound count at 48.  Any row dropped
  ## is reported, so this can never hide data.
  empty_id <- is.na(out$Compound) | out$Compound == ""
  if (any(empty_id)) {
    vals <- out[empty_id, names(patterns), drop = FALSE]
    n_val <- sum(vapply(vals, function(v) sum(is.finite(v)), integer(1)))
    msg("  ", sheet, ": ", sum(empty_id), " row(s) dropped with no Compound (spacer rows); ",
            "they carried ", n_val, " numeric value(s)",
            if (n_val > 0) "  *** CHECK: a dropped row held data ***" else " - no data lost.")
    if (n_val > 0)
      msg(sheet, ": a row with no Compound carried numeric data and was dropped. ",
              "Inspect the sheet.", call. = FALSE)
    out <- out[!empty_id, , drop = FALSE]
  }
  out
}

prop <- read_harmonized(sm2_path, SH_HARM_PROP, prop_patterns)
fate <- read_harmonized(sm2_path, SH_HARM_FATE, fate_patterns)

## --- D18 CLASS-LABEL CANONICALISATION (applied at source) -------------------
## Force one spelling for every therapeutic-class label, so that ALL tables
## (SM_3, SM_4, Table_SM_2.9/2.9_a, Table_SM_2.10) AND all figures use the same
## label regardless of how the input sheets spell it. Canonical forms match the
## manuscript prose and SM_1 ("Anti-inflammatory", "Antibiotic from other class").
canon_class_chr <- function(x) dplyr::recode(trimws(as.character(x)),
      "Anti-Inflammatory"             = "Anti-inflammatory",
      "anti-inflammatory"             = "Anti-inflammatory",
      "Anti-inflammatories"           = "Anti-inflammatory",
      "Other classes of Antibiotics"  = "Antibiotic from other class",
      "Antibiotics from other class"  = "Antibiotic from other class",
      "Antibiotic from other classes" = "Antibiotic from other class")
prop$PCs_class <- canon_class_chr(prop$PCs_class)
fate$PCs_class <- canon_class_chr(fate$PCs_class)
msg("D18: therapeutic-class labels canonicalised in prop and fate.")

## ---- numeric-recast report for part (B) ------------------------------------
num_report <- do.call(rbind, NUM$report)
num_bad    <- if (length(NUM$bad)) do.call(rbind, NUM$bad) else
  data.frame(sheet = character(0), column = character(0),
             row_in_sheet = integer(0), value = character(0))
msg("Numeric recast (part B): ", nrow(num_report), " analysis columns; ",
        sum(num_report$n_numeric), " values; ",
        sum(num_report$n_unparsable), " unparsable cell(s).")
if (sum(num_report$n_unparsable) > 0) {
  msg("Unparsable cells remain in the analysis columns.",
          call. = FALSE)
  if (!QUIET) print(num_bad, row.names = FALSE)
} else {
  msg("  -> all analysis columns are fully numeric after normalisation; ",
          "no value is silently lost.")
}

## --- ZERO POLICY (granular) ------------------------------------------------
ZERO_DROP <- "half_life"     # "half_life" | "conc" | c("conc","half_life") | character(0)
conc_cols <- names(conc_patterns); hl_cols <- names(hl_patterns)
drop_cols <- c(if ("conc"      %in% ZERO_DROP) conc_cols,
               if ("half_life" %in% ZERO_DROP) hl_cols)
N_ZERO_DROPPED <- 0L
if (length(drop_cols)) {
  ## count before removing, so the integrity census (Section 12b) can state how
  ## many records the zero policy actually removed rather than asserting six
  N_ZERO_DROPPED <- sum(vapply(intersect(drop_cols, names(fate)),
                        function(cc) sum(!is.na(fate[[cc]]) & fate[[cc]] == 0),
                        integer(1)))
  fate <- dplyr::mutate(fate, dplyr::across(dplyr::any_of(drop_cols),
            ~ ifelse(!is.na(.) & . == 0, NA_real_, .)))
  msg("ZERO_DROP = {", paste(ZERO_DROP, collapse = ", "),
          "}: ", N_ZERO_DROPPED, " literal zero(s) removed from ",
          length(drop_cols), " fate column(s).")
} else msg("ZERO_DROP = none: all zeros retained.")


## ===========================================================================
## 3C. SUBSTITUTION SENSITIVITY for ALL censored FATE variables (DL/2 vs DL vs exclusion)
STAGE("3C  substitution sensitivity -> Table_SM_2.7")

##     WHAT: every left-censored fate variable -- the four concentrations AND the six
##           half-lives -- is re-summarised (P25/median/P75/P95) under three treatments
##           of below-limit data, classifying EVERY raw entry by the harmonisation
##           rules -- not just one case:
##             rule i   n.d. / < LOQ / <x   (pure left-censored, no detected value)
##             rule ii  a-b , n.d.-y , <x-y (range -> midpoint; a detected upper bound
##                                           makes n.d.-y / <x-y DETECTIONS, not non-detects)
##             rule iii >x                  (right-censored -> x)
##             rule iv  ~x , x+/-SD         (approximate -> x)
##             rule v   >>x , <<x , N.A.    (excluded)
##           Variables with no censored entries simply show four identical scenarios (a
##           built-in robustness check).  PROPERTIES are excluded by design: they are not
##           detection-limited in the DL sense, their values can be negative (so the range
##           parser is invalid for them), and their raw left-censoring is negligible (~0.25%).
##

## ===========================================================================
NOLIMIT_DL <- "min_detected"   # proxy limit for a no-limit n.d./<LOQ in the SUBSTITUTION
                     # cross-check only (per-variable lowest detected value; Helsel 2012). No fixed
                     # 1e-4 constant. Concentration statistics OF RECORD are the censored estimates
                     # of Section 6C (the Conc_* sheets of SM_4_fate.xlsx), not this DL/2 reconstruction.

## parse one raw concentration cell -> list(kind, lo, hi).  Concentrations are
## non-negative, so every "-" (after protecting sci-notation exponents) is a
## range separator, never a minus sign.
parse_conc <- function(s) {
  out <- function(kind, lo = NA_real_, hi = NA_real_) list(kind = kind, lo = lo, hi = hi)
  if (is.na(s)) return(out("EXCL"))
  s <- trimws(as.character(s))
  if (s == "" || s == "-") return(out("EXCL"))
  if (grepl(">>", s, fixed = TRUE) || grepl("<<", s, fixed = TRUE)) return(out("EXCL"))
  low <- tolower(s)
  if (gsub("\\s+", "", low) %in% c("n.a.", "n.a", "na")) return(out("EXCL"))
  nd  <- grepl("n\\.?d\\.?", low) && !grepl("n\\.a", low)
  loq <- grepl("lo[dq]", low)
  lt  <- grepl("<", s, fixed = TRUE)
  gt  <- grepl(">", s, fixed = TRUE)
  apx <- grepl("[\u2245\u2248~]", s) || grepl("\u00B1", s, fixed = TRUE)
  t <- gsub("[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]", "-", s, perl = TRUE)            # unify dashes
  t <- gsub("\\s*[\u00d7x\u00b7*]\\s*10\\s*\\^?\\s*(?=[+-]?[0-9])", "e",                    # x10^-n -> e-n
            t, perl = TRUE, ignore.case = TRUE)
  ## pull EVERY numeric token (robust to x +/- y, sci-notation, ranges); a leading
  ## '<'/'>' or a separating '-' is never matched as a sign because the pattern has none.
  nums <- suppressWarnings(as.numeric(regmatches(t,
            gregexpr("[0-9]*\\.?[0-9]+(?:[eE][+-]?[0-9]+)?", t, perl = TRUE))[[1]]))
  nums <- nums[is.finite(nums)]
  ## separator present?  strip exponents first so the '-' inside e-5 is not counted.
  has_sep <- grepl("-", sub("^-", "", gsub("[eE][+-]?[0-9]+", "", t, perl = TRUE)))
  if ((nd || loq) && !lt && length(nums) == 0)     return(out("NOLIM"))             # n.d. / (rare) LOQ
  if (lt && length(nums) == 0)                      return(out("NOLIM"))             # < LOQ
  if ((nd || loq) && has_sep && length(nums) >= 1)  return(out("MIX_nolim", hi = max(nums))) # n.d.-y
  if (lt && has_sep && length(nums) >= 2)           return(out("MIX_lim", lo = min(nums), hi = max(nums))) # <x-y
  if (lt && has_sep && length(nums) == 1)           return(out("MIX_nolim", hi = max(nums))) # <LOQ-y
  if (lt && length(nums) >= 1)  return(out("LT_lim", lo = nums[1]))                 # <x
  if (gt && length(nums) >= 1)  return(out("GT",     lo = nums[1]))                 # >x
  if (apx && length(nums) >= 1) return(out("APPROX", lo = nums[1]))                 # ~x , x+/-SD
  if (has_sep && length(nums) >= 2) return(out("RANGE", lo = min(nums), hi = max(nums)))      # a-b
  if (length(nums) >= 1) return(out("NUM", lo = nums[1]))                           # plain value
  out("EXCL")
}

tryCatch({
  q4 <- function(v) { v <- v[is.finite(v)]
    if (length(v)) stats::quantile(v, c(.25, .5, .75, .95), names = FALSE, type = 7)
    else rep(NA_real_, 4) }
  r6 <- function(x) signif(x, 4)

  raw_fate  <- apply_corrections(read_sm2(SH_RAW_FATE),  SH_RAW_FATE)
  harm_fate <- apply_corrections(read_sm2(SH_HARM_FATE), SH_HARM_FATE)
  rnms <- norm_names(names(raw_fate)); hnms <- norm_names(names(harm_fate))

  ## display units (concentrations vs half-lives) for labelling the output table
  fate_unit <- c(Concentration_SW = "ug/L",  Concentration_SED = "ug/Kg",
                 Concentration_HZ = "ug/L",  Concentration_GW  = "ug/L",
                 half_life_SW = "days", half_life_SED = "days", half_life_HZ = "days",
                 half_life_aerobic = "days", half_life_anoxic = "days", half_life_GW = "days")

  sens <- list(); cnts <- list()
  for (cc in names(fate_patterns)) {          # ALL censored fate variables: 4 conc + 6 half-lives
    ## zero policy applied consistently with Section 3B / Sections 4-6 (e.g. a half-life of
    ## exactly 0 is physically impossible and is dropped under ZERO_DROP = "half_life").
    dz <- (cc %in% conc_cols && "conc"      %in% ZERO_DROP) ||
          (cc %in% hl_cols   && "half_life" %in% ZERO_DROP)

    P    <- lapply(raw_fate[[ find_col(rnms, fate_patterns[[cc]], cc) ]], parse_conc)
    kind <- vapply(P, `[[`, character(1), "kind")
    lo   <- vapply(P, `[[`, numeric(1),  "lo")
    hi   <- vapply(P, `[[`, numeric(1),  "hi")

    ## Core detections (plain values, >x, approx, true ranges), kept in every scenario.
    ## posd/lims below are used only by the optional data-derived proxy keywords.
    det_core <- c(lo[kind == "NUM"], lo[kind == "GT"], lo[kind == "APPROX"],
                  (lo[kind == "RANGE"] + hi[kind == "RANGE"]) / 2)
    posd  <- det_core[det_core > 0]; lims <- lo[kind == "LT_lim"]
    proxy <- if (is.numeric(NOLIMIT_DL)) NOLIMIT_DL else switch(NOLIMIT_DL,
               min_detected = if (length(posd)) min(posd) else NA_real_,
               min_limit    = if (length(lims)) min(lims) else if (length(posd)) min(posd) else NA_real_,
               median_limit = if (length(lims)) stats::median(lims) else NA_real_,
               if (length(posd)) min(posd) else NA_real_)

    ## DETECTED (kept in every scenario): core detections PLUS censored-floor ranges
    ## (rule ii). A censored lower bound is substituted under rule (i) and averaged with
    ## the measured upper bound b:  n.d.-b -> (DL/2 + b)/2 ;  <x-b -> (x/2 + b)/2.
    detected <- c(det_core,
                  (proxy / 2          + hi[kind == "MIX_nolim"]) / 2,   # n.d.-b : (DL/2 + b)/2
                  (lo[kind == "MIX_lim"] / 2 + hi[kind == "MIX_lim"]) / 2)  # <x-b  : (x/2  + b)/2

    ## CENSORED (vary across scenarios): pure <x (limit known) and pure n.d./<LOQ (no limit).
    sub_dl2 <- c(lo[kind == "LT_lim"] / 2, rep(proxy / 2, sum(kind == "NOLIM")))
    sub_dl  <- c(lo[kind == "LT_lim"],     rep(proxy,     sum(kind == "NOLIM")))

    pub <- clean_num(harm_fate[[ find_col(hnms, fate_patterns[[cc]], cc) ]])

    ## drop literal zeros when the zero policy says so, so the sensitivity is computed on the
    ## SAME data as the published statistics in Sections 4-6.
    add <- function(scn, v) { if (dz) v <- v[is.finite(v) & v != 0]
      data.frame(Variable = cc, Unit = fate_unit[[cc]], Scenario = scn,
              n = sum(is.finite(v)), P25 = r6(q4(v)[1]), Median = r6(q4(v)[2]),
              P75 = r6(q4(v)[3]), P95 = r6(q4(v)[4]), stringsAsFactors = FALSE) }
    sens[[length(sens) + 1]] <- add("DL/2 (published, Table_SM_2.6)", pub)
    sens[[length(sens) + 1]] <- add("DL/2 (reconstructed)",    c(detected, sub_dl2))
    sens[[length(sens) + 1]] <- add("DL (reconstructed)",      c(detected, sub_dl))
    sens[[length(sens) + 1]] <- add("Exclusion (detected only)", detected)

    ncl <- sum(kind == "LT_lim"); nnl <- sum(kind == "NOLIM")   # only PURE censored entries
    nan <- length(detected) + ncl + nnl
    cnts[[length(cnts) + 1]] <- data.frame(Variable = cc, Unit = fate_unit[[cc]], n_analyzed = nan,
              n_detected = length(detected), n_censored_with_limit = ncl,
              n_censored_no_limit = nnl,
              left_censoring_pct = if (nan) round(100 * (ncl + nnl) / nan, 1) else NA_real_,
              proxy_DL_used = r6(proxy), stringsAsFactors = FALSE)
  }
  comp_tbl <- do.call(rbind, sens); cnt_tbl <- do.call(rbind, cnts)

  notes <- data.frame(Notes = c(
    "SUBSTITUTION SENSITIVITY for ALL censored fate variables (4 concentrations + 6 half-lives)",
    "Units: concentrations SW/HZ/GW = ug/L, SED = ug/Kg; half-lives = days.",
    paste0("No-limit non-detect proxy (NOLIMIT_DL) = ",
           if (is.numeric(NOLIMIT_DL)) NOLIMIT_DL else NOLIMIT_DL),
    paste0("Zero policy (ZERO_DROP) = {", paste(ZERO_DROP, collapse = ", "),
           "} applied to the matching variable group, as in Sections 4-6."),
    "This table is the SUBSTITUTION cross-check only. The concentration statistics OF RECORD are the",
    "   censored estimates (KM primary; Turnbull/ROS/DL2 cross-checks) in Section 6C ->  (Conc_* sheets of SM_4_fate.xlsx)",
    "   the Conc_* sheets of SM_4_fate.xlsx. The scenarios below bound the substitution conventions using a",
    "   per-variable lowest-detected proxy; they are NOT the reported estimate. No-limit non-detects are no",
    "   longer substituted with a fixed 1e-4 constant (retired).",
    "Read robustness from 'DL/2 (published)' vs 'Exclusion (detected)'. A large gap (e.g. SED)",
    "   means the summary is governed by the non-detect treatment and is NOT a robust estimate.",
    "A variable whose four scenarios are identical has no censored entries (robustness confirmed).",
    "PROPERTIES are excluded by design: not detection-limited, can be negative (range parser invalid),",
    "   and their raw left-censoring is negligible (~0.25%)."),
    stringsAsFactors = FALSE)

  wb_s <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb_s, "Sensitivity")
  openxlsx::writeData(wb_s, "Sensitivity", notes, startRow = 1)
  r2 <- nrow(notes) + 3
  openxlsx::writeData(wb_s, "Sensitivity", comp_tbl, startRow = r2)
  openxlsx::writeData(wb_s, "Sensitivity", cnt_tbl, startRow = r2 + nrow(comp_tbl) + 3)
  openxlsx::saveWorkbook(wb_s, file.path(out_dir, "Table_SM_2.7.xlsx"), overwrite = TRUE)
  wrote("Table_SM_2.7.xlsx")
  msg("Wrote Table_SM_2.7.xlsx (", nrow(comp_tbl), " rows = ",
          length(fate_patterns), " fate variables x 4 scenarios).")
}, error = function(e) msg("Section 3C (sensitivity) skipped: ", conditionMessage(e)))


## ===========================================================================
## 4.  Summary engine  (NaN/Inf-safe; empty groups -> NA, never #NUM!)
## ===========================================================================
m_mean   <- function(x) { x <- x[is.finite(x)]; if (length(x))      mean(x)          else NA_real_ }
m_sd     <- function(x) { x <- x[is.finite(x)]; if (length(x) > 1L) stats::sd(x)     else NA_real_ }
m_median <- function(x) { x <- x[is.finite(x)]; if (length(x))      stats::median(x) else NA_real_ }
safe_q   <- function(x, p) { x <- x[is.finite(x)]
  if (length(x)) stats::quantile(x, p, names = FALSE, type = 7) else NA_real_ }

stat_set <- list(mean = ~m_mean(.), sd = ~m_sd(.), median = ~m_median(.),
                 p25 = ~safe_q(., 0.25), p75 = ~safe_q(., 0.75), p95 = ~safe_q(., 0.95))

summarise_block <- function(df)
  dplyr::summarise(df, dplyr::across(dplyr::where(is.numeric), stat_set,
                                     .names = "{.col}_{.fn}"), .groups = "drop")

sanitize <- function(df)
  dplyr::mutate(df, dplyr::across(dplyr::where(is.numeric),
                                  ~ ifelse(is.finite(.), ., NA_real_)))

as_overall <- function(block) {
  v <- unlist(block[1, ], use.names = TRUE)
  tibble::tibble(statistic = names(v), value = as.numeric(v))
}
as_by_class <- function(grouped_block)
  grouped_block %>%
    tidyr::pivot_longer(-PCs_class, names_to = "statistic", values_to = "v") %>%
    tidyr::pivot_wider(names_from = PCs_class, values_from = "v") %>%
    dplyr::rename(PCs_class = statistic)

## ---------------------------------------------------------------------------
## AGG REDUCTION RULE  --  manuscript Section 2.5:
##   
## ---------------------------------------------------------------------------
AGG_REDUCE  <- "median"                       # "median" (manuscript 2.5) | "mean"
agg_reducer <- switch(AGG_REDUCE, median = m_median, mean = m_mean,
                      stop("AGG_REDUCE must be 'median' or 'mean'.", call. = FALSE))
msg("AGG (compound-weighted) reduction: each compound -> its ", AGG_REDUCE,
        "  [manuscript 2.5 specifies the median]")

build_outputs <- function(df) {
  har_overall  <- summarise_block(df)
  har_by_class <- df %>% dplyr::group_by(PCs_class) %>% summarise_block()
  har_by_comp  <- df %>% dplyr::group_by(Compound)  %>% summarise_block()
  agg_tbl      <- df %>% dplyr::group_by(Compound, PCs_class) %>%
                  dplyr::summarise(dplyr::across(dplyr::where(is.numeric), ~agg_reducer(.)),
                                   .groups = "drop")
  agg_overall  <- summarise_block(agg_tbl %>% dplyr::select(-Compound, -PCs_class))
  agg_by_class <- agg_tbl %>% dplyr::group_by(PCs_class) %>%
                  dplyr::summarise(dplyr::across(dplyr::where(is.numeric), stat_set,
                                                 .names = "{.col}_{.fn}"), .groups = "drop")
  list(Harmonized_Overall = as_overall(har_overall),
       Harmonized_By_Class = as_by_class(har_by_class),
       Harmonized_By_Compound = har_by_comp,
       AGG_Overall = as_overall(agg_overall),
       AGG_By_Class = as_by_class(agg_by_class),
       AGG_By_Compound = agg_tbl)
}

write_book <- function(sheets, path) {
  wb <- openxlsx::createWorkbook()
  for (nm in names(sheets)) {
    openxlsx::addWorksheet(wb, nm)
    openxlsx::writeData(wb, nm, sanitize(sheets[[nm]]), keepNA = TRUE, na.string = NA_TOKEN)
  }
  openxlsx::saveWorkbook(wb, file.path(out_dir, path), overwrite = TRUE)
}

STAGE("6   descriptive stats -> SM_3 / SM_4")
write_book(build_outputs(prop), "SM_3_properties.xlsx")
## Concentration statistics come from the censored estimator (Section 6C), not from the
## point table, whose concentration columns exclude the censored-floor ranges and are
## therefore biased. Drop the four concentration columns here so SM_4_fate holds only the
## valid half-life descriptive statistics; Section 6C then adds the censored concentration
## sheets (Conc_*) to the same workbook.
fate_hl <- dplyr::select(fate, -dplyr::any_of(c("Concentration_SW","Concentration_SED",
                                                "Concentration_HZ","Concentration_GW")))
write_book(build_outputs(fate_hl), "SM_4_fate.xlsx")
wrote("SM_3_properties.xlsx")
wrote("SM_4_fate.xlsx")

## ===========================================================================
## SECTION 6C -- CENSORED-DATA ESTIMATION FOR CONCENTRATIONS
## Replaces DL/2 substitution for the four concentration variables (C_SW, C_HZ,
## C_SED, C_GW). Half-lives and properties are untouched.
##   Primary  : Kaplan-Meier (left-censored NPMLE via survival::survfit). Censored-
##              floor ranges (n.d.-y, <x-y), which carry a measured upper bound and
##              are compiled multi-observation summaries, are treated as detections
##              at their midpoint -- the stable, defensible choice for this data.
##   Sensitivity: Turnbull (same survfit, but the ranges entered as wide intervals
##              (m, y)). Reported as a conservative lower bound; on this dataset it
##              under-estimates the median because the compiled ranges become very
##              wide intervals when the lowest detected value m is tiny.
##   Cross-checks: robust ROS (NADA::ros, guarded) and DL/2 substitution (screening).
##   Mean via EnvStats::enparCensored (guarded).
##   No-limit non-detects take the lowest detected value m of their OWN variable as
##   the censoring limit (per matrix; no cross-matrix constant). AGG retains all-ND
##   compounds as left-censored (never deleted). BY-CLASS estimates are computed only
##   where a class has >=3 detections and n>=5, else reported as detection frequency.
##   Output the Conc_* sheets added to SM_4_fate.xlsx (Estimates, By_class, By_compound, Detection,
##   Notes): the KM rows are the concentration statistics OF RECORD, superseding the
##   concentration rows of SM_4_fate.xlsx (its half-life rows stand).
## Refs: Helsel (2006; 2012); Turnbull (1976); Shoari & Dube (2018).
## ===========================================================================

STAGE("6C  censored-data estimation for concentrations (KM primary; Turnbull/ROS/DL2 cross-checks; by-class)")

MIN_N_EST   <- 5    # minimum records for an estimate
MIN_DET_EST <- 3    # minimum detections for an estimate

tryCatch({

  have <- function(p) requireNamespace(p, quietly = TRUE)
  if (!have("survival")) stop("Section 6C needs the 'survival' package.", call.=FALSE)
  has_nada <- have("NADA"); has_env <- have("EnvStats")

  conc_vars  <- c("Concentration_SW","Concentration_SED","Concentration_HZ","Concentration_GW")
  conc_label <- c(Concentration_SW="C_SW", Concentration_SED="C_SED",
                  Concentration_HZ="C_HZ", Concentration_GW="C_GW")
  conc_unit  <- c(Concentration_SW="ug/L", Concentration_SED="ug/kg",
                  Concentration_HZ="ug/L", Concentration_GW="ug/L")

  raw_fate <- apply_corrections(read_sm2(SH_RAW_FATE), SH_RAW_FATE)
  rnms     <- norm_names(names(raw_fate))
  cmp_col  <- find_col(rnms, "^Compound", "Compound")
  cls_col  <- which(grepl("^(PCs|Clinical Application|Therapeutic|Class)", rnms, ignore.case=TRUE))[1]
  ff <- function(v){ v <- as.character(v)
    for (i in seq_along(v)) if (i>1 && (is.na(v[i]) || !nzchar(trimws(v[i])))) v[i] <- v[i-1]; v }
  Compound <- ff(raw_fate[[cmp_col]])
  PCs      <- if (is.na(cls_col)) rep(NA_character_, nrow(raw_fate)) else ff(raw_fate[[cls_col]])
  q_probs  <- c(.25, .5, .75, .95)

  ## ---- parse one variable into a per-record frame with BOTH range codings ----
  ## km_lo/km_hi  : ranges as EXACT detections at the midpoint (primary)
  ## tb_lo/tb_hi  : ranges as INTERVALS (m, y) (Turnbull sensitivity)
  ## coding for survival::Surv(type="interval2"): exact (v,v); left-cens (NA,x);
  ##   right-cens (x,NA); interval (a,b).
  build_frame <- function(var) {
    j    <- find_col(rnms, fate_patterns[[var]], var)
    P    <- lapply(raw_fate[[j]], parse_conc)
    kind <- vapply(P,`[[`,character(1),"kind"); lo <- vapply(P,`[[`,numeric(1),"lo"); hi <- vapply(P,`[[`,numeric(1),"hi")
    measured <- c(lo[kind %in% c("NUM","GT","APPROX")], lo[kind=="RANGE"], hi[kind=="RANGE"],
                  hi[kind %in% c("MIX_nolim","MIX_lim")])
    measured <- measured[is.finite(measured) & measured>0]; m <- if(length(measured)) min(measured) else NA_real_
    n <- length(P)
    km1<-rep(NA_real_,n); km2<-rep(NA_real_,n); tb1<-rep(NA_real_,n); tb2<-rep(NA_real_,n)
    obs<-rep(NA_real_,n); cen<-rep(NA,n); det<-rep(FALSE,n); keep<-rep(FALSE,n)
    for (i in seq_len(n)) {
      k<-kind[i]
      if (k %in% c("NUM","APPROX")) { v<-lo[i]; km1[i]<-v;km2[i]<-v; tb1[i]<-v;tb2[i]<-v; obs[i]<-v;cen[i]<-FALSE;det[i]<-TRUE;keep[i]<-TRUE }
      else if (k=="GT")   { v<-lo[i]; km1[i]<-v;km2[i]<-v; tb1[i]<-v;tb2[i]<-NA_real_; obs[i]<-v;cen[i]<-FALSE;det[i]<-TRUE;keep[i]<-TRUE }  # >x
      else if (k=="RANGE"){ mid<-(lo[i]+hi[i])/2; km1[i]<-mid;km2[i]<-mid; tb1[i]<-lo[i];tb2[i]<-hi[i]; obs[i]<-mid;cen[i]<-FALSE;det[i]<-TRUE;keep[i]<-TRUE }
      else if (k=="MIX_lim"){ mid<-(lo[i]+hi[i])/2; km1[i]<-mid;km2[i]<-mid; tb1[i]<-m;tb2[i]<-hi[i]; obs[i]<-mid;cen[i]<-FALSE;det[i]<-TRUE;keep[i]<-TRUE }   # <x-y
      else if (k=="MIX_nolim"){ mid<-(m+hi[i])/2; km1[i]<-mid;km2[i]<-mid; tb1[i]<-m;tb2[i]<-hi[i]; obs[i]<-mid;cen[i]<-FALSE;det[i]<-TRUE;keep[i]<-TRUE }    # n.d.-y
      else if (k=="LT_lim"){ km1[i]<-NA_real_;km2[i]<-lo[i]; tb1[i]<-NA_real_;tb2[i]<-lo[i]; obs[i]<-lo[i];cen[i]<-TRUE;keep[i]<-TRUE }  # <x
      else if (k=="NOLIM") { km1[i]<-NA_real_;km2[i]<-m;     tb1[i]<-NA_real_;tb2[i]<-m;     obs[i]<-m;    cen[i]<-TRUE;keep[i]<-TRUE }  # n.d.
    }
    data.frame(Compound=Compound, PCs=PCs, km1=km1,km2=km2, tb1=tb1,tb2=tb2,
               obs=obs, cen=cen, det=det, m=m, stringsAsFactors=FALSE)[keep,,drop=FALSE] -> d
    attr(d,"m")<-m; d
  }
  F <- lapply(conc_vars, build_frame); names(F) <- conc_vars

  ## ---- estimator helpers (guarded) ----------------------------------------
  surv_q <- function(t1,t2){
    ok <- !(is.na(t1)&is.na(t2)); t1<-t1[ok]; t2<-t2[ok]
    S <- survival::Surv(time=t1, time2=t2, type="interval2")
    f <- survival::survfit(S ~ 1)
    as.numeric(stats::quantile(f, probs=q_probs)$quantile)
  }
  ros_q <- function(obs,cen){ r<-NADA::ros(obs,cen); as.numeric(stats::quantile(r, probs=q_probs)) }
  env_mean <- function(obs,cen){ EnvStats::enparCensored(obs,cen,censoring.side="left")$parameters[["mean"]] }
  dl2_q <- function(obs,cen){ v<-ifelse(cen,obs/2,obs); v<-v[is.finite(v)&v>0]
    if(length(v)) as.numeric(stats::quantile(v,q_probs,names=FALSE,type=7)) else rep(NA_real_,4) }
  safe4 <- function(e) tryCatch(e, error=function(x) rep(NA_real_,4))
  safe1 <- function(e) tryCatch(e, error=function(x) NA_real_)

  mk <- function(var,weighting,estimator,q,mn,n,ndet,mval,primary){
    data.frame(Variable=conc_label[[var]],Unit=conc_unit[[var]],Weighting=weighting,Estimator=estimator,
               n=n,n_detected=ndet,detection_pct=if(n) round(100*ndet/n,1) else NA_real_,
               min_detected_limit=signif(mval,4),
               P25=signif(q[1],4),Median=signif(q[2],4),P75=signif(q[3],4),P95=signif(q[4],4),
               Mean=signif(mn,4),primary=primary,stringsAsFactors=FALSE) }

  ## ---- HAR (record-weighted) ----------------------------------------------
  har <- do.call(rbind, lapply(conc_vars, function(v){
    d<-F[[v]]; n<-nrow(d); ndet<-sum(d$det); m<-attr(d,"m")
    out <- mk(v,"HAR","KM", safe4(surv_q(d$km1,d$km2)),
              if(has_env) safe1(env_mean(d$obs,d$cen)) else NA_real_, n,ndet,m, TRUE)
    out <- rbind(out, mk(v,"HAR","Turnbull", safe4(surv_q(d$tb1,d$tb2)), NA_real_, n,ndet,m, FALSE))
    if (has_nada) out <- rbind(out, mk(v,"HAR","ROS", safe4(ros_q(d$obs,d$cen)),
                                       if(has_env) safe1(env_mean(d$obs,d$cen)) else NA_real_, n,ndet,m, FALSE))
    out <- rbind(out, mk(v,"HAR","DL/2", safe4(dl2_q(d$obs,d$cen)),
                         safe1(mean((function(x){x[is.finite(x)&x>0]})(ifelse(d$cen,d$obs/2,d$obs)))), n,ndet,m, FALSE))
    out }))

  ## ---- BY CLASS (record-weighted within therapeutic class) -----------------
  byclass <- do.call(rbind, lapply(conc_vars, function(v){
    d<-F[[v]]
    do.call(rbind, lapply(split(d, d$PCs), function(g){
      n<-nrow(g); ndet<-sum(g$det); m<-attr(d,"m")
      if (ndet < MIN_DET_EST || n < MIN_N_EST)
        return(data.frame(Variable=conc_label[[v]],Unit=conc_unit[[v]],PCs_class=g$PCs[1],
                          n=n,n_detected=ndet,detection_pct=if(n) round(100*ndet/n,1) else NA_real_,
                          Median="insufficient (report detection frequency)",
                          P25=NA_real_,P75=NA_real_,stringsAsFactors=FALSE))
      q<-safe4(surv_q(g$km1,g$km2))
      data.frame(Variable=conc_label[[v]],Unit=conc_unit[[v]],PCs_class=g$PCs[1],
                 n=n,n_detected=ndet,detection_pct=round(100*ndet/n,1),
                 Median=as.character(signif(q[2],4)),P25=signif(q[1],4),P75=signif(q[3],4),
                 stringsAsFactors=FALSE) })) }))

  ## ---- BY COMPOUND (record-weighted within compound) -----------------------
  bycompound <- do.call(rbind, lapply(conc_vars, function(v){
    d<-F[[v]]
    do.call(rbind, lapply(split(d, d$Compound), function(g){
      n<-nrow(g); ndet<-sum(g$det); m<-attr(d,"m")
      if (ndet < MIN_DET_EST || n < MIN_N_EST)
        return(data.frame(Variable=conc_label[[v]],Unit=conc_unit[[v]],Compound=g$Compound[1],
                          n=n,n_detected=ndet,detection_pct=if(n) round(100*ndet/n,1) else NA_real_,
                          Median="insufficient (report detection frequency)",
                          P25=NA_real_,P75=NA_real_,P95=NA_real_,stringsAsFactors=FALSE))
      q<-safe4(surv_q(g$km1,g$km2))
      data.frame(Variable=conc_label[[v]],Unit=conc_unit[[v]],Compound=g$Compound[1],
                 n=n,n_detected=ndet,detection_pct=round(100*ndet/n,1),
                 Median=as.character(signif(q[2],4)),P25=signif(q[1],4),P75=signif(q[3],4),
                 P95=signif(q[4],4),stringsAsFactors=FALSE) })) }))

  ## ---- AGG (compound-weighted; all-ND compounds retained as censored) ------
  agg <- do.call(rbind, lapply(conc_vars, function(v){
    d<-F[[v]]; m<-attr(d,"m")
    rc <- do.call(rbind, lapply(split(d, d$Compound), function(g){
      det<-g$obs[g$det & is.finite(g$obs) & g$obs>0]
      if (length(det)) data.frame(val=stats::median(det), cen=FALSE)
      else data.frame(val=suppressWarnings(min(g$obs[is.finite(g$obs)])), cen=TRUE) }))
    rc<-rc[is.finite(rc$val)&rc$val>0,,drop=FALSE]; n<-nrow(rc); ndet<-sum(!rc$cen)
    km1<-ifelse(rc$cen,NA_real_,rc$val); km2<-rc$val
    out <- mk(v,"AGG","KM", safe4(surv_q(km1,km2)),
              if(has_env) safe1(env_mean(rc$val,rc$cen)) else NA_real_, n,ndet,m, TRUE)
    out <- rbind(out, mk(v,"AGG","DL/2", safe4(dl2_q(rc$val,rc$cen)), NA_real_, n,ndet,m, FALSE))
    out }))

  det <- do.call(rbind, lapply(conc_vars, function(v){
    d<-F[[v]]; data.frame(Variable=conc_label[[v]],Unit=conc_unit[[v]],n=nrow(d),n_detected=sum(d$det),
      detection_pct=round(100*sum(d$det)/nrow(d),1), min_detected_value=signif(attr(d,"m"),4),
      stringsAsFactors=FALSE) }))

  notes <- data.frame(Notes=c(
    "CENSORED-DATA concentration statistics -- the concentration statistics OF RECORD for the four",
    "concentration variables, held in the Conc_* sheets of this workbook (SM_4_fate.xlsx). The point-table",
    "concentration descriptive statistics are NOT produced: they exclude the censored-floor ranges and are biased.",
    "Primary = Kaplan-Meier (left-censored NPMLE, survival::survfit); censored-floor ranges (n.d.-y, <x-y)",
    "  treated as detections at their midpoint. Turnbull (ranges as intervals) is a conservative lower-bound",
    "  sensitivity. ROS (NADA, if available) and DL/2 are further cross-checks. Mean via EnvStats::enparCensored.",
    "No-limit non-detects take the lowest detected value of their own variable as the censoring limit (per matrix).",
    "BY-CLASS and BY-COMPOUND: estimated only where the group has >=3 detections and n>=5; otherwise the group",
    "  is reported by detection frequency. AGG: compound -> median of its detected values; all-non-detect",
    "  compounds RETAINED as left-censored. Half-life descriptive statistics are in the Harmonized_*/AGG_* sheets.",
    "Units: C_SW/C_HZ/C_GW = ug/L ; C_SED = ug/kg."), stringsAsFactors=FALSE)

  ## Write the censored concentration statistics INTO SM_4_fate.xlsx as additional sheets,
  ## so every fate statistic lives in one file and the biased point-table concentration
  ## table is absent. The Conc_* sheets are the concentration record.
  sm4_path <- file.path(out_dir, "SM_4_fate.xlsx")
  wb <- openxlsx::loadWorkbook(sm4_path)
  add_conc_sheet <- function(nm, d){
    if (nm %in% openxlsx::sheets(wb)) openxlsx::removeWorksheet(wb, nm)
    openxlsx::addWorksheet(wb, nm); openxlsx::writeData(wb, nm, d) }
  add_conc_sheet("Conc_Estimates",   rbind(har,agg))
  add_conc_sheet("Conc_By_class",    byclass)
  add_conc_sheet("Conc_By_compound", bycompound)
  add_conc_sheet("Conc_Detection",   det)
  add_conc_sheet("Conc_Notes",       notes)
  openxlsx::saveWorkbook(wb, sm4_path, overwrite=TRUE)
  wrote("SM_4_fate.xlsx (concentration sheets added)")
  assign("conc_censored", rbind(har,agg), envir=.GlobalEnv)
  assign("conc_censored_byclass", byclass, envir=.GlobalEnv)
  assign("conc_censored_bycompound", bycompound, envir=.GlobalEnv)
  msg("Section 6C: KM primary + Turnbull/ROS/DL2 cross-checks + by-class + by-compound ",
      "written into SM_4_fate.xlsx (Conc_* sheets). Point-table concentration rows not produced.")

}, error = function(e) msg("Section 6C (censored estimation) skipped: ", conditionMessage(e)))


## ---- AGG reducer comparison (median vs mean), for the audit workbook -------
agg_compare <- function(df, group) {
  vars <- setdiff(names(df)[vapply(df, is.numeric, logical(1))], character(0))
  do.call(rbind, lapply(vars, function(v) {
    by_cmp <- split(df[[v]], df$Compound)
    med <- vapply(by_cmp, m_median, numeric(1)); med <- med[is.finite(med)]
    men <- vapply(by_cmp, m_mean,   numeric(1)); men <- men[is.finite(men)]
    if (!length(med)) return(NULL)
    s4 <- function(x) c(safe_q(x, .25), m_median(x), safe_q(x, .75), safe_q(x, .95))
    a <- s4(med); b <- s4(men)
    data.frame(group = group, variable = v, n_compounds = length(med),
               median_P25 = signif(a[1], 4), median_median = signif(a[2], 4),
               median_P75 = signif(a[3], 4), median_P95 = signif(a[4], 4),
               mean_P25   = signif(b[1], 4), mean_median   = signif(b[2], 4),
               mean_P75   = signif(b[3], 4), mean_P95      = signif(b[4], 4),
               identical  = isTRUE(all.equal(a, b)),
               stringsAsFactors = FALSE)
  }))
}
agg_check <- rbind(agg_compare(prop, "Properties"), agg_compare(fate, "Fate"))
msg("AGG reducer check: ", sum(!agg_check$identical), " of ", nrow(agg_check),
        " variables give different AGG percentiles under median- vs mean-reduction ",
        ".")


## ===========================================================================
## 5.  Figures  (overall / overview / per-compound)  -- unchanged logic
STAGE("5   figures (overall / overview / per-compound)")

## ===========================================================================
dir_overall  <- file.path(fig_dir, "overall")
dir_overview <- file.path(fig_dir, "overview")
dir_comp     <- file.path(fig_dir, "compounds")
for (d in c(dir_overall, dir_overview, dir_comp))
  dir.create(d, showWarnings = FALSE, recursive = TRUE)

class_levels <- c(
  "Analgesic", "Anti-inflammatory", "Antibiotic Beta-lactams",
  "Antibiotic Fluoroquinolones", "Antibiotic from other class",
  "Antibiotic Macrolides", "Antibiotic Sulfonamides",
  "Antibiotic Tetracyclines", "Antidepressant", "Antiepileptic",
  "Antihypertensive", "Beta-blocker", "Lipid regulator",
  "Stimulant", "Tranquilizers")
class_colors <- setNames(scales::hue_pal()(length(class_levels)), class_levels)

canon_class <- function(x) {
  x <- dplyr::recode(trimws(x),
        "Anti-Inflammatory"            = "Anti-inflammatory",
        "Other classes of Antibiotics" = "Antibiotic from other class")
  factor(x, levels = class_levels)
}

prop_vars <- c("log MW", "pKa", "log Kow", "log Koc", "log Dow", "log S")
## FIGURE AXIS LABELS - notation harmonised across every figure:
##   C_SW / C_SED / C_HZ / C_GW  (not "Concentration ...")
##   t\u00bd                          (not "t1/2")
##   log Kow / log Koc / log Dow  (not KOW / KOC / DOW)
fate_vars <- c("C_SW (\u00b5g/L)",   "C_SED (\u00b5g/Kg)",
               "C_HZ (\u00b5g/L)",   "C_GW (\u00b5g/L)",
               "t\u00bd SW (days)",  "t\u00bd SED (days)",  "t\u00bd HZ (days)",
               "t\u00bd aerobic (days)", "t\u00bd anoxic (days)", "t\u00bd GW (days)")

prop_fig <- prop %>%
  dplyr::transmute(PCs_class = canon_class(PCs_class), Compound,
    `log MW` = log10(Molecular_weight), pKa,
    `log Kow` = log_Kow, `log Koc` = `log _Koc`, `log Dow` = log_Dow, `log S` = log_S) %>%
  tidyr::pivot_longer(-c(PCs_class, Compound), names_to = "Variable", values_to = "Value") %>%
  dplyr::filter(!is.na(Value))

## NOTE: \u00b5 (micro sign) escapes are NOT allowed inside `backtick` column names in R,
## so the canonical columns are selected first and mapped to the display labels
## (fate_vars, where \u00b5 is fine because it sits in a normal string).
fate_fig <- fate %>%
  dplyr::transmute(PCs_class = canon_class(PCs_class), Compound,
    Concentration_SW, Concentration_SED, Concentration_HZ, Concentration_GW,
    half_life_SW, half_life_SED, half_life_HZ,
    half_life_aerobic, half_life_anoxic, half_life_GW) %>%
  tidyr::pivot_longer(-c(PCs_class, Compound), names_to = "Variable", values_to = "Value") %>%
  dplyr::filter(!is.na(Value))
fate_lab <- setNames(fate_vars,
  c("Concentration_SW", "Concentration_SED", "Concentration_HZ", "Concentration_GW",
    "half_life_SW", "half_life_SED", "half_life_HZ",
    "half_life_aerobic", "half_life_anoxic", "half_life_GW"))
fate_fig$Variable <- factor(unname(fate_lab[fate_fig$Variable]), levels = fate_vars)

drop_outliers <- function(df, group_cols)
  df %>% dplyr::group_by(dplyr::across(dplyr::all_of(group_cols))) %>%
    dplyr::filter({ q <- stats::quantile(Value, c(.25, .75), na.rm = TRUE, names = FALSE)
                    f <- 1.5 * (q[2] - q[1]); Value >= q[1] - f & Value <= q[2] + f }) %>%
    dplyr::ungroup()

theme_doc <- function(base = 14, angle = 0)
  ggplot2::theme_bw(base_size = base) +
  ggplot2::theme(
    plot.title = ggplot2::element_text(face = "bold", size = base + 4),
    axis.title = ggplot2::element_text(face = "bold", size = base + 1),
    axis.text  = ggplot2::element_text(face = "bold", colour = "black", size = base - 2),
    axis.text.x = ggplot2::element_text(face = "bold", colour = "black", angle = angle,
                                        hjust = if (angle > 0) 1 else 0.5),
    strip.text = ggplot2::element_text(face = "bold", size = base - 1),
    strip.background = ggplot2::element_rect(fill = "grey90"),
    legend.position = "none", panel.grid.minor = ggplot2::element_blank())

## ---- AXIS NUMBER FORMAT ----------------------------------------------------
## Every numeric axis label in every figure is printed with a FIXED number of
## decimals (default 2), so 1.3 prints as "1.30", 0.009 as "0.01", 2 as "2.00".
##
## FIG_SCI_ZERO_GUARD protects the log-scaled concentration axes, which run down
## to 1e-06.  A non-zero value whose fixed-decimal label would collapse to
## "0.00" is shown in compact scientific form instead, so that several decade
## ticks cannot all carry the same "0.00" label.  The guard therefore fires ONLY
## where the fixed-decimal label would be meaningless; every value that renders
## to a non-zero label (0.009 -> "0.01" included) is left in decimal form.
## Set FIG_SCI_ZERO_GUARD <- FALSE to force fixed decimals everywhere.
FIG_DECIMALS       <- 2
FIG_SCI_ZERO_GUARD <- TRUE

plain_lab <- function(x) vapply(x, function(v) {
  if (is.na(v)) return("")
  lab <- formatC(v, format = "f", digits = FIG_DECIMALS, big.mark = "")
  if (FIG_SCI_ZERO_GUARD && v != 0 && as.numeric(lab) == 0)
    return(format(v, scientific = TRUE, trim = TRUE, digits = 2))
  lab
}, character(1))

## Wide-range LOG axes get compact scientific labels instead of fixed decimals.
## The concentration panel spans ~10 orders of magnitude (1e-06 to 1e+04): at two
## decimals the upper ticks ("100.00", "10000.00") are wide enough to collide,
## while the lower ticks lose all resolution.  Scientific labels are uniform in
## width, so they neither overlap nor round away.  Every LINEAR axis keeps the
## fixed 2-decimal format set by plain_lab().
sci_lab <- function(x) vapply(x, function(v) {
  if (is.na(v)) return("")
  if (v == 0) return("0")
  format(v, scientific = TRUE, trim = TRUE, digits = 1)
}, character(1))

breaks3 <- function(lims) { r <- lims[2] - lims[1]
  if (!is.finite(r) || r <= 0) return(lims)
  unique(signif(lims[1] + r * c(0.2, 0.5, 0.8), 2)) }

decade_breaks <- function(lims) { e1 <- floor(log10(min(lims[lims > 0])))
  e2 <- ceiling(log10(max(lims))); 10^seq(e1, e2, by = 2) }

## Facet panels in the overview figures are roughly a third the width of an
## overall-figure panel, so a fixed 2-decade step crowds a wide facet and leaves
## a narrow one with a single tick.  The step adapts to the span instead.
decade_breaks_facet <- function(lims) {
  pos <- lims[lims > 0]
  if (!length(pos)) return(lims)
  e1 <- floor(log10(min(pos))); e2 <- ceiling(log10(max(pos)))
  by <- if (e2 - e1 > 6) 3 else if (e2 - e1 > 3) 2 else 1
  10^seq(e1, e2, by = by)
}

## ---- OVERALL FIGURES: IDENTICAL PANEL GEOMETRY -----------------------------
## Figures 1-3 (properties, concentrations, half-lives) are read next to one
## another, so their plotting panels are pinned to the same absolute size and the
## same left-hand offset.  Left to ggplot the panel width follows the longest
## y-axis label - "Concentration SED (ug/Kg)" against "log S" - so the three
## x-axes start at different positions and do not line up when the figures are
## stacked.  Pinning the geometry keeps every box the same thickness and every
## axis the same length, so nothing is rescaled and no value is distorted.
OV_PANEL_W <- 4.60   # in - plotting panel width, identical in Figs 1-3
OV_ROW_H   <- 0.62   # in - height per boxplot row, identical box thickness
OV_RIGHT_W <- 0.34   # in - right margin, reserved so the last x-axis tick label
                     #      (centred on the panel edge) is not clipped by the canvas.
OV_AXIS_PAD <- 0.06  # in - slack added to the y-axis label column, whose width is
                     #      measured from the rendered labels rather than hard-coded.

## Panel alignment is achieved by padding the LEFT MARGIN column rather than
## resizing the y-axis label column (ggplot sizes that to its own labels; forcing
## a width pushes the text off the canvas).  Column widths are set absolutely, not
## via convertWidth(), because grobwidth resolution depends on the device's font
## metrics.  The measuring device below matches the one ggsave() draws with (ragg
## when installed, otherwise cairo), since text metrics differ between them.
open_measure_dev <- function(width = 8, height = 6, res = 300) {
  f <- tempfile(fileext = ".png")
  if (requireNamespace("ragg", quietly = TRUE))
    ragg::agg_png(f, width = width, height = height, units = "in", res = res)
  else
    grDevices::png(f, width = width, height = height, units = "in", res = res)
}

fix_overall_geometry <- function(p, n_rows, pad) {
  g   <- ggplot2::ggplotGrob(p)
  pan <- g$layout[g$layout$name == "panel", ]
  g$widths[pan$l]  <- grid::unit(OV_PANEL_W, "in")
  g$heights[pan$t] <- grid::unit(OV_ROW_H * n_rows, "in")
  g$widths[1] <- grid::unit(
    grid::convertWidth(g$widths[1], "in", valueOnly = TRUE) + pad, "in")
  g$widths[length(g$widths)] <- grid::unit(OV_RIGHT_W, "in")
  g
}

ov_w  <- 7.0; ov_h  <- 4.5
ovw_w <- 7.0; ovw_h <- 9.7
cmp_w <- 7.0; cmp_h <- 9.3

build_overall <- function(df, var_order, logx = FALSE) {
  df <- drop_outliers(df, "Variable")
  if (logx) df <- df[df$Value > 0, , drop = FALSE]
  df$Variable <- factor(df$Variable, levels = var_order)
  p <- ggplot2::ggplot(df, ggplot2::aes(x = Value, y = Variable)) +
    ggplot2::geom_boxplot(fill = "white", outlier.shape = NA, linewidth = 0.6) +
    ## No baked-in title and no y-axis title: Figures 1-3 are captioned in the
    ## manuscript, and "Variable" is redundant against the row labels themselves.
    ## Removing both also frees the width that was pushing the last x-axis label
    ## off the canvas.
    ggplot2::labs(title = NULL, x = "Value", y = NULL) + theme_doc(base = 15)
  if (logx) p + ggplot2::scale_x_log10(breaks = decade_breaks, labels = sci_lab)
  else      p + ggplot2::scale_x_continuous(labels = plain_lab)
}

## The three overall figures are built first, then saved with a SHARED y-axis
## label width, so their panels start at the same x-position and are the same
## size.  The shared width is the widest the three actually need, measured from
## the rendered grobs inside a device matching the output resolution.
ov_specs <- list(
  list(p = build_overall(prop_fig, rev(prop_vars)),
       n = length(prop_vars),  file = "1_properties.png"),
  list(p = build_overall(dplyr::filter(fate_fig, Variable %in% fate_vars[1:4]),
                         rev(fate_vars[1:4]), logx = TRUE),
       n = 4L,                 file = "2_concentrations.png"),
  ## Half-lives are plotted on log10, as concentrations already are: the compiled
  ## t1/2 span more than four orders of magnitude (hours in the hyporheic zone to
  ## years in sediment), so on a linear axis every distribution below ~100 d
  ## collapses onto the origin and only the sediment tail is legible.  This
  ## affects the FIGURE ONLY; the statistics in SM_4 remain untransformed.
  list(p = build_overall(dplyr::filter(fate_fig, Variable %in% fate_vars[5:10]),
                         rev(fate_vars[5:10]), logx = TRUE),
       n = 6L,                 file = "3_half_lives.png"))

## All measurement happens inside ONE open drawing-device pass, so the numbers
## used to size the canvas are the numbers ggsave will honour.
open_measure_dev()
ov_natural <- vapply(ov_specs, function(s) {
  g   <- ggplot2::ggplotGrob(s$p)
  pan <- g$layout[g$layout$name == "panel", ]
  sum(grid::convertWidth(g$widths[seq_len(pan$l - 1L)], "in", valueOnly = TRUE))
}, numeric(1))
ov_left_target <- max(ov_natural) + OV_AXIS_PAD
ov_built <- lapply(seq_along(ov_specs), function(i) {
  s <- ov_specs[[i]]
  g <- fix_overall_geometry(s$p, s$n, pad = ov_left_target - ov_natural[i])
  list(g = g,
       ## width is deterministic: every column left of the panel now sums to
       ## ov_left_target, and the only column right of it is OV_RIGHT_W
       w = ov_left_target + OV_PANEL_W + OV_RIGHT_W,
       h = sum(grid::convertHeight(g$heights, "in", valueOnly = TRUE)))
})
grDevices::dev.off()

for (i in seq_along(ov_specs))
  ggplot2::ggsave(file.path(dir_overall, ov_specs[[i]]$file), ov_built[[i]]$g,
                  width = ov_built[[i]]$w, height = ov_built[[i]]$h, dpi = 300)

overview_plot <- function(df, var_order, title, file, logx = FALSE) {
  df <- drop_outliers(df, c("Variable", "PCs_class"))
  ## a log axis cannot show non-positive values; none occur in the fate data once
  ## rules (i) and (vi) have been applied, but the guard keeps the figure honest
  ## if a future compilation introduces one
  if (logx) df <- df[is.finite(df$Value) & df$Value > 0, , drop = FALSE]
  df$Variable <- factor(df$Variable, levels = var_order)
  p <- ggplot2::ggplot(df, ggplot2::aes(x = Value, y = PCs_class, fill = PCs_class)) +
    ggplot2::geom_boxplot(outlier.shape = NA, linewidth = 0.35) +
    ggplot2::facet_wrap(~ Variable, scales = "free_x", ncol = 3) +
    ggplot2::scale_fill_manual(values = class_colors, drop = FALSE) +
    ggplot2::scale_y_discrete(limits = rev(class_levels)) +
    (if (logx) ggplot2::scale_x_log10(breaks = decade_breaks_facet, labels = sci_lab)
     else      ggplot2::scale_x_continuous(breaks = breaks3, labels = plain_lab)) +
    ggplot2::labs(title = title, x = "Value", y = NULL) + theme_doc(base = 11) +
    ggplot2::theme(axis.text.y = ggplot2::element_text(face = "bold", size = 7.5),
                   axis.text.x = ggplot2::element_text(face = "bold", size = 8, angle = 30, hjust = 1),
                   panel.spacing.x = grid::unit(0.7, "lines"))
  ggplot2::ggsave(file.path(dir_overview, file), p, width = ovw_w, height = ovw_h, dpi = 300)
}
overview_plot(prop_fig, prop_vars, "Properties \u2014 comparison across PCs classes", "properties_overview.png")
## Fate overview on log10 for BOTH concentrations and half-lives, so the panels
## of one figure share a scale convention.  Properties stay linear: their values
## are already logarithms or bounded quantities.  Per-compound fate figures and
## the summary figures are unchanged.
overview_plot(fate_fig, fate_vars, "Fate \u2014 comparison across PCs classes",
              "fate_overview.png", logx = TRUE)

compound_class <- prop_fig %>% dplyr::distinct(Compound, PCs_class)
panel <- function(d, var_order, fillcol, panel_title) {
  d <- drop_outliers(d, "Variable"); d$Variable <- factor(d$Variable, levels = var_order)
  ggplot2::ggplot(d, ggplot2::aes(x = Value, y = Variable)) +
    ggplot2::geom_boxplot(fill = fillcol, outlier.shape = NA, linewidth = 0.5) +
    ggplot2::scale_y_discrete(limits = var_order, drop = FALSE) +
    ggplot2::scale_x_continuous(labels = plain_lab) +
    ggplot2::labs(title = panel_title, x = "Value", y = NULL) + theme_doc(base = 12)
}
n_ok <- 0L
for (i in seq_len(nrow(compound_class))) {
  cmp <- compound_class$Compound[i]; cls <- as.character(compound_class$PCs_class[i])
  tryCatch({
    col <- class_colors[[cls]]
    p_prop <- panel(dplyr::filter(prop_fig, Compound == cmp), prop_vars, col, "Properties")
    p_fate <- panel(dplyr::filter(fate_fig, Compound == cmp), fate_vars, col, "Fate")
    combined <- patchwork::wrap_plots(p_prop, p_fate, ncol = 1, heights = c(1, 1.4)) +
      patchwork::plot_annotation(title = paste(cls, "\u2014", cmp),
        theme = ggplot2::theme(plot.title = ggplot2::element_text(face = "bold", size = 17)))
    fn <- paste0(gsub("[^A-Za-z0-9]+", "_", cls), "_", gsub("[^A-Za-z0-9]+", "_", cmp), ".png")
    ggplot2::ggsave(file.path(dir_comp, fn), combined, width = cmp_w, height = cmp_h, dpi = 300)
    n_ok <- n_ok + 1L
  }, error = function(e) msg("  compound figure skipped (", cmp, "): ", conditionMessage(e)))
}
wrote(paste0("figures -> ", fig_dir))
msg("Figures written to: ", fig_dir, "  (overall/=3, overview/=2, compounds/=",
        n_ok, " of ", nrow(compound_class), ")")


## ===========================================================================
## 6.  Compound-level property -> half-life regressions
## ===========================================================================
prop_cmpd <- prop %>% dplyr::group_by(Compound) %>%
  dplyr::summarise(`MW` = m_median(Molecular_weight), `pKa` = m_median(pKa),
                   `log Kow` = m_median(log_Kow), `log Koc` = m_median(`log _Koc`),
                   `log Dow` = m_median(log_Dow), .groups = "drop")
fate_cmpd <- fate %>% dplyr::group_by(Compound) %>%
  dplyr::summarise(`SW t1/2` = m_median(half_life_SW), `SED t1/2` = m_median(half_life_SED),
                   `HZ t1/2` = m_median(half_life_HZ), `Aerobic t1/2` = m_median(half_life_aerobic),
                   `Anoxic t1/2` = m_median(half_life_anoxic), `GW t1/2` = m_median(half_life_GW),
                   .groups = "drop")
only_p <- setdiff(prop_cmpd$Compound, fate_cmpd$Compound)
only_f <- setdiff(fate_cmpd$Compound, prop_cmpd$Compound)
if (length(only_p) || length(only_f))
  msg("NOTE - compound names not matched across sheets:",
          if (length(only_p)) paste0("\n   only in properties: ", paste(only_p, collapse = ", ")) else "",
          if (length(only_f)) paste0("\n   only in fate: ",       paste(only_f, collapse = ", ")) else "")

pred_names <- c("MW", "pKa", "log Kow", "log Koc", "log Dow")
## Response labels are emitted in the form used by the manuscript and by the
## delivered Table_SM_2.8, so that the sheet needs no manual relabelling after
## generation.  The 1/2 glyph is written as a \u escape rather than embedded, so
## the source file stays ASCII and survives any editor or locale.  The rename is
## applied BEFORE the join, so `reg` carries the final labels.
resp_names <- c("t\u00bd SW", "t\u00bd SED", "t\u00bd HZ",
                "t\u00bd aerobic", "t\u00bd anoxic", "t\u00bd GW")
names(fate_cmpd) <- c("Compound", resp_names)

reg <- dplyr::inner_join(prop_cmpd, fate_cmpd, by = "Compound")
sig_code <- function(p) ifelse(is.na(p), "", ifelse(p < .001, "***", ifelse(p < .01, "**",
              ifelse(p < .05, "*", ifelse(p < .1, ".", "-")))))
fmt_p <- function(p) ifelse(is.na(p), NA, ifelse(p < .001, "<0.001", sprintf("%.3f", p)))

## ---------------------------------------------------------------------------
## REGRESSION ENGINE.  Every model is fitted TWICE on exactly the same paired
## compound-level data:
##
##   (a) on the untransformed t1/2 -- the PRIMARY analysis, reported in the main
##       text.  Its slope is directly interpretable (days per predictor unit),
##       which is what makes the sediment t1/2 ~ log KOC result readable: a
##       slope of 0.26 d per log unit with a 95% interval of -138.2 to 138.7 d
##       says something a standardised coefficient does not.
##
##   (b) on log10(t1/2)  -- the SENSITIVITY analysis.  The compiled half-lives
##       span more than four orders of magnitude (0.16 d to 1620.83 d) and are
##       strongly right-skewed.  Fitted untransformed, the residuals are
##       non-normal in every model that reaches significance (Shapiro-Wilk
##       p = 1e-06 to 7e-04) and heteroscedastic in three of them, and the
##       fitted lines predict NEGATIVE half-lives over part of the predictor
##       range (intercepts of -24.09 d for SW ~ MW, -13.75 d for SW ~ pKa,
##       -16.68 d for SED ~ pKa).  Refitting on log10 restores residual
##       normality (Shapiro-Wilk p = 0.06-0.51) and shows how far the reported
##       conclusions depend on the scale.  It is tabulated rather than adopted,
##       so the assumption is auditable instead of assumed either way.
##
## Both are reported with two additional diagnostics:
##
##   * Spearman's rank correlation, computed on the ORIGINAL (untransformed)
##     response.  Because Spearman uses only ranks it is invariant under any
##     monotone transformation, so its value is identical for (a) and (b) by
##     construction.  It is therefore the distribution-free arbiter: where OLS
##     and Spearman agree, the association does not depend on the scale chosen.
##
##   * the Shapiro-Wilk p-value of the model residuals, so that the reader can
##     see directly which of the two fits satisfies the OLS normality
##     assumption rather than having to take it on trust.
##
## A response of t1/2 <= 0 cannot be log-transformed.  The zero policy (Section
## 6 of the description; ZERO_DROP) has already removed the six t1/2 = 0 records
## and no negative half-life exists, so log10 is applied to the full paired
## sample and n is identical in (a) and (b).  A model is fitted only where at
## least three paired observations exist and the predictor has non-zero
## variance, exactly as before.
## ---------------------------------------------------------------------------
fit_one <- function(x, y, rn, pn, logged) {
  ok <- is.finite(x) & is.finite(y)
  if (logged) ok <- ok & y > 0
  if (sum(ok) < 3 || stats::sd(x[ok]) == 0) return(NULL)
  xx <- x[ok]; yraw <- y[ok]
  yy <- if (logged) log10(yraw) else yraw
  m  <- stats::lm(yy ~ xx); sm <- summary(m)
  b0 <- unname(coef(m)[1]); b1 <- unname(coef(m)[2]); pv <- sm$coefficients[2, 4]
  ## Spearman on the ORIGINAL response: rank-based, hence transformation-invariant
  rho <- suppressWarnings(stats::cor(xx, yraw, method = "spearman"))
  sp  <- tryCatch(suppressWarnings(
           stats::cor.test(xx, yraw, method = "spearman", exact = FALSE)$p.value),
           error = function(e) NA_real_)
  res <- stats::residuals(m)
  shp <- if (length(res) >= 3 && length(res) <= 5000 && stats::sd(res) > 0)
           tryCatch(stats::shapiro.test(res)$p.value, error = function(e) NA_real_)
         else NA_real_
  lhs <- if (logged) "log10(t)" else "t"
  data.frame(
    Response = rn, Predictor = pn,
    Formula = sprintf("%s = %.3f %s %.3f x %s", lhs, b0,
                      ifelse(b1 < 0, "-", "+"), abs(b1), pn),
    n_compounds = sum(ok), `R2` = round(sm$r.squared, 3),
    ## Full statistical reporting: slope with standard error and 95%
    ## confidence interval, so a reader can judge precision and power and
    ## not only the coefficient of determination.
    slope        = round(b1, 4),
    slope_SE     = round(sm$coefficients[2, 2], 4),
    slope_CI_low = round(stats::confint(m)[2, 1], 4),
    slope_CI_high= round(stats::confint(m)[2, 2], 4),
    intercept    = round(b0, 3),
    `p_value` = fmt_p(pv), Significance = sig_code(pv),
    ## distribution-free control and the residual-normality diagnostic
    Spearman_rho = round(rho, 3), Spearman_p = fmt_p(sp),
    Spearman_Significance = sig_code(sp),
    residual_normality_p = fmt_p(shp),
    ## exact p retained as a number so that the Bonferroni guard and the
    ## comparison labels are computed from the value, not from its rendering
    p_value_exact = signif(pv, 4),
    check.names = FALSE, stringsAsFactors = FALSE)
}

## the 30 univariate models, fitted on ANY table of compound-level responses
fit_models <- function(dat, logged = TRUE) {
  rows <- list()
  for (rn in resp_names) {
    if (!rn %in% names(dat)) next
    y <- dat[[rn]]
    for (pn in pred_names) {
      r <- fit_one(dat[[pn]], y, rn, pn, logged)
      if (!is.null(r)) rows[[length(rows) + 1]] <- r
    }
  }
  do.call(rbind, rows)
}

## POOLED responses (measured + EPI Suite estimates)
STAGE("8   regressions -> Table_SM_2.8")
tab8_raw <- fit_models(reg, logged = FALSE)   # PRIMARY      untransformed t1/2
tab8_log <- fit_models(reg, logged = TRUE)    # SENSITIVITY  log10(t1/2)
## Table_SM_2.8 is written in Section 7b, together with the MEASURED-ONLY models,
## which need the EPI Suite provenance split computed in Section 7.


## ===========================================================================
## 7.  Persistence (P/vP) screening -- EPI Suite (provenance) sensitivity -> Table_SM_2.9
## ===========================================================================
## WHAT: re-runs the REACH Annex XIII / EMA (2024) persistence flagging at the
##       COMPOUND level twice -- once on all harmonised t1/2, once after removing
##       every EPI Suite estimate -- so the share of P/vP flags that rests on
##       model-derived rather than measured persistence is made explicit.
## EPI Suite values are identified by PROVENANCE, not by magnitude: any t1/2 whose
##       source Reference is "U.S. EPA (2012)" in the harmonised fate sheet
##       (Table_SM_2.6) is treated as model-derived. EPI Suite contributes one
##       surface-water and one sediment estimate per compound (48 + 48 = 96 values).
## RULES (manuscript 2.7): P  if compound-median t1/2 > 40 d (SW) or > 120 d (SED);
##                         vP if compound-median t1/2 > 60 d (SW) or > 180 d (SED).
##       Flags are read per compartment from the COMPOUND MEDIAN; upper
##       percentiles / single literature extremes never trigger a category.
## NOTE: comparisons are strictly ">", so SW = 60 d gives P (not vP); change to ">="
##       to let boundary values trigger the higher class.

EPISUITE_REF_PATTERN <- "^\\s*U\\.S\\. EPA"   # References tag of EPI Suite estimates

## re-read the harmonised fate sheet WITH its References column
## (read_harmonized keeps only the analysis variables and drops References)
hf  <- apply_corrections(read_sm2(SH_HARM_FATE), SH_HARM_FATE)
hn  <- norm_names(names(hf))
ci_compound   <- find_col(hn, "^Compound",   "Compound")
ci_t12_SW     <- find_col(hn, "^t.* SW",      "t1/2 SW")
ci_t12_SED    <- find_col(hn, "^t.* SED",     "t1/2 SED")
ci_t12_GW     <- find_col(hn, "^t.* GW",      "t1/2 GW")
ci_t12_HZ     <- find_col(hn, "^t.* HZ",      "t1/2 HZ")
ci_t12_AER    <- find_col(hn, "^t.* aerobic", "t1/2 aerobic")
ci_t12_ANOX   <- find_col(hn, "^t.* anoxic",  "t1/2 anoxic")
ci_references <- find_col(hn, "^References", "References")

zdrop <- function(v) { v[!is.na(v) & v == 0] <- NA_real_; v }   # ZERO_DROP = half_life
mm    <- function(v) { v <- v[is.finite(v)]; if (length(v)) stats::median(v) else NA_real_ }

scr_raw <- data.frame(
  Compound  = trimws(hf[[ci_compound]]),
  Reference = trimws(hf[[ci_references]]),
  is_epi    = grepl(EPISUITE_REF_PATTERN, hf[[ci_references]]),
  t_SW   = zdrop(clean_num(hf[[ci_t12_SW]])),
  t_SED  = zdrop(clean_num(hf[[ci_t12_SED]])),
  t_GW   = zdrop(clean_num(hf[[ci_t12_GW]])),
  t_HZ   = zdrop(clean_num(hf[[ci_t12_HZ]])),
  t_AER  = zdrop(clean_num(hf[[ci_t12_AER]])),
  t_ANOX = zdrop(clean_num(hf[[ci_t12_ANOX]])),
  stringsAsFactors = FALSE)

## STUDY-LEVEL median: reduce each STUDY to its own median first, then take the
## median across STUDIES.  This removes within-study pseudo-replication (Hurlbert,
## 1984): several records reported by one study no longer outweigh a single record
## reported by another.  Applied to MEASURED values only (EPI Suite excluded).
study_med <- function(v, ref, epi) {
  keep <- is.finite(v) & !epi
  if (!any(keep)) return(NA_real_)
  sm <- tapply(v[keep], ref[keep], stats::median)
  stats::median(as.numeric(sm))
}
scr_raw <- scr_raw[!is.na(scr_raw$Compound) & scr_raw$Compound != "", ]

## per-compound medians: ALL records vs MEASURED-ONLY (EPI Suite removed)
scr <- scr_raw %>%
  dplyr::group_by(Compound) %>%
  dplyr::summarise(
    n_SW         = sum(is.finite(t_SW)),
    n_epi_SW     = sum(is.finite(t_SW) & is_epi),
    t12_SW       = mm(t_SW),
    t12_SW_meas  = mm(t_SW[!is_epi]),
    n_SED        = sum(is.finite(t_SED)),
    n_epi_SED    = sum(is.finite(t_SED) & is_epi),
    t12_SED      = mm(t_SED),
    t12_SED_meas = mm(t_SED[!is_epi]),
    t12_GW       = mm(t_GW),
    ## measured-only medians for the remaining responses (no EPI Suite records
    ## exist for HZ / aerobic / anoxic / GW, so these equal the pooled medians;
    ## they are recomputed here so the measured-only regression set is complete)
    t12_HZ        = mm(t_HZ),   t12_HZ_meas   = mm(t_HZ[!is_epi]),
    t12_AER       = mm(t_AER),  t12_AER_meas  = mm(t_AER[!is_epi]),
    t12_ANOX      = mm(t_ANOX), t12_ANOX_meas = mm(t_ANOX[!is_epi]),
    t12_GW_meas   = mm(t_GW[!is_epi]),
    ## measured evidence base (records and INDEPENDENT STUDIES behind each flag)
    n_meas_SW     = sum(is.finite(t_SW)  & !is_epi),
    n_meas_SED    = sum(is.finite(t_SED) & !is_epi),
    n_stud_SW     = dplyr::n_distinct(Reference[is.finite(t_SW)  & !is_epi]),
    n_stud_SED    = dplyr::n_distinct(Reference[is.finite(t_SED) & !is_epi]),
    ## study-level (pseudo-replication-corrected) medians
    t12_SW_study  = study_med(t_SW,  Reference, is_epi),
    t12_SED_study = study_med(t_SED, Reference, is_epi),
    .groups = "drop")

## attach therapeutic class from the harmonised properties/fate already in memory
cls <- dplyr::distinct(fate[, c("Compound", "PCs_class")])
scr <- dplyr::left_join(scr, cls, by = "Compound")

## flagging helpers (vectorised over the compound-median columns)
flag_sw  <- function(m) ifelse(is.na(m), "\u2014", ifelse(m > 60,  "vP", ifelse(m > 40,  "P", "\u2014")))
flag_sed <- function(m) ifelse(is.na(m), "\u2014", ifelse(m > 180, "vP", ifelse(m > 120, "P", "\u2014")))
combine  <- function(fsw, fsed)
  mapply(function(a, b) {
    p <- c(if (a != "\u2014") paste0(a, " (SW)"), if (b != "\u2014") paste0(b, " (SED)"))
    if (length(p)) paste(p, collapse = "; ") else "\u2014"
  }, fsw, fsed, USE.NAMES = FALSE)
sev <- function(f) ifelse(f == "vP", 2L, ifelse(f == "P", 1L, 0L))

screen <- scr %>%
  dplyr::mutate(
    P_SW_all  = flag_sw(t12_SW),        P_SED_all  = flag_sed(t12_SED),
    P_SW_meas = flag_sw(t12_SW_meas),   P_SED_meas = flag_sed(t12_SED_meas),
    P_SW_stud = flag_sw(t12_SW_study),  P_SED_stud = flag_sed(t12_SED_study),
    flag_all  = combine(P_SW_all,  P_SED_all),
    flag_meas = combine(P_SW_meas, P_SED_meas),
    flag_stud = combine(P_SW_stud, P_SED_stud),
    sev_all   = pmax(sev(P_SW_all),  sev(P_SED_all)),
    sev_meas  = pmax(sev(P_SW_meas), sev(P_SED_meas)),
    sev_stud  = pmax(sev(P_SW_stud), sev(P_SED_stud)),
    estimate_dependent = ifelse(flag_all == flag_meas, "no", "yes"),
    change = dplyr::case_when(
      sev_all >  sev_meas    ~ "weakened (estimate-driven)",
      sev_all <  sev_meas    ~ "strengthened",
      flag_all != flag_meas  ~ "compartment shift",
      TRUE                   ~ "unchanged")) %>%
  dplyr::arrange(dplyr::desc(estimate_dependent == "yes"),
                 dplyr::desc(sev_all), Compound)

## tidy, rounded screening sheet
rnd2 <- function(x) ifelse(is.finite(x), round(x, 2), NA_real_)
tab9 <- screen %>%
  dplyr::transmute(
    Compound, PCs_class,
    n_SW, `n EPI SW` = n_epi_SW,
    `t1/2 SW (all)` = rnd2(t12_SW), `t1/2 SW (measured only)` = rnd2(t12_SW_meas),
    n_SED, `n EPI SED` = n_epi_SED,
    `t1/2 SED (all)` = rnd2(t12_SED), `t1/2 SED (measured only)` = rnd2(t12_SED_meas),
    `t1/2 GW` = rnd2(t12_GW),
    `flag (all t1/2)` = flag_all, `flag (EPI Suite removed)` = flag_meas,
    estimate_dependent, change,
    `n measured SW` = n_meas_SW, `n studies SW` = n_stud_SW,
    `n measured SED` = n_meas_SED, `n studies SED` = n_stud_SED)

## one-line summary sheet
summ9 <- data.frame(
  metric = c("Compounds with any P/vP flag - all t1/2",
             "Compounds with any P/vP flag - EPI Suite (U.S. EPA 2012) removed",
             "Compounds whose flag weakened when EPI Suite removed",
             "Compounds losing ALL P/vP flags when EPI Suite removed",
             "EPI Suite t1/2 removed - surface water",
             "EPI Suite t1/2 removed - sediment",
             "EPI Suite t1/2 removed - total"),
  value  = c(sum(screen$sev_all  > 0),
             sum(screen$sev_meas > 0),
             sum(screen$sev_all > screen$sev_meas),
             sum(screen$sev_all > 0 & screen$sev_meas == 0),
             sum(screen$n_epi_SW),
             sum(screen$n_epi_SED),
             sum(screen$n_epi_SW) + sum(screen$n_epi_SED)),
  check.names = FALSE, stringsAsFactors = FALSE)

## ---------------------------------------------------------------------------
## 7a.  STUDY-LEVEL AGGREGATION SENSITIVITY  (pseudo-replication)
## ---------------------------------------------------------------------------
## The compound median used above is a median across RECORDS: a study reporting
## several half-lives for one compound therefore weighs more than a study
## reporting one.  Here the same flags are recomputed from a median across
## STUDIES (each study first reduced to its own median).  Flags that survive both
## aggregations are robust; flags that appear only under record-level
## aggregation are decided by within-study replication and must be reported as
## such (manuscript Section 4.5).
tab9b <- screen %>%
  dplyr::filter(n_meas_SW > 0 | n_meas_SED > 0) %>%
  dplyr::transmute(
    Compound, PCs_class,
    `n measured SW` = n_meas_SW, `n studies SW` = n_stud_SW,
    `t1/2 SW (measured, record median)` = rnd2(t12_SW_meas),
    `t1/2 SW (measured, study median)`  = rnd2(t12_SW_study),
    `n measured SED` = n_meas_SED, `n studies SED` = n_stud_SED,
    `t1/2 SED (measured, record median)` = rnd2(t12_SED_meas),
    `t1/2 SED (measured, study median)`  = rnd2(t12_SED_study),
    `flag (measured, record-level)` = flag_meas,
    `flag (measured, study-level)`  = flag_stud,
    `flag robust to aggregation` = ifelse(flag_meas == flag_stud, "yes", "no"),
    ## a flag rests on single-study evidence when the compartment that CARRIES it
    ## is supported by only one independent study
    `flag on single-study evidence` = ifelse(
      (P_SW_meas  != "\u2014" & n_stud_SW  <= 1) |
      (P_SED_meas != "\u2014" & n_stud_SED <= 1), "yes", "no")) %>%
  dplyr::arrange(`flag (measured, record-level)` == "\u2014", Compound)

n_meas_flag  <- sum(screen$sev_meas > 0)
n_stud_flag  <- sum(screen$sev_stud > 0)
n_flip       <- sum(screen$sev_meas > 0 & screen$flag_meas != screen$flag_stud)
n_single     <- sum(screen$sev_meas > 0 &
                    ((screen$P_SW_meas  != "\u2014" & screen$n_stud_SW  <= 1) |
                     (screen$P_SED_meas != "\u2014" & screen$n_stud_SED <= 1)))

summ9 <- rbind(summ9, data.frame(
  metric = c("Compounds with any P/vP flag - measured, STUDY-level median",
             "Compounds whose measured flag changes under study-level aggregation",
             "Measured flags resting on a SINGLE independent study",
             "Measured flags supported by >1 independent study"),
  value  = c(n_stud_flag, n_flip, n_single, n_meas_flag - n_single),
  check.names = FALSE, stringsAsFactors = FALSE))

openxlsx::write.xlsx(
  list(screening = tab9, summary = summ9, study_level_sensitivity = tab9b),
  file.path(out_dir, "Table_SM_2.9.xlsx"),
  overwrite = TRUE, keepNA = TRUE, na.string = NA_TOKEN)
wrote("Table_SM_2.9.xlsx")

## ---------------------------------------------------------------------------
## 7b.  MEASURED-ONLY REGRESSIONS  -> Table_SM_2.8
## ---------------------------------------------------------------------------
## WHY: in the pooled responses of Section 6 the compound median IS the EPI Suite
##      estimate for every compound with no measured value (15 of 48 in SW, 27 of
##      48 in SED).  EPI Suite half-lives are structure-derived (BIOWIN) and take
##      a small number of discrete values, so a property -> t1/2 regression run on
##      the pooled responses partly regresses a structural descriptor on a model
##      output built from molecular structure.  The models are therefore refitted
##      on MEASURED-ONLY compound medians (EPI Suite removed) and the two sets are
##      reported side by side.  Any association present in the pooled models but
##      absent in the measured-only models is estimate-driven and must not be
##      interpreted mechanistically.
resp_meas <- screen %>% dplyr::transmute(
  Compound,
  t12_SW_meas, t12_SED_meas, t12_HZ_meas,
  t12_AER_meas, t12_ANOX_meas, t12_GW_meas)
names(resp_meas) <- c("Compound", resp_names)   # same labels as the pooled fits

reg_meas   <- dplyr::inner_join(prop_cmpd, resp_meas, by = "Compound")
tab8m_raw  <- fit_models(reg_meas, logged = FALSE)   # PRIMARY
tab8m_log  <- fit_models(reg_meas, logged = TRUE)    # SENSITIVITY

## --- pooled vs measured-only comparison, on the PRIMARY (untransformed) fits
## The conclusion label is computed from the untransformed p-values, so this
## sheet reproduces the published comparison exactly; the log10 refit is carried
## alongside as extra columns rather than replacing it.
BONF_ALPHA <- 0.05 / 30      # thirty exploratory models inspected as one family

cmp78 <- merge(
  tab8_raw [, c("Response", "Predictor", "n_compounds", "R2", "p_value",
                "p_value_exact", "Spearman_rho", "Spearman_p")],
  tab8m_raw[, c("Response", "Predictor", "n_compounds", "R2", "p_value",
                "Spearman_rho", "Spearman_p")],
  by = c("Response", "Predictor"), all = TRUE,
  suffixes = c("_pooled", "_measured_only"))
cmp78 <- cmp78[order(match(cmp78$Response, resp_names),
                     match(cmp78$Predictor, pred_names)), ]

sig05 <- function(v) !is.na(v) & v < 0.05
p_exact_m <- tab8m_raw$p_value_exact[
  match(paste(cmp78$Response, cmp78$Predictor),
        paste(tab8m_raw$Response, tab8m_raw$Predictor))]

cmp78$conclusion <- ifelse(
  is.na(p_exact_m), "no measured data",
  ifelse(sig05(cmp78$p_value_exact),
    ifelse(!sig05(p_exact_m),
           "significant ONLY with EPI Suite estimates (estimate-driven)",
           "significant in both"),
    "not significant in either"))

## Bonferroni guard.  The thirty models are exploratory and are inspected as a
## family, so the manuscript reports an interpretive threshold of
## alpha = 0.05 / 30 = 0.00167 alongside the nominal p-values.  No adjustment is
## applied to the tabulated p-values themselves; this column simply states,
## model by model, whether the pooled fit would survive it.
cmp78$survives_Bonferroni_pooled <- ifelse(
  is.na(cmp78$p_value_exact), NA,
  ifelse(cmp78$p_value_exact < BONF_ALPHA, "yes", "no"))

## log10 sensitivity carried alongside, so the comparison sheet answers the
## scale question without the reader having to open another tab
log_key <- match(paste(cmp78$Response, cmp78$Predictor),
                 paste(tab8_log$Response, tab8_log$Predictor))
cmp78$R2_pooled_log10      <- tab8_log$R2[log_key]
cmp78$p_value_pooled_log10 <- tab8_log$p_value[log_key]
logm_key <- match(paste(cmp78$Response, cmp78$Predictor),
                  paste(tab8m_log$Response, tab8m_log$Predictor))
cmp78$R2_measured_only_log10      <- tab8m_log$R2[logm_key]
cmp78$p_value_measured_only_log10 <- tab8m_log$p_value[logm_key]

## original column order first, additions appended, so the sheet is a strict
## superset of the previously delivered comparison
cmp78 <- cmp78[, c("Response", "Predictor",
                   "n_compounds_pooled", "R2_pooled", "p_value_pooled",
                   "n_compounds_measured_only", "R2_measured_only",
                   "p_value_measured_only", "conclusion",
                   "Spearman_rho_pooled", "Spearman_p_pooled",
                   "Spearman_rho_measured_only", "Spearman_p_measured_only",
                   "survives_Bonferroni_pooled",
                   "R2_pooled_log10", "p_value_pooled_log10",
                   "R2_measured_only_log10", "p_value_measured_only_log10")]

## --- transformation check: untransformed vs log10, side by side ------------
## Purpose: make the choice of scale auditable rather than assumed.  For each
## model the sheet reports R2, the p-value and the residual-normality p-value
## under both scales, together with the (scale-invariant) Spearman statistic,
## and labels every model on whether the two scales lead to the same conclusion.
tcheck <- merge(
  tab8_raw[, c("Response", "Predictor", "n_compounds", "R2", "p_value",
               "p_value_exact", "residual_normality_p")],
  tab8_log[, c("Response", "Predictor", "R2", "p_value", "p_value_exact",
               "residual_normality_p", "Spearman_rho", "Spearman_p")],
  by = c("Response", "Predictor"), all = TRUE,
  suffixes = c("_untransformed", "_log10"))
tcheck <- tcheck[order(match(tcheck$Response, resp_names),
                       match(tcheck$Predictor, pred_names)), ]
tcheck$agreement <- ifelse(
  sig05(tcheck$p_value_exact_untransformed) == sig05(tcheck$p_value_exact_log10),
  "same conclusion on both scales",
  ifelse(sig05(tcheck$p_value_exact_log10),
         "significant only after log10 transformation",
         "significant only untransformed (scale-dependent)"))
sp_exact <- suppressWarnings(ifelse(tcheck$Spearman_p %in% "<0.001", 0.0005,
                                    as.numeric(tcheck$Spearman_p)))
tcheck$Spearman_supports_untransformed_result <- ifelse(
  sig05(sp_exact) == sig05(tcheck$p_value_exact_untransformed), "yes", "no")
tcheck$p_value_exact_untransformed <- NULL
tcheck$p_value_exact_log10 <- NULL

readme_8 <- data.frame(Notes = c(
  "Table_SM_2.8 - compound-level property -> half-life regressions.",
  "",
  "PRIMARY ANALYSIS (reported in the main text): sheets pooled_t12 and",
  "  measured_only_t12.  Response = t1/2 in days, untransformed, so the slope is",
  "  interpretable in days per predictor unit.  pooled_t12 uses every harmonised",
  "  half-life; measured_only_t12 refits the same models after removing every",
  "  EPI Suite estimate.  comparison places the two side by side and labels each",
  "  model as significant in both, significant only with EPI Suite estimates",
  "  (estimate-driven), or not significant in either.",
  "",
  "SENSITIVITY: sheets log10_pooled_t12 and log10_measured_only_t12 refit the",
  "  same thirty models on log10(t1/2).  Compiled half-lives span more than four",
  "  orders of magnitude and are strongly right-skewed, so the untransformed",
  "  residuals are non-normal and the fitted lines can predict negative",
  "  half-lives.  The transformed fits are tabulated, not adopted: they show how",
  "  far each reported conclusion depends on the scale.",
  "  Three of the thirty models change conclusion between the two scales -",
  "  sediment t1/2 ~ pKa loses significance, and t1/2 ~ log KOC under aerobic and",
  "  anoxic conditions gain it.  None of the three is supported by the rank",
  "  correlation or survives the Bonferroni guard, and the aerobic and anoxic",
  "  models rest on 13 and 11 compounds respectively.",
  "",
  "Spearman_rho / Spearman_p are computed on the untransformed response.  Being",
  "  rank-based they are invariant under any monotone transformation, so they",
  "  are identical in the untransformed and log10 sheets by construction and act",
  "  as the distribution-free control on the OLS result.",
  "residual_normality_p is the Shapiro-Wilk p-value of the model residuals.",
  "p_value_exact is the unrounded p-value; p_value is its rendered form.",
  "",
  "survives_Bonferroni_pooled: the thirty models are exploratory and are",
  "  inspected as one family, so the manuscript reports an interpretive",
  "  threshold of alpha = 0.05/30 = 0.00167.  No adjustment is applied to the",
  "  tabulated p-values; the column states whether each pooled model would",
  "  survive that threshold.",
  "",
  "transformation_check: untransformed versus log10 for every model, with the",
  "  residual-normality diagnostic on both scales and an explicit agreement label.",
  "",
  "Each observation is one compound (median per parameter); n varies between",
  "  models because parameter coverage is uneven.  GW regressions (n = 3) are",
  "  shown for completeness only and are not interpreted.",
  "The models are exploratory and hypothesis-generating, not predictive."),
  stringsAsFactors = FALSE)

openxlsx::write.xlsx(
  list(README                  = readme_8,
       pooled_t12              = tab8_raw,     # PRIMARY
       measured_only_t12       = tab8m_raw,    # PRIMARY
       comparison              = cmp78,
       log10_pooled_t12        = tab8_log,     # SENSITIVITY
       log10_measured_only_t12 = tab8m_log,    # SENSITIVITY
       transformation_check    = tcheck),
  file.path(out_dir, "Table_SM_2.8.xlsx"),
  overwrite = TRUE, keepNA = TRUE, na.string = NA_TOKEN)
wrote("Table_SM_2.8.xlsx")


## ===========================================================================
## 8.  Data audit (console)
## ===========================================================================
audit <- function(df, vars, label) {
  if (QUIET) return(invisible(NULL))
  cat("\n--- audit:", label, " (ZERO_DROP = {", paste(ZERO_DROP, collapse = ","), "}) ---\n", sep = "")
  for (v in vars) { x <- df[[v]]; fin <- x[is.finite(x)]
    cat(sprintf("  %-20s n=%4d  zeros=%3d  missing=%4d  min=%s  median=%s  max=%s\n",
        v, length(fin), sum(fin == 0), sum(is.na(x)),
        ifelse(length(fin), formatC(min(fin),    format = "g", digits = 4), "-"),
        ifelse(length(fin), formatC(stats::median(fin), format = "g", digits = 4), "-"),
        ifelse(length(fin), formatC(max(fin),    format = "g", digits = 4), "-"))) }
}
if (!QUIET) audit(fate, names(fate_patterns), "fate (Table_SM_2.6)")

msg("Done.")
## ###########################################################################
## SECTION 9 - TWO SYNTHESIS FIGURES FOR THE MAIN TEXT
STAGE("9   synthesis figures")

## ---------------------------------------------------------------------------
## Figure A  Regulatory screening map: persistence (t1/2 SED) x mobility (log KOC)
##           x bioaccumulation (log DOW), with MEASURED vs EPI-Suite-estimate-driven
##           persistence distinguished. One picture = the whole Section 4.5 + Table_SM_2.9.
## Figure B  Compartmental concentration gradient (SW-HZ-SED-GW) + within-compound
##           HZ/SW retention ratio. One picture = Section 3.2.2.2 + Section 4.2.
## ###########################################################################

## auto-install extra dependencies -------------------------------------------
for (pkg in c("ggrepel", "patchwork"))
  if (!pkg %in% rownames(installed.packages()))
    install.packages(pkg, repos = "https://cloud.r-project.org")
suppressWarnings(suppressMessages({
    library(ggplot2); library(dplyr); library(tidyr); library(ggrepel); library(patchwork)
}))

fig_sum <- file.path(fig_dir, "summary")
dir.create(fig_sum, showWarnings = FALSE, recursive = TRUE)
DASH <- "\u2014"   # em dash, the "no flag" token used in Section 7

## ===========================================================================
## FIGURE A  Regulatory screening map: persistence x mobility x bioaccumulation
##           TWO PANELS (SW | SED), each with its own REACH Annex XIII thresholds
##           and its own per-compartment provenance colour coding.
##           ALL COMPOUNDS labeled (JCH requirement: legible at 100% zoom).
##
## Visual design:
##   FILL colour  = provenance (blue = Measured; red = Estimate-driven EPI Suite;
##                  light grey = Not persistent above any threshold)
##   SHAPE        = circle (log DOW < 3) vs filled triangle-up (log DOW >= 3,
##                  EMA Phase II B trigger)
##   POINT SIZE   = proportional to log DOW (larger = more lipophilic)
##   STROKE       = solid 0.5 pt white border for legibility on coloured background
##   LABELS       = ggrepel; ALL compounds; 7 pt regular; segment lines to points
##   LEGEND       = single vertical legend block at right; separate rows for
##                  provenance (fill swatches), shape (B-trigger symbol), and
##                  a discrete size guide (3 reference values)
## ===========================================================================
mob <- prop %>% dplyr::group_by(Compound) %>%
    dplyr::summarise(logKOC = m_median(`log _Koc`),
                      logDOW = m_median(log_Dow), .groups = "drop")

mapdat <- screen %>%
    dplyr::select(Compound, PCs_class,
                   t_SW = t12_SW, t_SW_meas = t12_SW_meas,
                   t_SED = t12_SED, t_SED_meas = t12_SED_meas,
                   P_SW_all, P_SW_meas, P_SED_all, P_SED_meas) %>%
    dplyr::left_join(mob, by = "Compound") %>%
    dplyr::mutate(
        prov_SW  = dplyr::case_when(
            P_SW_all  == DASH ~ "Not persistent",
            P_SW_meas == DASH ~ "Estimate-driven",
            TRUE               ~ "Measured"),
        prov_SED = dplyr::case_when(
            P_SED_all  == DASH ~ "Not persistent",
            P_SED_meas == DASH ~ "Estimate-driven",
            TRUE                ~ "Measured"),
        Bioaccum = ifelse(is.finite(logDOW) & logDOW >= 3,
                          "log Dow \u2265 3 (B trigger)", "log Dow < 3"),
        sizeDOW  = pmax(ifelse(is.finite(logDOW), logDOW, 0), 0))

## Provenance palette: blue/red/grey matching original style
pal_prov <- c("Measured"        = "#2166ac",
              "Estimate-driven" = "#d6604d",
              "Not persistent"  = "#bdbdbd")

## Shapes: filled circle vs filled upward triangle (B-trigger)
## Filled shapes (21, 24) allow independent fill + white stroke
shp_prov <- c("log Dow \u2265 3 (B trigger)" = 24,   # filled triangle-up
              "log Dow < 3"                  = 21)   # filled circle

## ---- Figure A geometry -----------------------------------------------------
FIGA_W        <- 16     # in
FIGA_H        <- 8      # in
FIGA_DPI      <- 600    # raster resolution (PDF companion is vector)
FIGA_LAB_SIZE <- 3.4    # mm - compound ID label size (was 2.6, illegible)

make_panelA <- function(yvar, provvar, p_line, vp_line, ylab_text, panel_title) {
    d   <- mapdat %>% dplyr::filter(is.finite(.data[[yvar]]), is.finite(logKOC))
    ## ALL compounds labeled; segment drawn from label to point
    ggplot(d, aes(x = logKOC, y = .data[[yvar]])) +
        ## threshold lines --------------------------------------------------
        geom_vline(xintercept = 3, linetype = "dashed", colour = "grey55", linewidth = 0.55) +
        geom_vline(xintercept = 2, linetype = "solid",  colour = "grey55", linewidth = 0.55) +
        geom_hline(yintercept = p_line,  linetype = "dashed", colour = "grey55", linewidth = 0.55) +
        geom_hline(yintercept = vp_line, linetype = "solid",  colour = "grey55", linewidth = 0.55) +
        ## threshold annotations (placed inside plot area, near top-left) --
        annotate("text", x = -0.3, y = p_line,
                 label = paste0("P (", p_line, " d)"),
                 hjust = 1, vjust = -0.5, size = 2.8, fontface = "italic", colour = "grey40") +
        annotate("text", x = -0.3, y = vp_line,
                 label = paste0("vP (", vp_line, " d)"),
                 hjust = 1, vjust = -0.5, size = 2.8, fontface = "italic", colour = "grey40") +
        ## points: fill = provenance, shape = B-trigger, size = log DOW ----
        geom_point(aes(fill  = .data[[provvar]],
                       shape = Bioaccum,
                       size  = sizeDOW),
                   colour = "white", stroke = 0.5, alpha = 0.92) +
        ## ALL compound labels via ggrepel ----------------------------------
        ggrepel::geom_text_repel(
            aes(label = Compound),
            ## Enlarged with a white halo so labels stay readable where they
            ## overlap points or gridlines.
            size          = FIGA_LAB_SIZE,
            fontface      = "plain",
            colour        = "grey10",
            bg.colour     = "white",
            bg.r          = 0.16,
            segment.size  = 0.3,
            segment.colour= "grey60",
            box.padding   = 0.45,
            point.padding = 0.30,
            min.segment.length = 0.2,
            max.overlaps  = Inf,
            seed          = 42) +
        ## scales -----------------------------------------------------------
        scale_y_log10(
            labels = plain_lab,                    # fixed 2-decimal labels
            breaks = scales::breaks_log(n = 6)) +
        scale_x_reverse(labels = plain_lab) +      # fixed 2-decimal labels
        scale_fill_manual(
            values = pal_prov,
            name   = "Persistence provenance",
            guide  = guide_legend(order = 1,
                                  override.aes = list(shape = 21, size = 4,
                                                      colour = "white", stroke = 0.5))) +
        scale_shape_manual(
            values = shp_prov,
            name   = "Bioaccumulation trigger",
            guide  = guide_legend(order = 2,
                                  override.aes = list(fill = "grey50", size = 4,
                                                      colour = "white", stroke = 0.5))) +
        scale_size_continuous(
            range  = c(2.5, 8),
            breaks = c(0, 1, 2, 3, 4),
            labels = plain_lab,                    # fixed 2-decimal legend labels
            name   = "log Dow (lipophilicity)",
            guide  = guide_legend(order = 3,
                                  override.aes = list(fill = "grey70", shape = 21,
                                                      colour = "white", stroke = 0.4))) +
        labs(title = panel_title,
             ## NOTE: no Unicode arrow here.  U+2192 is absent from the default
             ## device font and rendered as the literal escape "<U+2192>".  The
             ## wording below is glyph-free and matches the Figure 6 caption
             ## ("x-axis reversed, so mobility increases rightwards").
             ## plain string (not plotmath) so the notation matches every other
             ## figure exactly: "log Koc", not "log K_OC"
             x     = "log Koc  (axis reversed: mobility increases rightwards)",
             y     = ylab_text) +
        theme_bw(base_size = 11) +
        theme(
            plot.title       = element_text(face = "bold", size = 11, hjust = 0),
            axis.title       = element_text(size = 10),
            axis.text        = element_text(size = 9),
            legend.position  = "right",
            legend.title     = element_text(face = "bold", size = 9),
            legend.text      = element_text(size = 8.5),
            legend.key.size  = unit(0.45, "cm"),
            legend.spacing.y = unit(0.15, "cm"),
            panel.grid.minor = element_blank(),
            panel.grid.major = element_line(colour = "grey93", linewidth = 0.35))
}

pA_SW  <- make_panelA("t_SW",  "prov_SW",  40,  60,
                       "t\u00bd SW (days)",
                       "Surface water  \u2014  P > 40 d, vP > 60 d")
pA_SED <- make_panelA("t_SED", "prov_SED", 120, 180,
                       "t\u00bd SED (days)",
                       "Sediment  \u2014  P > 120 d, vP > 180 d")

## Combine panels: shared legend on right; figure caption BELOW panels
pA <- pA_SW + pA_SED +
    patchwork::plot_layout(guides = "collect", widths = c(1, 1)) +
    patchwork::plot_annotation(
        caption = paste0(
            "Figure A. Regulatory screening map: persistence \u00d7 mobility \u00d7 bioaccumulation trigger ",
            "(48 PCs, two-compartment view). Each point is one compound; colour indicates ",
            "whether persistence flags are supported by measured data (blue) or rest solely on ",
            "EPI Suite estimates (red); grey points fall below all persistence thresholds. ",
            "Triangle symbols flag the EMA Phase II bioaccumulation trigger (log Dow \u2265 3); ",
            "point size scales with log Dow. Vertical lines: log Koc = 3 (M, dashed) and 2 (vM, solid). ",
            "Horizontal lines: P (dashed) and vP (solid) thresholds (REACH Annex XIII). ",
            "All 48 compounds labeled (ggrepel; seed = 42).")) &
    theme(
        legend.position = "right",
        plot.caption    = element_text(size = 8, hjust = 0, colour = "grey30",
                                       margin = margin(t = 8)))

## FIGA_DPI = 600 so the raster stays sharp when zoomed to read compound IDs.
## The PDF is vector and is the version to submit.
ggsave(file.path(fig_sum, "FigA_screening_map.png"), pA,
       width = FIGA_W, height = FIGA_H, dpi = FIGA_DPI)
ggsave(file.path(fig_sum, "FigA_screening_map.pdf"), pA,
       width = FIGA_W, height = FIGA_H)
wrote("FigA_screening_map.png / .pdf")
msg("Figure A written to: ",
        fig_sum, "  (FigA_screening_map.*)")
## ===========================================================================

## ###########################################################################
## SECTION 10 - RECONCILIATION OF MANUSCRIPT NARRATIVE NUMBERS
STAGE("10  narrative reconciliation -> Table_SM_2.1_a")

## ---------------------------------------------------------------------------
## Makes reproducible the numbers that previously lived only in the running text,
## written to two appendix workbooks:
##   Table_SM_2.1_a.xlsx (appendix to the reference table SM_2.1):
##     (A) study / reference counts          (reconciles the reference keys of the
##                                            REVISED Table_SM_2.1 against the WoS export)
##     (R) reference integrity + row-by-row key <-> WoS alignment check
##     (B) bibliometrics                      (publication year, WoS category, country
##                                             of the corresponding author) from Table_SM_2.1
##     (P) PRISMA count template              (study total auto-filled; funnel manual)
##   Table_SM_2.9_a.xlsx (appendix to the persistence screening SM_2.9):
##     (C) regulatory P/M/B screening table   (Section 4.5 read off log KOC / log DOW
##                                             + the Section 7 persistence flags)
## Reuses in-memory objects: prop, fate, screen, m_median, m_mean, sm2_path,
##   find_col, norm_names, clean_num, out_dir, NA_TOKEN.
## NOT derivable here (must stay manual; recorded as a template in sheet 'PRISMA'):
##   the upstream search/screening funnel 233 -> 198 -> 120 -> 35 (+52) -> 87,
##   because those counts come from the screening log, not from SM_2.
## ###########################################################################
suppressWarnings(suppressMessages({library(dplyr); library(tidyr); library(stringr)}))

## ---- (A) + (B) read Table_SM_2.1 (raw WoS export embedded to the right) ----
## Table_SM_2.1 holds the review's reference table in the leftmost columns
## (col A = "References all" = the reference key) AND a full raw Web-of-Science
## export further right (standard WoS headers: Publication Year, WoS Categories,
## Reprint Addresses, Addresses, ...).  There is NO pivot-style "Row Labels"
## column; the reference key is column A.  WoS availability is decided on the
## presence of the standard export columns (Publication Year + WoS Categories),
## found ANYWHERE in the header row, not on a "Row Labels" column.
SH_REFS <- "Table_SM_2.1"
refs <- tryCatch(read_sm2(SH_REFS),
  error = function(e) { msg("NOTE: sheet ", SH_REFS, " not found."); NULL })
## Drop anything pasted BELOW the reference list (bibliometric / PRISMA appendix)
## so that R2 (orphan keys) measures the reference list, not the appendix.
refs <- trim_ref_block(refs, key_col = 1L)
rn <- if (!is.null(refs)) norm_names(names(refs)) else character(0)

## helper: first column whose (normalised) header matches; NA if none/ambiguous.
## Uses the FIRST match so duplicated/blank trailing headers never break it.
safe_col <- function(nms, pat) {
  hit <- which(grepl(pat, nms, ignore.case = TRUE))
  if (length(hit) >= 1L) hit[1] else NA_integer_
}
ci_key  <- safe_col(rn, "^References all")   # reference key = column A
if (is.na(ci_key)) ci_key <- safe_col(rn, "^Reference")
if (is.na(ci_key)) ci_key <- 1L              # fallback: first column
ci_year <- safe_col(rn, "^Publication Year")
ci_cat  <- safe_col(rn, "^WoS Categor")
ci_rp   <- safe_col(rn, "^Reprint Address")
ci_addr <- safe_col(rn, "^Addresses$")
if (is.na(ci_addr)) ci_addr <- safe_col(rn, "^Addresses")

## WoS available if the export's core columns are present (year + categories)
wos_available <- !is.null(refs) && !is.na(ci_year) && !is.na(ci_cat)
if (!wos_available) {
  msg("NOTE (Section 10): Table_SM_2.1 has no embedded WoS export ",
          "(columns 'Publication Year' and 'WoS Categories' not both found). ",
          "Bibliometric sheets omitted from Table_SM_2.1_a.xlsx. ",
          "Headers seen: ", paste(head(rn[nzchar(rn)], 10), collapse = " | "))
} else {
  msg("Section 10: WoS export detected in Table_SM_2.1 (Publication Year @ col ",
          ci_year, ", WoS Categories @ col ", ci_cat, ", key @ col ", ci_key, ").")
}

## ---- counts that don't require WoS columns --------------------------------
study_counts <- data.frame(
    metric = c("Distinct compounds (harmonised data)", "Therapeutic classes"),
    value  = c(dplyr::n_distinct(prop$Compound), dplyr::n_distinct(prop$PCs_class)),
    stringsAsFactors = FALSE)

## ---- WoS-dependent blocks (skipped when sheet is not a WoS export) -------
keys <- character(0); Nwos <- NA_integer_
yr_tab <- yr_bins <- cat_tab <- ctry_tab <- cont_tab <- NULL
n_unclassified <- NA_integer_

pctW <- function(k) round(100 * k / Nwos, 1)   # defined here for reuse in README
if (wos_available) {
  keys  <- trimws(as.character(refs[[ci_key]]))
  keys  <- keys[!is.na(keys) & keys != ""]
  year  <- suppressWarnings(as.integer(refs[[ci_year]]))
  is_wos <- is.finite(year)
  Nwos  <- sum(is_wos)
  pctW  <- function(k) round(100 * k / Nwos, 1)
  study_counts <- data.frame(
    metric = c("Reference keys in Table_SM_2.1 (column A)",
               "WoS records (with Publication Year) = included studies",
               "Non-WoS references (databases / regulatory / methods)",
               "Distinct compounds (harmonised data)", "Therapeutic classes"),
    value  = c(length(keys), Nwos, length(keys) - Nwos,
               dplyr::n_distinct(prop$Compound), dplyr::n_distinct(prop$PCs_class)),
    stringsAsFactors = FALSE)
  ## B1 year distribution
  yy     <- year[is_wos]
  yr_tab <- as.data.frame(table(Year = yy), stringsAsFactors = FALSE)
  yr_tab$Year <- as.integer(yr_tab$Year); yr_tab$pct <- pctW(yr_tab$Freq)
  yr_bins <- data.frame(
    window = c("2000-2010","2011-2018","2019-2024","2025-2026","peak 2021-2022"),
    n      = c(sum(yy >= 2000 & yy <= 2010), sum(yy >= 2011 & yy <= 2018),
               sum(yy >= 2019 & yy <= 2024), sum(yy >= 2025 & yy <= 2026),
               sum(yy >= 2021 & yy <= 2022)))
  yr_bins$pct <- pctW(yr_bins$n)
  ## B2 WoS categories
  cat_split <- unlist(lapply(refs[[ci_cat]][is_wos], function(x)
    if (is.na(x)) character(0) else trimws(strsplit(x, ";", fixed = TRUE)[[1]])))
  cat_split <- cat_split[nzchar(cat_split)]
  cat_tab <- as.data.frame(table(WoS_Category = cat_split), stringsAsFactors = FALSE)
  cat_tab <- cat_tab[order(-cat_tab$Freq), ]; cat_tab$pct_of_studies <- pctW(cat_tab$Freq)
}

## ---- (B3) corresponding-author country  -----------------------------------
ctry_pat <- c("Peoples R China","China","USA","U Arab Emirates","Germany","Portugal",
              "Spain","Czech Republic","England","Scotland","Wales","Switzerland","Sweden",
              "France","Finland","Netherlands","Denmark","Croatia","Serbia","Greece","Norway",
              "South Korea","Japan","Canada","Brazil","India","Iran","Kenya","Australia",
              "Ireland","Italy","Belgium","Austria","Poland")
ctry_canon <- c("China","China","USA","UAE","Germany","Portugal","Spain","Czech Republic",
                "England","Scotland","Wales","Switzerland","Sweden","France","Finland",
                "Netherlands","Denmark","Croatia","Serbia","Greece","Norway","South Korea",
                "Japan","Canada","Brazil","India","Iran","Kenya","Australia","Ireland",
                "Italy","Belgium","Austria","Poland")
eur  <- c("Germany","Portugal","Spain","Czech Republic","England","Scotland","Wales",
          "Switzerland","Sweden","France","Finland","Netherlands","Denmark","Croatia",
          "Serbia","Greece","Norway","Ireland","Italy","Belgium","Austria","Poland")
asia <- c("China","India","Japan","South Korea","Iran","UAE")
contin <- function(c) ifelse(c %in% eur, "Europe",
                       ifelse(c %in% asia, "Asia",
                       ifelse(c %in% c("USA","Canada"), "North America",
                       ifelse(c == "Brazil", "South America",
                       ifelse(c == "Kenya", "Africa",
                       ifelse(c == "Australia", "Oceania", NA_character_))))))
last_country <- function(s) {
    if (is.na(s) || s == "") return(NA_character_)
    pos <- vapply(ctry_pat, function(p) { m <- gregexpr(p, s, fixed = TRUE)[[1]]
                  if (m[1] == -1L) -1L else max(m) }, integer(1))
    if (max(pos) < 0) return(NA_character_)
    ctry_canon[which.max(pos)]
}
country_of <- function(rp, addr) {                       # corresponding-author country
    rp <- ifelse(is.na(rp), "", rp); addr <- ifelse(is.na(addr), "", addr)
    if (nzchar(rp)) {
        segs <- strsplit(rp, ";", fixed = TRUE)[[1]]
        ca <- segs[grepl("corresponding author|reprint author", segs, ignore.case = TRUE)]
        if (length(ca)) { c1 <- last_country(ca[1]); if (!is.na(c1)) return(c1) }
        c2 <- last_country(rp); if (!is.na(c2)) return(c2)
    }
    last_country(strsplit(addr, ";", fixed = TRUE)[[1]][1])
}
if (wos_available && !is.na(ci_rp) && !is.na(ci_addr)) {
  rp_v <- refs[[ci_rp]][is_wos]; ad_v <- refs[[ci_addr]][is_wos]
  ctry <- mapply(country_of, rp_v, ad_v, USE.NAMES = FALSE)
  ctry_tab <- as.data.frame(table(Country = ctry), stringsAsFactors = FALSE)
  ctry_tab <- ctry_tab[order(-ctry_tab$Freq), ]
  ctry_tab$pct <- pctW(ctry_tab$Freq); ctry_tab$Continent <- contin(ctry_tab$Country)
  cont_tab <- aggregate(Freq ~ Continent, data = ctry_tab, sum)
  cont_tab$pct <- pctW(cont_tab$Freq); cont_tab <- cont_tab[order(-cont_tab$Freq), ]
  n_unclassified <- sum(is.na(ctry))
}

## ---------------------------------------------------------------------------
## (R) REFERENCE INTEGRITY  --  revised Table_SM_2.1, aligned row-by-row with WoS
## ---------------------------------------------------------------------------
## Three checks, in both directions, so the reference table, the data sheets and
## the WoS export can no longer drift apart silently:
##   R1  every reference CITED in the data sheets (Table_SM_2.2/2.3/2.5/2.6) has
##       a key in Table_SM_2.1  -> otherwise the study is uncatalogued;
##   R2  every key in Table_SM_2.1 is CITED at least once -> otherwise it is an
##       orphan entry in the reference list;
##   R3  ROW-BY-ROW alignment: on each row, the leading surname of the reference
##       key occurs in that row's WoS author list, and the key's year equals that
##       row's Publication Year.  This is the check that the revised table was
##       built to satisfy.
## Reference keys are compared after alias canonicalisation (Section 1b), accent
## folding, whitespace collapse and case folding.
nref <- function(x) tolower(fold_ascii(trimws(sub("\\.$", "", norm_names(x)))))

cited <- unique(unlist(lapply(
  c(SH_RAW_PROP, SH_RAW_FATE, SH_HARM_PROP, SH_HARM_FATE), function(sh) {
    d  <- read_sm2(sh); nn <- norm_names(names(d))
    ci <- which(grepl("^References", nn))[1]
    if (is.na(ci)) return(character(0))
    v <- trimws(as.character(d[[ci]]))
    v[!is.na(v) & nzchar(v)]
  })))

ref_keys   <- if (!is.null(refs)) trimws(as.character(refs[[ci_key]])) else character(0)
ref_keys   <- ref_keys[!is.na(ref_keys) & nzchar(ref_keys)]
k_norm     <- nref(ref_keys)
cited_norm <- nref(cited)

R1_missing <- sort(unique(cited[!cited_norm %in% k_norm]))            # cited, not catalogued
R2_orphan  <- sort(unique(ref_keys[!k_norm %in% cited_norm]))         # catalogued, never cited

## R3 row-by-row key <-> WoS alignment
align <- NULL
if (wos_available) {
  ci_au  <- safe_col(rn, "^Authors")
  ci_auf <- safe_col(rn, "^Author Full Names")
  yr_all <- suppressWarnings(as.integer(refs[[ci_year]]))
  au_all <- if (!is.na(ci_au))  as.character(refs[[ci_au]])  else rep(NA_character_, nrow(refs))
  af_all <- if (!is.na(ci_auf)) as.character(refs[[ci_auf]]) else rep(NA_character_, nrow(refs))
  kk     <- trimws(as.character(refs[[ci_key]]))
  keep   <- !is.na(kk) & nzchar(kk)
  ## Compare on letters only: WoS abbreviates compound surnames inconsistently
  ## between its two author fields (e.g. "Garcia-Valverde et al." appears as
  ## "Valverde, MG" in Authors but as "Garcia Valverde, M." in Author Full Names;
  ## "Nguyen et al." appears as "Giang, CND" / "Chau Nguyen Dang Giang").  Both
  ## fields are therefore searched, with punctuation and spacing removed.
  letters_only <- function(x) gsub("[^a-z]", "", tolower(fold_ascii(x)))
  surname <- letters_only(sub("^([A-Za-z\u00C0-\u024F'-]+).*$", "\\1", kk))
  kyear   <- suppressWarnings(as.integer(sub(".*\\((\\d{4})[a-z]?\\).*", "\\1", kk)))
  aulist  <- letters_only(paste(ifelse(is.na(au_all), "", au_all),
                                ifelse(is.na(af_all), "", af_all)))
  has_wos <- is.finite(yr_all)
  ok_name <- mapply(function(s, a) nzchar(s) && grepl(s, a, fixed = TRUE), surname, aulist)
  ok_year <- is.finite(kyear) & is.finite(yr_all) & kyear == yr_all
  align <- data.frame(
    row_in_sheet = which(keep) + 1L,
    reference_key = kk[keep],
    has_WoS_record = has_wos[keep],
    key_year = kyear[keep], WoS_Publication_Year = yr_all[keep],
    WoS_first_author = ifelse(is.na(au_all[keep]), "",
                              sub(";.*$", "", au_all[keep])),
    check = dplyr::case_when(
      !has_wos[keep]                 ~ "no WoS record (database / regulatory source)",
      ok_name[keep] & ok_year[keep]  ~ "aligned",
      !ok_year[keep]                 ~ "YEAR MISMATCH - verify row alignment",
      TRUE                           ~ "surname not in WoS author list - verify citation key"),
    stringsAsFactors = FALSE)
}

ref_integrity <- data.frame(
  check = c("R1 references cited in data sheets but absent from Table_SM_2.1",
            "R2 keys in Table_SM_2.1 never cited in any data sheet",
            "R3 WoS rows aligned with their reference key",
            "R3 WoS rows needing verification",
            "Reference keys in Table_SM_2.1",
            "  of which carry a WoS record (= included studies)",
            "  of which are non-WoS sources (database / regulatory)",
            "Distinct references cited across the data sheets"),
  value = c(length(R1_missing), length(R2_orphan),
            if (is.null(align)) NA_integer_ else sum(align$check == "aligned"),
            if (is.null(align)) NA_integer_ else
              sum(!align$check %in% c("aligned", "no WoS record (database / regulatory source)")),
            length(ref_keys),
            if (is.na(Nwos)) NA_integer_ else Nwos,
            if (is.na(Nwos)) NA_integer_ else length(ref_keys) - Nwos,
            length(unique(cited_norm))),
  detail = c(if (length(R1_missing)) paste(R1_missing, collapse = "; ") else "none",
             if (length(R2_orphan))  paste(R2_orphan,  collapse = "; ") else "none",
             "leading surname of the key found in the row's WoS author list AND key year = Publication Year",
             "listed row-by-row in sheet 'Reference_alignment'",
             "column A of Table_SM_2.1, after alias canonicalisation",
             "rows carrying a Publication Year",
             "rows with no Publication Year",
             "after alias canonicalisation (REF_ALIASES)"),
  stringsAsFactors = FALSE)

msg("Reference integrity: ", length(ref_keys), " keys | ",
        if (is.na(Nwos)) "?" else Nwos, " WoS studies | ",
        length(R1_missing), " uncatalogued citation(s) | ",
        length(R2_orphan), " orphan key(s)",
        if (!is.null(align)) paste0(" | ", sum(align$check == "aligned"), "/",
                                    sum(align$has_WoS_record), " WoS rows aligned") else "")
if (length(R1_missing))
  msg("References cited but absent from Table_SM_2.1: ",
          paste(R1_missing, collapse = "; "), call. = FALSE)
if (length(R2_orphan))
  msg("  orphan keys (in Table_SM_2.1, never cited): ", paste(R2_orphan, collapse = "; "))
if (!is.null(align)) {
  chk <- align[!align$check %in% c("aligned", "no WoS record (database / regulatory source)"), ]
  if (nrow(chk)) {
    msg("  rows to verify (", nrow(chk), "):")
    if (!QUIET) print(chk[, c("row_in_sheet", "reference_key", "WoS_Publication_Year",
                  "WoS_first_author", "check")], row.names = FALSE)
  }
}

## ---- (C) regulatory P / M / B screening (per compound) --------------------
pmb <- prop %>% group_by(Compound, PCs_class) %>%
    summarise(logKOC = m_median(`log _Koc`), logDOW = m_median(log_Dow),
              logKOW = m_median(log_Kow), .groups = "drop") %>%
    left_join(select(screen, Compound, flag_all, flag_meas, estimate_dependent),
              by = "Compound") %>%
    mutate(
        Mobility = ifelse(is.na(logKOC), "-", ifelse(logKOC < 2, "vM",
                          ifelse(logKOC < 3, "M", "-"))),
        B_trigger_PhaseII = ((is.finite(logKOW) & logKOW >= 3) |
                             (is.finite(logDOW) & logDOW >= 3)),
        PBT_screen_PhaseI = is.finite(logKOW) & logKOW > 4.5,
        Persistent_all  = !is.na(flag_all)  & flag_all  != "\u2014",
        Persistent_meas = !is.na(flag_meas) & flag_meas != "\u2014",
        PM_candidate = Persistent_all  & Mobility %in% c("M", "vM"),
        BP_candidate = B_trigger_PhaseII & Persistent_all) %>%
    arrange(desc(PM_candidate | BP_candidate), PCs_class, Compound) %>%
    transmute(Compound, PCs_class,
              `log KOC` = round(logKOC, 2), Mobility,
              `log DOW` = round(logDOW, 2), `log KOW` = round(logKOW, 2),
              `B trigger (Phase II)` = B_trigger_PhaseII,
              `PBT screen (Phase I)` = PBT_screen_PhaseI,
              `P/vP (all t1/2)` = flag_all, `P/vP (measured)` = flag_meas,
              `P/M candidate` = PM_candidate, `B/P candidate` = BP_candidate,
              persistence_provenance = estimate_dependent)

pmb_summary <- data.frame(
    metric = c("M or vM (mobile)", "  of which vM", "B-trigger (Phase II)",
               "PBT/vPvB screen (Phase I, log KOW > 4.5)",
               "P/M candidates (all t1/2)", "B/P candidates (all t1/2)",
               "P/M candidates retaining persistence on measured data"),
    value  = c(sum(pmb$Mobility %in% c("M","vM")), sum(pmb$Mobility == "vM"),
               sum(pmb$`B trigger (Phase II)`), sum(pmb$`PBT screen (Phase I)`),
               sum(pmb$`P/M candidate`), sum(pmb$`B/P candidate`),
               sum(pmb$`P/M candidate` &
                   pmb$`P/vP (measured)` != "\u2014" & !is.na(pmb$`P/vP (measured)`))),
    stringsAsFactors = FALSE)

## ---- PRISMA template (manual counts; not derivable from SM_2) --------------
prisma <- data.frame(
    stage = c("Records identified (WoS searches)", "Duplicates removed",
              "Records screened", "Records excluded (title/abstract)",
              "Full-text assessed", "Full-text with extractable data",
              "Additional records (citation tracking)", "Studies included (total)"),
    count = c(NA, NA, NA, NA, NA, NA, NA,
              if (is.na(Nwos)) NA_integer_ else Nwos),
    note  = c(rep("ENTER from screening log", 7),
              if (is.na(Nwos)) "ENTER manually (WoS sheet not detected)"
              else "auto: WoS records in Table_SM_2.1"),
    stringsAsFactors = FALSE)

## ---- README sheets ---------------------------------------------------------
readme_1a <- data.frame(Notes = c(
    "Table_SM_2.1_a - bibliometric, study-count and reference-integrity appendix to Table_SM_2.1.",
    if (wos_available)
      paste0("Bibliometrics computed from the WoS export embedded in the REVISED Table_SM_2.1, ",
             "whose reference keys (column A) are aligned ROW BY ROW with the export. ",
             "N(WoS studies) = ", Nwos, "; reference keys = ", length(keys), "; ",
             "non-WoS sources = ", length(keys) - Nwos, ".")
    else
      paste0("Table_SM_2.1 in SM_2.xlsx does not appear to carry a WoS export; ",
             "bibliometric sheets were not produced. To obtain them, paste a WoS export ",
             "(with columns Publication Year, WoS Categories, Reprint Addresses, Addresses) ",
             "into Table_SM_2.1, one row per reference key."),
    "Reference_integrity: R1 = references cited in the data sheets but absent from Table_SM_2.1;",
    "  R2 = keys in Table_SM_2.1 never cited; R3 = row-by-row alignment of each key with its WoS record.",
    "Reference_alignment: the R3 check, row by row (key surname vs WoS author list; key year vs Publication Year).",
    "PRISMA sheet: the search/screening funnel is NOT contained in SM_2; the included-study",
    "  total is filled automatically and the upstream counts must be entered from the screening log."),
    stringsAsFactors = FALSE)

readme_9a <- data.frame(Notes = c(
    "Table_SM_2.9_a - regulatory P/M/B screening appendix to Table_SM_2.9.",
    "Mobility: M when log KOC < 3, vM when log KOC < 2 (Commission Delegated Reg. (EU) 2023/707).",
    "Bioaccumulation: Phase II BCF trigger log KOW >= 3 or log DOW >= 3; Phase I PBT screen log KOW > 4.5 (EMA 2024).",
    "Persistence (P/vP) flags are inherited from the screening module (Table_SM_2.9): the",
    "  '(all t1/2)' column uses every harmonised half-life, the '(measured)' column excludes EPI Suite estimates.",
    "P/M candidate = persistent (any compartment, all t1/2) AND mobile; B/P candidate = B-trigger AND persistent.",
    "persistence_provenance flags whether a compound's persistence rests on EPI Suite estimates.",
    "",
    "PROVENANCE SHEETS. The estimate-sensitivity test applied to persistence in",
    "  Table_SM_2.9 is applied here to the other two screening axes, so that all",
    "  three carry the same measured-versus-estimated distinction.",
    "Property_provenance: per descriptor, the number of compiled records, how many",
    "  are database-derived (EPI Suite, PubChem, DrugBank, identified by source",
    "  reference and never by magnitude), and how many compounds have no measured",
    "  record of that descriptor.",
    "Mobility_provenance: per compound, log KOC and its mobility class computed on",
    "  all records and again on measured records only, with the compounds whose",
    "  class changes listed first. A compound with no measured log KOC is reported",
    "  as 'undetermined' rather than retaining its estimate-derived class.",
    "Bioaccumulation_provenance: per flagged compound, which descriptor fires the",
    "  EMA Phase II trigger and whether that descriptor has measured support."),
    stringsAsFactors = FALSE)

## ---- Table_SM_2.1_a : study counts + bibliometrics + PRISMA template -------
wb_1a <- openxlsx::createWorkbook()
add1a <- function(nm, df) { openxlsx::addWorksheet(wb_1a, nm)
    openxlsx::writeData(wb_1a, nm, df, keepNA = TRUE, na.string = NA_TOKEN) }
add1a("README",              readme_1a)
add1a("Study_counts",        study_counts)
add1a("Reference_integrity", ref_integrity)
if (!is.null(align)) add1a("Reference_alignment", align)
if (wos_available) {
  add1a("Bibliometrics_year",      yr_bins)
  add1a("Bibliometrics_year_full", yr_tab)
  if (!is.null(cat_tab))  add1a("Bibliometrics_area",      cat_tab)
  if (!is.null(ctry_tab)) add1a("Bibliometrics_country",   ctry_tab)
  if (!is.null(cont_tab)) add1a("Bibliometrics_continent", cont_tab)
}
add1a("PRISMA", prisma)
openxlsx::saveWorkbook(wb_1a, file.path(out_dir, "Table_SM_2.1_a.xlsx"), overwrite = TRUE)
wrote("Table_SM_2.1_a.xlsx")

## ---- (D) PROVENANCE OF THE MOBILITY AND BIOACCUMULATION DESCRIPTORS -------
## WHY: the estimate-sensitivity machinery of Section 7 is applied to HALF-LIVES
##      only, so up to this point the persistence column of Table_SM_2.9_a
##      carried a measured-versus-estimated distinction that the mobility and
##      bioaccumulation columns did not.  That asymmetry is not defensible once
##      the manuscript states that the mobility axis is provisional: the claim
##      has to be tabulated, not asserted.  This block supplies it on exactly
##      the same principle used for persistence -- a record is database-derived
##      when its SOURCE REFERENCE is one of the three reference databases, never
##      when its value looks unusual.  Identifying estimates by magnitude would
##      build the conclusion into the selection.
## DATABASES: EPI Suite (U.S. EPA, 2012), PubChem (Kim et al., 2023) and
##      DrugBank (Wishart et al., 2018).  EPI Suite is the only one of the three
##      that supplies log KOC, and it supplies exactly one estimate per compound.
DB_REF_PATTERN <- "^\\s*(U\\.S\\. EPA|Kim et al|Wishart)"

## re-read the harmonised PROPERTY sheet WITH its References column
## (read_harmonized keeps only the analysis variables and drops References)
hp_full   <- apply_corrections(read_sm2(SH_HARM_PROP), SH_HARM_PROP)
hpn       <- norm_names(names(hp_full))
ci_comp_p <- find_col(hpn, "^Compound",   "Compound")
ci_ref_p  <- find_col(hpn, "^References", "References")

prop_prov <- data.frame(
  Compound  = trimws(as.character(hp_full[[ci_comp_p]])),
  Reference = trimws(as.character(hp_full[[ci_ref_p]])),
  stringsAsFactors = FALSE)
prop_prov$is_db <- grepl(DB_REF_PATTERN, prop_prov$Reference)
for (canon in names(prop_patterns))
  prop_prov[[canon]] <- clean_num(hp_full[[ find_col(hpn, prop_patterns[[canon]], canon) ]])
prop_prov <- prop_prov[!is.na(prop_prov$Compound) & prop_prov$Compound != "", ]
## Compound is the join key throughout; trimws matches the treatment applied in
## Section 3D and Section 7 (Table_SM_2.5 stores " Sulfamethoxazole" with a
## leading space, a different key from the fate sheet's "Sulfamethoxazole").
prop_prov$Compound <- trimws(prop_prov$Compound)

## ---- (D1) descriptor-level provenance census ------------------------------
## One row per physicochemical descriptor: how many compiled records it has,
## how many of them are database-derived, and for how many compounds it has no
## measured record at all.  This is the table behind the manuscript statements
## that log DOW carries no database-derived record and that log KOW does.
prov_census <- do.call(rbind, lapply(names(prop_patterns), function(v) {
  x  <- prop_prov[[v]]; ok <- is.finite(x)
  if (!any(ok)) return(NULL)
  cmp   <- prop_prov$Compound[ok]
  isdb  <- prop_prov$is_db[ok]
  nmeas <- tapply(!isdb, cmp, sum)
  data.frame(
    Descriptor                    = v,
    n_records                     = sum(ok),
    n_database_derived            = sum(isdb),
    pct_database_derived          = round(100 * sum(isdb) / sum(ok), 1),
    n_compounds_with_any_value    = length(unique(cmp)),
    n_compounds_with_measured     = sum(nmeas > 0),
    n_compounds_estimate_only     = sum(nmeas == 0),
    stringsAsFactors = FALSE)
}))

## ---- (D2) per-compound MOBILITY provenance sensitivity ---------------------
## The mobility class is recomputed from measured log KOC alone and compared
## with the published class, exactly as the persistence flags are recomputed in
## Section 7.  A compound with no measured log KOC becomes "undetermined"
## rather than silently retaining its estimate-derived class.
mob_class <- function(k) ifelse(is.na(k) | !is.finite(k), "undetermined",
                        ifelse(k < 2, "vM", ifelse(k < 3, "M", "not mobile")))

mob_prov <- prop_prov %>%
  dplyr::group_by(Compound) %>%
  dplyr::summarise(
    n_logKOC          = sum(is.finite(`log _Koc`)),
    n_logKOC_database = sum(is.finite(`log _Koc`) &  is_db),
    n_logKOC_measured = sum(is.finite(`log _Koc`) & !is_db),
    `log KOC (all)`      = m_median(`log _Koc`),
    `log KOC (measured)` = m_median(ifelse(is_db, NA_real_, `log _Koc`)),
    .groups = "drop") %>%
  dplyr::mutate(
    `log KOC (all)`       = round(`log KOC (all)`, 3),
    `log KOC (measured)`  = round(`log KOC (measured)`, 3),
    `Mobility (all)`      = mob_class(`log KOC (all)`),
    `Mobility (measured)` = mob_class(`log KOC (measured)`),
    class_changes = `Mobility (all)` != `Mobility (measured)`,
    evidence_base = ifelse(n_logKOC_measured == 0, "EPI Suite estimate only",
                    ifelse(n_logKOC_database == 0, "measured only",
                                                   "measured + estimate"))) %>%
  dplyr::left_join(dplyr::select(pmb, Compound, PCs_class), by = "Compound") %>%
  dplyr::select(Compound, PCs_class, dplyr::everything()) %>%
  dplyr::arrange(desc(class_changes), PCs_class, Compound)

mob_summary <- data.frame(
  metric = c("Compounds with a log KOC value (any source)",
             "  of which mobile (M or vM), all sources",
             "  of which very mobile (vM), all sources",
             "EPI Suite log KOC estimates in the compilation",
             "Compounds with NO measured log KOC (EPI Suite estimate only)",
             "Compounds with a measured log KOC",
             "  of which mobile (M or vM), measured only",
             "  of which very mobile (vM), measured only",
             "Compounds whose mobility class CHANGES when estimates are removed",
             "Compounds becoming undetermined when estimates are removed"),
  value = c(sum(mob_prov$n_logKOC > 0),
            sum(mob_prov$`Mobility (all)` %in% c("M", "vM")),
            sum(mob_prov$`Mobility (all)` == "vM"),
            sum(mob_prov$n_logKOC_database),
            sum(mob_prov$n_logKOC_measured == 0),
            sum(mob_prov$n_logKOC_measured > 0),
            sum(mob_prov$`Mobility (measured)` %in% c("M", "vM")),
            sum(mob_prov$`Mobility (measured)` == "vM"),
            sum(mob_prov$class_changes),
            sum(mob_prov$`Mobility (measured)` == "undetermined")),
  stringsAsFactors = FALSE)

## ---- (D3) per-compound BIOACCUMULATION trigger provenance ------------------
## Which descriptor fires the EMA Phase II trigger for each flagged compound,
## and whether that descriptor has measured support for that compound.  This is
## the table behind the manuscript statement that the log DOW exceedances rest
## entirely on measured data while most Phase II triggers fire on log KOW.
kow_counts <- prop_prov %>% dplyr::group_by(Compound) %>%
  dplyr::summarise(n_logKOW = sum(is.finite(log_Kow)),
                   n_logKOW_database = sum(is.finite(log_Kow) &  is_db),
                   n_logDOW = sum(is.finite(log_Dow)),
                   n_logDOW_database = sum(is.finite(log_Dow) &  is_db),
                   .groups = "drop")

bio_prov <- pmb %>%
  dplyr::filter(`B trigger (Phase II)`) %>%
  dplyr::select(Compound, PCs_class, `log KOW`, `log DOW`,
                `PBT screen (Phase I)`) %>%
  dplyr::left_join(kow_counts, by = "Compound") %>%
  dplyr::mutate(
    trigger_descriptor = ifelse(
      !is.na(`log DOW`) & `log DOW` >= 3,
      ifelse(!is.na(`log KOW`) & `log KOW` >= 3, "log DOW and log KOW", "log DOW"),
      "log KOW"),
    logKOW_has_measured = (n_logKOW - n_logKOW_database) > 0,
    logDOW_has_measured = (n_logDOW - n_logDOW_database) > 0,
    ## where both descriptors fire, log DOW governs: it is the ionisation-corrected
    ## descriptor and is the one the manuscript uses for ionisable compounds
    trigger_has_measured_support = ifelse(
      grepl("log DOW", trigger_descriptor), logDOW_has_measured, logKOW_has_measured)) %>%
  dplyr::arrange(trigger_descriptor, PCs_class, Compound)

bio_summary <- data.frame(
  metric = c("Compounds above the EMA Phase II trigger",
             "  triggered on log DOW (>= 3)",
             "  triggered on log KOW only (>= 3)",
             "  whose triggering descriptor has NO measured record",
             "Compounds above the Phase I PBT screen (log KOW > 4.5)",
             "log DOW records in the compilation",
             "  of which database-derived",
             "log KOW records in the compilation",
             "  of which database-derived"),
  value = c(nrow(bio_prov),
            sum(grepl("log DOW", bio_prov$trigger_descriptor)),
            sum(bio_prov$trigger_descriptor == "log KOW"),
            sum(!bio_prov$trigger_has_measured_support),
            sum(bio_prov$`PBT screen (Phase I)`),
            prov_census$n_records[prov_census$Descriptor == "log_Dow"],
            prov_census$n_database_derived[prov_census$Descriptor == "log_Dow"],
            prov_census$n_records[prov_census$Descriptor == "log_Kow"],
            prov_census$n_database_derived[prov_census$Descriptor == "log_Kow"]),
  stringsAsFactors = FALSE)

msg("  mobility provenance: ", sum(mob_prov$class_changes), " of ",
        nrow(mob_prov), " mobility classes change when EPI Suite log KOC is removed")

## ---- Table_SM_2.9_a : regulatory P/M/B screening ---------------------------
wb_9a <- openxlsx::createWorkbook()
add9a <- function(nm, df) { openxlsx::addWorksheet(wb_9a, nm)
    openxlsx::writeData(wb_9a, nm, df, keepNA = TRUE, na.string = NA_TOKEN) }
add9a("README",                     readme_9a)
add9a("Regulatory_PMB",             pmb)
add9a("Regulatory_PMB_summary",     pmb_summary)
add9a("Property_provenance",        prov_census)
add9a("Mobility_provenance",        mob_prov)
add9a("Mobility_provenance_summary", mob_summary)
add9a("Bioaccumulation_provenance", bio_prov)
add9a("Bioaccumulation_prov_summary", bio_summary)
openxlsx::saveWorkbook(wb_9a, file.path(out_dir, "Table_SM_2.9_a.xlsx"), overwrite = TRUE)
wrote("Table_SM_2.9_a.xlsx")

msg("Reconciliation written: Table_SM_2.1_a.xlsx, Table_SM_2.9_a.xlsx")
if (wos_available && !is.null(yr_bins) && !is.null(cat_tab) && !is.null(cont_tab))
  msg("  bibliometrics: ", Nwos, " WoS studies | ",
          yr_bins$pct[yr_bins$window == "2019-2024"], "% in 2019-2024 | Env.Sci ",
          cat_tab$pct_of_studies[cat_tab$WoS_Category == "Environmental Sciences"][1],
          "% | Europe ", cont_tab$pct[cont_tab$Continent == "Europe"], "%") else
  msg("  bibliometrics: skipped (Table_SM_2.1 is not a WoS export). ",
          "Study counts: ", dplyr::n_distinct(prop$Compound), " compounds, ",
          dplyr::n_distinct(prop$PCs_class), " therapeutic classes.")
msg("  screening: ", sum(pmb$`P/M candidate`), " P/M candidates, ",
        sum(pmb$`B/P candidate`), " B/P candidates")
msg("Done (Sections 9-10).")

## ###########################################################################
## SECTION 11 - PROVENANCE MODULE: Table_SM_2.10, Table_SM_2.11, Figure B
STAGE("11  provenance -> Table_SM_2.10 / 2.11")

## ---------------------------------------------------------------------------
## Builds, from the corrected harmonised fate sheet (Table_SM_2.6, already
## patched in Section 1b) and the corrected study classification (cls_path):
##   Table_SM_2.10  long-format half-life table: one reported t1/2 per row,
##                  tagged with compartment/condition, source reference and
##                  provenance Type (Field / Laboratory / Database).
##   Table_SM_2.11  three half-life ratio analyses:
##                    (i)   overall         Lab vs Field  (per compound)
##                    (ii)  per compartment Lab vs Field  (per compound)
##                    (iii) per compartment measured vs EPI Suite
##   Figure B       per-compound half-life by data origin, faceted by
##                  compartment (SW | SED | HZ, the compartments with
##                  genuine multi-source overlap).
##
## METHOD (matching / filtering / restructuring)
##   1. The harmonised fate sheet is melted: each of the six half-life
##      columns becomes rows, one per non-empty value, carrying PCs class,
##      Compound and References (already forward-filled by read_harmonized
##      logic; re-applied here since References is not kept by `fate`).
##   2. Each reference is joined to Table_SM_2.1_corrected and reduced to a
##      Type in {Field, Laboratory, Database}. Studies classified as MIXED
##      (Field/Laboratory) are resolved BY MEASUREMENT CONDITION: aerobic/
##      anoxic -> Laboratory (controlled incubation); SW/SED/HZ/GW -> Field
##      (environmental compartment). This is the only rule applied to mixed
##      sources.
##   3. Empty / null / non-detect cells create no row.
##   4. Ratios use compound-level medians so each compound is its own
##      control; within-compartment ratios (ii, iii) avoid mixing
##      compartments.
## Reuses: sm2_path, cls_path, SH_HARM_FATE, hl_patterns, norm_names,
##   find_col, clean_num, out_dir, fig_sum.
## ###########################################################################
suppressWarnings(suppressMessages({
  library(dplyr); library(tidyr); library(stringr); library(readxl); library(openxlsx)

}))

## ---- 11.1  long-format join (Table_SM_2.10) -------------------------------
hf_full <- apply_corrections(read_sm2(SH_HARM_FATE), SH_HARM_FATE)
hfn <- norm_names(names(hf_full))
ci_class11 <- find_col(hfn, id_pat_class, "PCs_class")
ci_comp11  <- find_col(hfn, id_pat_comp,  "Compound")
ci_ref11   <- find_col(hfn, "^References", "References")

base11 <- tibble::tibble(
  PCs_class  = trimws(as.character(hf_full[[ci_class11]])),
  Compound   = trimws(as.character(hf_full[[ci_comp11]])),
  References = trimws(as.character(hf_full[[ci_ref11]])))
## forward-fill block-structured id columns
for (cc11 in c("PCs_class", "Compound", "References")) {
  v <- base11[[cc11]]
  for (i in seq_along(v)) if (is.na(v[i]) || v[i] == "") v[i] <- if (i > 1) v[i - 1] else NA_character_
  base11[[cc11]] <- v
}
base11$PCs_class <- canon_class_chr(base11$PCs_class)  # D18: canonical class labels in Table_SM_2.10
for (canon in names(hl_patterns))
  base11[[canon]] <- as.character(hf_full[[ find_col(hfn, hl_patterns[[canon]], canon) ]])

EMPTY_TOK <- c("", "n.d.", "nd", "na", "n/a", "-", "\u2014")
is_empty11 <- function(x) is.na(x) | tolower(trimws(x)) %in% EMPTY_TOK

long11 <- base11 %>%
  tidyr::pivot_longer(cols = dplyr::all_of(names(hl_patterns)),
                       names_to = "canon", values_to = "t12_raw") %>%
  dplyr::mutate(Compartment = dplyr::recode(canon,
      half_life_SW = "SW", half_life_SED = "SED", half_life_HZ = "HZ",
      half_life_aerobic = "aerobic", half_life_anoxic = "anoxic", half_life_GW = "GW")) %>%
  dplyr::filter(!is_empty11(t12_raw)) %>%
  dplyr::filter(is.na(suppressWarnings(as.numeric(t12_raw))) |
                suppressWarnings(as.numeric(t12_raw)) != 0) %>%  # D10: drop physically-implausible t1/2 = 0 (rule vii), consistent with ZERO_DROP
  dplyr::mutate(`t1/2 (days)` = trimws(t12_raw))

## Embedded study classification (91 references) from the abstract-level audit.
## Used when no external classification file / column is available, so the
## provenance module always runs.  Values: Field-derived, Laboratory-derived,
## Database-derived, Field/Laboratory (mixed).
CLS_EMBEDDED <- c(
  "Aminot et al. (2018)"               = "Laboratory-derived",
  "Arsand et al. (2020)"               = "Field-derived",
  "Babi\u0107 et al. (2007)"           = "Database-derived",
  "Banzhaf et al. (2013)"              = "Field-derived",
  "Barkow et al. (2021)"               = "Field-derived",
  "Bertelkamp et al. (2016)"           = "Laboratory-derived",
  "Bielen et al. (2017)"               = "Field-derived",
  "Burke et al. (2013)"                = "Laboratory-derived",
  "Burke et al. (2014)"                = "Laboratory-derived",
  "Carlson and Mabury (2006)"          = "Field/Laboratory (mixed)",
  "Cardoza et al. (2005)"              = "Field/Laboratory (mixed)",
  "Chen et al. (2013)"                 = "Laboratory-derived",
  "Chen et al. (2018a)"                = "Field-derived",
  "Cheng et al. (2014)"                = "Field-derived",
  "Dibyanshu et al. (2024)"            = "Field-derived",
  "Drillia et al. (2005)"              = "Laboratory-derived",
  "Fan et al. (2025)"                  = "Laboratory-derived",
  "Fernandes et al. (2020)"            = "Field-derived",
  "Fu et al. (2022)"                   = "Field-derived",
  "Gan et al. (2023)"                  = "Laboratory-derived",
  "Garcia-Valverde et al. (2021)"      = "Laboratory-derived",
  "Gros et al. (2021)"                 = "Field-derived",
  "Guan et al. (2023)"                 = "Field-derived",
  "Guillet et al. (2019)"              = "Field-derived",
  "Hari et al. (2005)"                 = "Laboratory-derived",
  "He et al. (2015)"                   = "Field-derived",
  "Hollender et al. (2018)"            = "Field-derived",
  "Hu and Coats (2007)"                = "Laboratory-derived",
  "Hu et al. (2012)"                   = "Field-derived",
  "Hu et al. (2020)"                   = "Field-derived",
  "Jaeger et al. (2021)"               = "Laboratory-derived",
  "Jaukovi\u0107 et al. (2014)"        = "Field-derived",
  "Jones et al. (2002)"                = "Database-derived",
  "Jurado et al. (2020)"               = "Field-derived",
  "Kairigo et al. (2020a)"             = "Field-derived",
  "Kang et al. (2012)"                 = "Laboratory-derived",
  "Kibuye et al. (2019)"               = "Field-derived",
  "Kim et al. (2023)"                  = "Database-derived",
  "Kivits et al. (2018)"               = "Field-derived",
  "Kode\u0161ov\u00e1 et al. (2015)"   = "Laboratory-derived",
  "Kunkel and Radke (2011)"            = "Field-derived",
  "Lahti and Oikari (2011)"            = "Laboratory-derived",
  "Lewandowski et al. (2011)"          = "Field-derived",
  "Li et al. (2010)"                   = "Laboratory-derived",
  "Li and Radke (2015)"                = "Laboratory-derived",
  "Li et al. (2019)"                   = "Laboratory-derived",
  "Li S. et al. (2022)"                = "Field-derived",
  "Li Y. et al. (2023)"                = "Field-derived",
  "Lin and Gan (2011)"                 = "Laboratory-derived",
  "Liu et al. (2012)"                  = "Laboratory-derived",
  "Liu et al. (2019)"                  = "Laboratory-derived",
  "Ma et al. (2015)"                   = "Field-derived",
  "Ma et al. (2022)"                   = "Field-derived",
  "Matviichuk et al. (2022)"           = "Field-derived",
  "Mechelke et al. (2019)"             = "Field/Laboratory (mixed)",
  "Menz et al. (2018)"                 = "Laboratory-derived",
  "Mueller et al. (2021)"              = "Field-derived",
  "Mueller et al. (2022)"              = "Field-derived",
  "Nguyen et al. (2015)"               = "Field/Laboratory (mixed)",
  "Palma et al. (2020)"                = "Field-derived",
  "Pan and Chu (2016)"                 = "Laboratory-derived",
  "Peralta\u2013Maraver et al. (2019)" = "Field-derived",
  "Pereira et al. (2017)"              = "Field-derived",
  "Posselt et al. (2018)"              = "Field-derived",
  "Posselt et al. (2020)"              = "Laboratory-derived",
  "Li Q. et al. (2023)"                = "Laboratory-derived",  # key as spelled in Table_SM_2.1
  "Rab\u00f8lle and Spliid (2000)"     = "Laboratory-derived",
  "Schaper et al. (2018a)"             = "Field-derived",
  "Schaper et al. (2018b)"             = "Field-derived",
  "Schaper et al. (2019)"              = "Field-derived",
  "Scheytt et al. (2006)"              = "Laboratory-derived",
  "Schmidtov\u00e1 et al. (2020)"      = "Laboratory-derived",
  "Sharma et al. (2019)"               = "Field-derived",
  "Sonne et al. (2017)"                = "Field/Laboratory (mixed)",
  "Stuart et al. (2014)"               = "Field-derived",
  "U.S. EPA (2012)"                    = "Database-derived",
  "Viana et al. (2021)"                = "Field-derived",
  "Wishart et al. (2018)"              = "Database-derived",
  "Xie et al. (2022)"                  = "Laboratory-derived",
  "Xu et al. (2025)"                   = "Laboratory-derived",
  "Yamamoto et al. (2009)"             = "Laboratory-derived",
  "Yang et al. (2018)"                 = "Field-derived",
  "Yang et al. (2022)"                 = "Laboratory-derived",
  "Yang J. et al. (2022)"              = "Field-derived",
  "Zhang et al. (2021)"                = "Field-derived",
  "Zhang et al. (2022)"                = "Laboratory-derived",
  "Zhang et al. (2024)"                = "Field/Laboratory (mixed)",
  "Zhao et al. (2024)"                 = "Laboratory-derived",
  "Zhu et al. (2023)"                  = "Field/Laboratory (mixed)",
  "Zuo et al. (2021)"                  = "Field/Laboratory (mixed)"
)

## ---- classification -> Type (mixed resolved BY MEASUREMENT CONDITION) -----
## The provenance Type of every half-life comes from the study classification.
##
##
## SOURCE PRIORITY
##   1. Table_SM_2.1_corrected.xlsx (Stage-1 abstract audit) when present.
##   2. A classification column inside SM_2.xlsx > Table_SM_2.1, if one exists.
##   3. Otherwise: stop with an explicit, actionable message (no silent garbage).
## key normaliser: trim, collapse whitespace, drop trailing period, fold accents
## (accent folding makes matching robust to encoding differences on Windows)
fold_ascii <- function(x) {
  y <- iconv(x, from = "UTF-8", to = "ASCII//TRANSLIT")
  y[is.na(y)] <- x[is.na(y)]
  y
}
nref11 <- function(x) tolower(fold_ascii(trimws(sub("\\.$", "", norm_names(x)))))


CLS_VALUE_PAT  <- "field[- ]?derived|laborator|database[- ]?derived|modeling|field/laborator"
CLS_REJECT_PAT <- "^(both|fate|properties)$"
looks_like_classification <- function(v) {
  v <- v[!is.na(v) & nzchar(trimws(v))]
  if (!length(v)) return(FALSE)
  if (mean(grepl(CLS_REJECT_PAT, trimws(v), ignore.case = TRUE)) >= 0.5) return(FALSE)
  ## require a clear majority of short, category-like strings (citations are long)
  hits  <- grepl(CLS_VALUE_PAT, v, ignore.case = TRUE)
  short <- nchar(v) <= 60
  mean(hits & short) >= 0.5
}

ref_lookup11 <- character(0)
cls_source   <- NA_character_

## ---- source 1: Stage-1 corrected classification file -----------------------
if (!is.null(cls_path)) {
  cls_ov <- tryCatch(readxl::read_excel(cls_path, col_types = "text"),
                     error = function(e) NULL)
  if (!is.null(cls_ov)) {
    nm_ov <- norm_names(names(cls_ov)); lo <- tolower(nm_ov)
    k_r <- which(lo == "reference")[1]
    if (is.na(k_r)) k_r <- which(grepl("^ref", lo))[1]
    k_t <- which(lo == "type_corrected")[1]
    if (is.na(k_t)) k_t <- which(grepl("type_correct", lo))[1]
    if (is.na(k_t)) k_t <- which(vapply(cls_ov, function(cc)
                        looks_like_classification(as.character(cc)), logical(1)))[1]
    if (!is.na(k_r) && !is.na(k_t)) {
      ref_lookup11 <- setNames(tolower(trimws(as.character(cls_ov[[k_t]]))),
                               nref11(as.character(cls_ov[[k_r]])))
      cls_source <- paste0(basename(cls_path), " [", nm_ov[k_r], " -> ", nm_ov[k_t], "]")
    }
  }
}

## ---- source 2: a genuine classification column inside SM_2.xlsx ------------
if (!length(ref_lookup11)) {
  sm2_cls_raw <- tryCatch(trim_ref_block(read_sm2("Table_SM_2.1"), quiet = TRUE),
                          error = function(e) NULL)
  if (!is.null(sm2_cls_raw)) {
    nm_raw <- norm_names(names(sm2_cls_raw))
    k_ref  <- which(grepl("^References all|^Reference", nm_raw, ignore.case = TRUE))[1]
    if (is.na(k_ref)) k_ref <- 1L
    ## scan EVERY column for classification-like CONTENT (no positional fallback)
    cand <- which(vapply(sm2_cls_raw, function(cc)
                    looks_like_classification(as.character(cc)), logical(1)))
    cand <- setdiff(cand, k_ref)
    if (length(cand)) {
      k_typ <- cand[1]
      ref_lookup11 <- setNames(tolower(trimws(as.character(sm2_cls_raw[[k_typ]]))),
                               nref11(as.character(sm2_cls_raw[[k_ref]])))
      cls_source <- paste0("SM_2.xlsx > Table_SM_2.1 [", nm_raw[k_ref],
                           " -> ", nm_raw[k_typ], "]")
    }
  }
}

## ---- source 3: embedded classification (always available) ------------------
## Guarantees the provenance module runs even with no external file and no
## classification column in SM_2.xlsx.  This is the audited classification
## itself, held in the script so it is versioned with the code and visible.
if (!length(ref_lookup11)) {
  ref_lookup11 <- setNames(tolower(trimws(unname(CLS_EMBEDDED))),
                           nref11(names(CLS_EMBEDDED)))
  cls_source   <- "embedded classification in this script (abstract-level audit)"
}

## ---- clean the lookup ------------------------------------------------------
ref_lookup11 <- ref_lookup11[nzchar(names(ref_lookup11)) &
                             names(ref_lookup11) != "na" &
                             !is.na(ref_lookup11) & nzchar(ref_lookup11)]

if (!length(ref_lookup11))
  stop("Section 11: no study classification available (embedded table is empty).",
       call. = FALSE)

msg("  Section 11 classification: ", length(ref_lookup11),
        " entries from ", cls_source)


study_type11 <- function(ref, comp) {
  ## use [ ] (not [[ ]]): a missing name yields NA rather than an error
  s <- unname(ref_lookup11[ nref11(ref) ])
  if (length(s) == 0L || is.na(s) || s == "") return("UNMATCHED")
  if (grepl("database", s)) return("Database")
  if (grepl("field", s) && grepl("laborator", s))
    return(if (comp %in% c("aerobic", "anoxic")) "Lab" else "Field")
  if (grepl("field", s))      return("Field")
  if (grepl("laborator", s))  return("Lab")
  "UNMATCHED"
}
long11$Type <- mapply(study_type11, long11$References, long11$Compartment)

## ---- data-quality guard: unclassified references ---------------------------

um11    <- sort(unique(long11$References[long11$Type == "UNMATCHED"]))
n_um11  <- sum(long11$Type == "UNMATCHED")
pct_um  <- if (nrow(long11)) round(100 * n_um11 / nrow(long11), 1) else 0
msg("  Section 11 provenance: ",
        paste(names(table(long11$Type)), table(long11$Type),
              sep = "=", collapse = ", "),
        "  (", pct_um, "% unclassified)")
if (length(um11))
  msg("  WARNING (Section 11) - unclassified references (", length(um11), "): ",
          paste(um11, collapse = "; "))
if (pct_um > 5)
  stop("Section 11: ", pct_um, "% of half-life records are UNCLASSIFIED (", n_um11,
       " of ", nrow(long11), ").\n",
       "  The study classification does not cover the references in Table_SM_2.6.\n",
       "  Classification source used: ", cls_source, "\n",
       "  ACTION: supply Table_SM_2.1_corrected.xlsx (columns 'Reference' and\n",
       "  'Type_corrected') beside SM_2.xlsx so every reference is classified.\n",
       "  Aborting rather than writing empty/meaningless ratio tables.", call. = FALSE)

comp_levels11 <- c("SW", "SED", "HZ", "aerobic", "anoxic", "GW")
tab210 <- long11 %>%
  dplyr::transmute(`PCs class` = PCs_class, Compound,
                    `Compartments/Conditions` = factor(Compartment, levels = comp_levels11),
                    `t1/2 (days)`, Type, References) %>%
  dplyr::arrange(`PCs class`, Compound, `Compartments/Conditions`) %>%
  dplyr::mutate(`Compartments/Conditions` = as.character(`Compartments/Conditions`))

wb210 <- createWorkbook(); addWorksheet(wb210, "Table_SM_2.10")
writeData(wb210, "Table_SM_2.10", tab210, keepNA = TRUE, na.string = NA_TOKEN)
saveWorkbook(wb210, file.path(out_dir, "Table_SM_2.10.xlsx"), overwrite = TRUE)
wrote("Table_SM_2.10.xlsx")
msg("Table_SM_2.10: ", nrow(tab210), " rows -> Table_SM_2.10.xlsx  | Types: ",
        paste(names(table(long11$Type)), table(long11$Type), sep = "=", collapse = ", "))

## numeric working frame
to_num11 <- function(x) {
  ## defensive: read_sm2() has already mapped U+2212/NBSP, but repeat it here so
  ## the parser is correct even if called on un-normalised text.  A range such as
  ## "0.5-8.7" must reduce to its MIDPOINT (rule ii) - if the separator were not
  ## recognised the parser would silently return the LOWER bound instead.
  s <- trimws(norm_text(x))
  if (is.na(s) || tolower(s) %in% EMPTY_TOK) return(NA_real_)
  s <- gsub("[\u2245\u2248~<>]", "", s)
  s <- trimws(sub("\u00b1.*$", "", s))
  m <- regmatches(s, regexec("^\\s*([0-9]+\\.?[0-9]*)\\s*-\\s*([0-9]+\\.?[0-9]*)\\s*$", s))[[1]]
  if (length(m) == 3) return((as.numeric(m[2]) + as.numeric(m[3])) / 2)
  v <- regmatches(s, regexpr("-?[0-9]+\\.?[0-9]*([eE]-?[0-9]+)?", s))
  if (length(v) == 0 || v == "") return(NA_real_)
  as.numeric(v)
}
long11$val <- vapply(long11$`t1/2 (days)`, to_num11, numeric(1))
num11 <- dplyr::filter(long11, is.finite(val) & val > 0)
mf11  <- dplyr::filter(num11, Type %in% c("Field", "Lab"))

big11 <- dplyr::filter(num11, val > 1000) %>%
  dplyr::select(Compound, Compartment, Type, val, References)
if (nrow(big11)) {
  msg("DATA-QC (Section 11): half-lives > 1000 d (verify units/source):")
  if (!QUIET) print(as.data.frame(big11))
}

## ---- 11.2  Table_SM_2.11 ratio analyses -----------------------------------
lab_i   <- mf11 %>% dplyr::filter(Type == "Lab")   %>% dplyr::group_by(Compound) %>%
  dplyr::summarise(Lab_median_d = median(val), n_Lab = dplyr::n(), .groups = "drop")
field_i <- mf11 %>% dplyr::filter(Type == "Field") %>% dplyr::group_by(Compound) %>%
  dplyr::summarise(Field_median_d = median(val), n_Field = dplyr::n(), .groups = "drop")
tab_i <- dplyr::inner_join(lab_i, field_i, by = "Compound") %>%
  dplyr::mutate(ratio_Lab_over_Field = Lab_median_d / Field_median_d,
                fold_difference = pmax(ratio_Lab_over_Field, 1 / ratio_Lab_over_Field)) %>%
  dplyr::select(Compound, n_Lab, Lab_median_d, n_Field, Field_median_d,
                ratio_Lab_over_Field, fold_difference) %>%
  dplyr::arrange(ratio_Lab_over_Field)

lab_ii   <- mf11 %>% dplyr::filter(Type == "Lab")   %>% dplyr::group_by(Compound, Compartment) %>%
  dplyr::summarise(Lab_median_d = median(val), n_Lab = dplyr::n(), .groups = "drop")
field_ii <- mf11 %>% dplyr::filter(Type == "Field") %>% dplyr::group_by(Compound, Compartment) %>%
  dplyr::summarise(Field_median_d = median(val), n_Field = dplyr::n(), .groups = "drop")
tab_ii <- dplyr::inner_join(lab_ii, field_ii, by = c("Compound", "Compartment")) %>%
  dplyr::mutate(ratio_Lab_over_Field = Lab_median_d / Field_median_d) %>%
  dplyr::rename(`Compartments/Conditions` = Compartment) %>%
  dplyr::select(Compound, `Compartments/Conditions`, n_Lab, Lab_median_d,
                n_Field, Field_median_d, ratio_Lab_over_Field) %>%
  dplyr::arrange(`Compartments/Conditions`, ratio_Lab_over_Field)

epi11  <- num11 %>% dplyr::filter(Type == "Database") %>% dplyr::group_by(Compound, Compartment) %>%
  dplyr::summarise(EPI_Suite_d = median(val), .groups = "drop")
meas11 <- mf11 %>% dplyr::group_by(Compound, Compartment, Type) %>%
  dplyr::summarise(Measured_median_d = median(val), .groups = "drop")
tab_iii <- dplyr::inner_join(meas11, epi11, by = c("Compound", "Compartment")) %>%
  dplyr::mutate(ratio_Measured_over_EPI = Measured_median_d / EPI_Suite_d) %>%
  dplyr::rename(`Compartments/Conditions` = Compartment, Measured_type = Type) %>%
  dplyr::select(Compound, `Compartments/Conditions`, Measured_type,
                Measured_median_d, EPI_Suite_d, ratio_Measured_over_EPI) %>%
  dplyr::arrange(`Compartments/Conditions`, Measured_type, ratio_Measured_over_EPI)

## ---- provenance-coverage sheet: WHY each compound is / isn't in the ratios --
## For EVERY compound that has any half-life, record how many Field / Lab /
## Database t1/2 records it has, and flag which ratio comparisons are therefore
## computable.  A Lab-vs-Field ratio needs >=1 Lab AND >=1 Field value; a
## measured-vs-EPI ratio needs >=1 measured (Field or Lab) AND >=1 Database
## value.  This documents, transparently, that ratios are reported only where
## a comparison is mathematically defined (not all 48 PCs qualify).
cov11 <- num11 %>%
  dplyr::group_by(Compound) %>%
  dplyr::summarise(
    n_Field    = sum(Type == "Field"),
    n_Lab      = sum(Type == "Lab"),
    n_Database = sum(Type == "Database"),
    .groups = "drop") %>%
  dplyr::mutate(
    in_ratio_Lab_vs_Field  = ifelse(n_Lab > 0 & n_Field > 0, "yes", "no"),
    in_ratio_measured_vsEPI = ifelse((n_Lab > 0 | n_Field > 0) & n_Database > 0,
                                      "yes", "no"),
    reason_if_no_LabField  = dplyr::case_when(
      n_Lab > 0 & n_Field > 0            ~ "",
      n_Lab == 0 & n_Field > 0           ~ "field only (no laboratory value)",
      n_Lab > 0 & n_Field == 0           ~ "laboratory only (no field value)",
      TRUE                                ~ "no measured value")) %>%
  dplyr::arrange(dplyr::desc(in_ratio_Lab_vs_Field), Compound)

n_qual_i <- sum(cov11$in_ratio_Lab_vs_Field == "yes")

## ---- (iii-b) POOLED measured-vs-EPI: one measured median per compound and -----
## compartment.  Sheet (iii) compares the FIELD median and the LABORATORY median
## against the same EPI Suite value SEPARATELY, so a compound holding both source
## types contributes two rows: the 68 rows of sheet (iii) are 68 comparisons over
## 54 distinct compound-compartment pairs.  Sheet (iii-b) pools field and
## laboratory records into a single measured median per compound and compartment,
## so every pair is counted once.  Reported as the robustness check for the
## measured-to-estimated ratio (manuscript Section 3.4).
meas11p <- mf11 %>% dplyr::group_by(Compound, Compartment) %>%
  dplyr::summarise(n_measured = dplyr::n(),
                   n_source_types = dplyr::n_distinct(Type),
                   Measured_median_d = median(val), .groups = "drop")
tab_iiib <- dplyr::inner_join(meas11p, epi11, by = c("Compound", "Compartment")) %>%
  dplyr::mutate(ratio_Measured_over_EPI = Measured_median_d / EPI_Suite_d) %>%
  dplyr::rename(`Compartments/Conditions` = Compartment) %>%
  dplyr::select(Compound, `Compartments/Conditions`, n_measured, n_source_types,
                Measured_median_d, EPI_Suite_d, ratio_Measured_over_EPI) %>%
  dplyr::arrange(`Compartments/Conditions`, ratio_Measured_over_EPI)

## ---- (iv) HYPORHEIC vs REDOX half-lives ---------------------------------
## Compounds carrying BOTH a hyporheic-zone half-life and at least one incubation
## half-life (aerobic and/or anoxic).  The incubation reference value is the
## median of the compound's aerobic and anoxic medians.  Quantifies the divergence
## discussed in manuscript Section 4.3 (operational vs intrinsic half-lives).
inc11 <- fate_cmpd %>%
  ## columns are addressed through resp_names, so this block is unaffected by
  ## the labelling of the response columns (Section 6)
  dplyr::transmute(Compound,
                   HZ      = .data[[resp_names[3]]],
                   aerobic = .data[[resp_names[4]]],
                   anoxic  = .data[[resp_names[5]]]) %>%
  dplyr::filter(is.finite(HZ) & (is.finite(aerobic) | is.finite(anoxic)))
inc11$incubation_median_d <- vapply(seq_len(nrow(inc11)), function(i) {
  v <- c(inc11$aerobic[i], inc11$anoxic[i]); v <- v[is.finite(v)]
  stats::median(v)
}, numeric(1))
tab_iv <- inc11 %>%
  dplyr::mutate(
    fold_difference = pmax(HZ, incubation_median_d) / pmin(HZ, incubation_median_d),
    direction = ifelse(HZ < incubation_median_d, "HZ shorter", "HZ longer")) %>%
  dplyr::rename(`t1/2 HZ (d)` = HZ, `t1/2 aerobic (d)` = aerobic,
                `t1/2 anoxic (d)` = anoxic) %>%
  dplyr::arrange(dplyr::desc(fold_difference))

summ11 <- data.frame(
  metric = c("(i)   compounds with BOTH lab and field half-lives",
             "(i)   median lab/field ratio",
             "(ii)  compound-compartment pairs with BOTH lab and field",
             "(ii)  median lab/field ratio",
             "(iii) measured-vs-EPI comparisons (field and lab separately)",
             "(iii) distinct compound-compartment pairs behind them",
             "(iii) median measured/EPI ratio",
             "(iii) comparisons where measured > EPI",
             "(iii-b) pooled compound-compartment pairs (field+lab combined)",
             "(iii-b) median measured/EPI ratio (pooled)",
             "(iii-b) implied EPI Suite over-prediction factor (pooled)",
             "(iii-b) pairs where measured > EPI (pooled)",
             "(iv)  compounds with HZ and incubation half-lives",
             "(iv)  median HZ-vs-incubation fold difference",
             "(iv)  compounds with HZ shorter than incubation"),
  value = c(nrow(tab_i),  round(median(tab_i$ratio_Lab_over_Field), 3),
            nrow(tab_ii), round(median(tab_ii$ratio_Lab_over_Field), 3),
            nrow(tab_iii),
            nrow(dplyr::distinct(tab_iii[, c("Compound", "Compartments/Conditions")])),
            round(median(tab_iii$ratio_Measured_over_EPI), 3),
            sum(tab_iii$ratio_Measured_over_EPI > 1),
            nrow(tab_iiib), round(median(tab_iiib$ratio_Measured_over_EPI), 3),
            round(1 / median(tab_iiib$ratio_Measured_over_EPI), 2),
            sum(tab_iiib$ratio_Measured_over_EPI > 1),
            nrow(tab_iv), round(median(tab_iv$fold_difference), 2),
            sum(tab_iv$direction == "HZ shorter")),
  stringsAsFactors = FALSE)

wb211 <- createWorkbook()
addWorksheet(wb211, "coverage_all_PCs");        writeData(wb211, "coverage_all_PCs", cov11)
addWorksheet(wb211, "i_overall_Lab_vs_Field");  writeData(wb211, "i_overall_Lab_vs_Field", tab_i)
addWorksheet(wb211, "ii_by_compartment");       writeData(wb211, "ii_by_compartment", tab_ii)
addWorksheet(wb211, "iii_vs_EPI_Suite");        writeData(wb211, "iii_vs_EPI_Suite", tab_iii)
addWorksheet(wb211, "iiib_vs_EPI_pooled");      writeData(wb211, "iiib_vs_EPI_pooled", tab_iiib)
addWorksheet(wb211, "iv_HZ_vs_incubation");     writeData(wb211, "iv_HZ_vs_incubation", tab_iv)
addWorksheet(wb211, "summary");                 writeData(wb211, "summary", summ11)
saveWorkbook(wb211, file.path(out_dir, "Table_SM_2.11.xlsx"), overwrite = TRUE)
wrote("Table_SM_2.11.xlsx")

## ---- 11.3  Figure B: per-compound source comparison -----------------------
## LEGIBILITY AT 100% ZOOM (JCH requirement)
## ---------------------------------------------------------------------------

FIGB_W   <- 7.5    # inches (190 mm, JCH double column) - do not resize in Word
FIGB_H   <- 9.0    # inches (fits the 33-row SW panel legibly; content fills ~98% of height)
FIGB_DPI <- 600    # publication raster resolution
FIGB_FS  <- 1.00   # global font multiplier (raise to enlarge ALL text at once)
fs <- function(pt) pt * FIGB_FS

PANELS11 <- c("SW", "SED", "HZ")   # compartments with genuine multi-source overlap
med11 <- num11 %>% dplyr::group_by(Compartment, Compound, Type) %>%
  dplyr::summarise(val = median(val), .groups = "drop")

pal11 <- c(Field = "#2166ac", Lab = "#d6604d", Database = "#737373")
shp11 <- c(Field = 21L,       Lab = 24L,       Database = 22L)
lab11 <- c(Field = "Field (in-situ)", Lab = "Laboratory (controlled)",
           Database = "Database (EPI Suite)")

make_panelB <- function(comp) {
  sub  <- dplyr::filter(med11, Compartment == comp)
  if (nrow(sub) == 0) return(NULL)
  keep <- sub %>% dplyr::group_by(Compound) %>%
    dplyr::summarise(nt = dplyr::n_distinct(Type), .groups = "drop") %>%
    dplyr::filter(nt >= 2) %>% dplyr::pull(Compound)
  sub  <- dplyr::filter(sub, Compound %in% keep)
  if (nrow(sub) == 0) return(NULL)

  ord <- sub %>% dplyr::group_by(Compound) %>%
    dplyr::summarise(
      k = if (any(Type == "Field")) stats::median(val[Type == "Field"])
          else stats::median(val), .groups = "drop") %>%
    dplyr::arrange(k) %>% dplyr::pull(Compound)
  sub$Compound <- factor(sub$Compound, levels = ord)
  sub$Type     <- factor(sub$Type, levels = c("Field", "Lab", "Database"))

  seg <- sub %>% dplyr::group_by(Compound) %>%
    dplyr::summarise(lo = min(val), hi = max(val), .groups = "drop")

  ggplot(sub, aes(x = val, y = Compound)) +
    geom_segment(data = seg,
                 aes(x = lo, xend = hi, y = Compound, yend = Compound),
                 inherit.aes = FALSE, colour = "grey75", linewidth = 0.45) +
    geom_point(aes(fill = Type, shape = Type),
               size = 2.4, colour = "white", stroke = 0.35) +
    scale_fill_manual(values = pal11, breaks = names(lab11), labels = lab11,
                      drop = FALSE, name = NULL) +
    scale_shape_manual(values = shp11, breaks = names(lab11), labels = lab11,
                       drop = FALSE, name = NULL) +
    ## few, well-spaced x breaks: crowded tick labels are the second cause of
    ## illegibility once the canvas is narrow
    scale_x_log10(n.breaks = 4, labels = plain_lab) +   # fixed 2-decimal labels
    labs(title = paste0(comp, "  (n = ", length(ord), ")"),
         x = "t\u00bd (days, log scale)", y = NULL) +
    theme_bw(base_size = fs(9)) +
    theme(
      plot.title         = element_text(face = "bold", size = fs(9.5), hjust = 0.5,
                                        margin = margin(b = 3)),
      axis.title.x       = element_text(size = fs(8.5), margin = margin(t = 3)),
      axis.text.y        = element_text(size = fs(8), face = "bold", colour = "grey10",
                                        margin = margin(r = 1)),
      axis.text.x        = element_text(size = fs(7.5)),
      panel.grid.minor   = element_blank(),
      panel.grid.major.y = element_line(colour = "grey93", linewidth = 0.25),
      panel.grid.major.x = element_line(colour = "grey88", linewidth = 0.3),
      panel.spacing      = unit(4, "pt"),
      plot.margin        = margin(t = 2, r = 4, b = 2, l = 2),
      ## Legend at the BOTTOM: with guides = "collect", patchwork wedges a
      ## top-placed legend between the panel titles and panels, cropping its keys.
      ## Bottom placement is also the usual convention for a shared legend.
      legend.position    = "bottom",
      legend.text        = element_text(size = fs(8)),
      legend.margin      = margin(t = 6, r = 2, b = 2, l = 2),
      legend.key.size    = unit(16, "pt")) +
    guides(fill  = guide_legend(nrow = 1, override.aes = list(size = 2.8)),
           shape = guide_legend(nrow = 1, override.aes = list(size = 2.8)))
}

panelsB <- Filter(Negate(is.null), lapply(PANELS11, make_panelB))

if (length(panelsB) == 0) {
  msg("Figure B: no compartment has compounds with >=2 provenance types; figure skipped.")
} else {
  figB <- patchwork::wrap_plots(panelsB, nrow = 1) +
    patchwork::plot_layout(guides = "collect")
  figB <- figB & theme(legend.position = "bottom",
                       legend.box.spacing = unit(10, "pt"),
                       plot.margin = margin(t = 6, r = 6, b = 10, l = 5))

  wrote("FigB_source_comparison.png / .pdf")
  ggsave(file.path(fig_sum, "FigB_source_comparison.png"), figB,
         width = FIGB_W, height = FIGB_H, units = "in", dpi = FIGB_DPI)
  ggsave(file.path(fig_sum, "FigB_source_comparison.pdf"), figB,
         width = FIGB_W, height = FIGB_H, units = "in")   # vector: always sharp
  msg("Figure B written (", FIGB_W, " x ", FIGB_H, " in @ ", FIGB_DPI,
          " dpi; y-axis labels ", fs(8), " pt) -> ", fig_sum,
          "  (FigB_source_comparison.*)")
  msg("  Insert at NATIVE size (7.5 in / 190 mm wide). Do not resize in Word, ",
          "or the text will shrink proportionally. Raise FIGB_FS to enlarge all text.")
}


FIGB_CAPTION <- paste0(
  "Figure B. Per-compound median half-life by data origin, within the three ",
  "compartments with genuine multi-source overlap (surface water, sediment and the ",
  "hyporheic zone). Only compounds with at least two distinct provenance types are ",
  "shown; n is the number of compounds per panel. Compounds are ordered by their ",
  "field median (or overall median where no field value is available). The grey ",
  "segment spans the minimum-to-maximum range across provenance types. Symbols: ",
  "field (blue circle), laboratory (red triangle) and database / EPI Suite ",
  "(grey square). The x-axis is logarithmic.")
msg("Done (Section 11 - Provenance module).")
msg("Section 11 complete.")

## ###########################################################################
## SECTION 12 - DATA-INTEGRITY AUDIT (internal; no workbook is exported)
STAGE("12  data-integrity audit")

## ---------------------------------------------------------------------------
## The evidence that the input file is clean, that nothing was silently lost and
## that nothing was silently changed.  Seven sheets:
##   encoding_by_column   every Unicode substitution, summarised by sheet+column
##   encoding_cells       every single cell that was normalised (before -> after)
##   reference_aliases    every reference key canonicalised (REF_ALIASES)
##   numeric_recast       the 16 analysis columns of part (B): non-empty cells,
##                        cells that parsed to a number, cells that did not
##   unparsable_cells     any cell that failed the recast (must be EMPTY)
##   AGG_reducer_check    AGG percentiles under median- vs mean-reduction
##   extreme_halflives    every t1/2 > 1000 d, FLAGGED for verification - the
##                        pipeline never alters them (includes Tylosin 45,145 d)
## SM_2.xlsx itself is never modified.
## ###########################################################################
extreme_hl <- num11 %>%
  dplyr::filter(val > 1000) %>%
  dplyr::transmute(Compound, `Compartments/Conditions` = Compartment,
                   `t1/2 (days)` = val, Type, References,
                   note = "retained as reported; flagged for verification only") %>%
  dplyr::arrange(dplyr::desc(`t1/2 (days)`))

readme_audit <- data.frame(Notes = c(
  "SM_2_encoding_audit.xlsx - data-integrity audit of the INPUT workbook SM_2.xlsx.",
  "Produced by analysis_pipeline_2026.R. SM_2.xlsx is NEVER modified: all",
  "  normalisation is applied in memory, at read time, and is logged here in full.",
  "",
  "ENCODING (Section 1b of the script):",
  "  U+00A0 non-breaking space (and the other invisible spaces) -> DELETED.",
  "  U+2212 minus sign (and U+2010..U+2015) -> REPLACED by the ASCII hyphen '-'.",
  "  U+2212 is NOT stripped: in SM_2 it serves as a negative sign (log Koc = -1.16),",
  "    as a scientific-notation exponent (Kd = 2.62E-01) and as a range separator",
  "    (foc = 0.5-8.7). Deleting it would flip the sign of the first and multiply",
  "    the second by 100; replacing it preserves all three readings.",
  "  Greek mu (U+03BC) in headers -> micro sign (U+00B5).",
  "",
  "NUMERIC RECAST (part B): every non-empty cell of the sixteen analysis columns",
  "  must parse to a number after normalisation. 'unparsable_cells' must be empty;",
  "  if it is not, those values would otherwise have become NA silently.",
  "",
  "VALUE CORRECTIONS: none. Every compiled value is used exactly as reported,",
  "  including Tylosin t1/2 SED = 45,145 d (Hu and Coats, 2007), which the author",
  "  has verified against the primary source. 'extreme_halflives' flags every",
  "  t1/2 > 1000 d for reader verification; no value is altered.",
  "",
  "AGG_reducer_check: the compound-weighted scheme reduces each compound to its",
  "  MEDIAN (manuscript Section 2.5). This sheet shows what the AGG percentiles",
  "  would be under mean-reduction instead, so the consequence of the rule is explicit."),
  stringsAsFactors = FALSE)

## The data-integrity audit (encoding substitutions, numeric-recast report, AGG
## reducer check, extreme half-lives) is COMPUTED but no longer exported as a
## separate workbook: SM_2_encoding_audit.xlsx is not written.  Set QUIET <- FALSE
## to print the audit to the console.
if (!QUIET) {
  print(num_report, row.names = FALSE)
  print(agg_check,  row.names = FALSE)
  print(extreme_hl, row.names = FALSE)
}

## ===========================================================================
## 12b.  RUN INTEGRITY CENSUS  (always printed, QUIET or not)
## ===========================================================================
## Warnings are no longer globally suppressed (Section 0), so a coercion failure
## now surfaces as it happens.  This block is the second half of that policy: it
## states, unconditionally and in one place, whether anything was silently lost.
## A clean run prints four zeros.  A non-zero count is a defect to be fixed in
## SM_2.xlsx before the outputs are used.
integrity <- data.frame(
  check = c("Non-empty cells in the 16 analysis columns that failed to parse",
            "Compounds present in one harmonised sheet but not the other",
            "Reference keys with no provenance classification",
            "Half-life records dropped by the zero policy"),
  n = c(
    sum(vapply(NUM$report, function(r) r$n_unparsable, numeric(1))),
    length(union(setdiff(prop_cmpd$Compound, fate_cmpd$Compound),
                 setdiff(fate_cmpd$Compound, prop_cmpd$Compound))),
    if (exists("unclassified_refs")) length(unclassified_refs) else 0L,
    if (exists("N_ZERO_DROPPED")) N_ZERO_DROPPED else NA_integer_),
  stringsAsFactors = FALSE)
## NOTE ON GRAPHICS WARNINGS.  With warnings now visible, some devices emit
## "conversion failure ... mbcsToSbcs" while drawing the figure captions, which
## contain the multiplication sign and the >= glyph.  These are font/locale
## messages from the PDF device and affect the rendering of caption glyphs only;
## they touch no datum.  Run in a UTF-8 locale to silence them.
base::message("Run integrity:")
for (i in seq_len(nrow(integrity)))
  base::message(sprintf("   %-62s %s", integrity$check[i],
                        ifelse(is.na(integrity$n[i]), "n/a", integrity$n[i])))
if (isTRUE(integrity$n[1] > 0))
  base::message("   *** unparsable cells present - see the numeric-recast report ",
                "(set QUIET <- FALSE) before using these outputs ***")

base::message("Analysis complete.")

STAGE("DONE  all sections completed")
