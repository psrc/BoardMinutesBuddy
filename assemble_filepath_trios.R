library(data.table)
library(stringr)
library(lubridate)

# Load and clean filepaths
filepaths <- fread("transcript_filepaths.txt", sep = ",")
filepaths[, filepath := gsub("\\\\", "/", filepath)]

# Extract metadata
rgx_board <- "(OC|EB|GMPB|TPB|EDD)"
filepaths <- filepaths[grepl(rgx_board, filepath)]
filepaths[, `:=`(
  board = str_extract(filepath, rgx_board),
  filename = basename(filepath),
  filetype = fifelse(str_detect(filepath, "\\.vtt$"), "transcript", NA_character_)
)]

filepaths[str_detect(filepath, "agenda"), filetype := "agenda"]
filepaths[str_detect(filepath, "minutes"), filetype := "minutes"]
filepaths[, filepath := str_remove(filepath, paste0(filename, "$"))]

# Text month labels & abbreviations to match against
abbr_mo <- c(month.abb, "June", "July", "Sept", month.name) %>%
  c(tolower(.)) %>% paste0(collapse="|") %>% paste0("(", ., ")")

# Generate pattern groups and associated date parsers
date_patterns <- list(
  # Begins with formats using text month
  ymd_text = list(
    patterns = c(
      paste0("(?<!\\d)20\\d{2}", abbr_mo, "(0?[1-9]|[12][0-9]|3[01])(?!\\d)"),
      paste0("(?<!\\d)20\\d{2}[-_ .]?", abbr_mo, "[-_ .]?(0?[1-9]|[12][0-9]|3[01])(?!\\d)")
    ),
    parser = ymd
  ),
  ydm_text = list(
    patterns = c(
      paste0("(?<!\\d)20\\d{2}(0?[1-9]|[12][0-9]|3[01])", abbr_mo),
      paste0("(?<!\\d)20\\d{2}[-_ .]?(0?[1-9]|[12][0-9]|3[01])[-_ .]?", abbr_mo)
    ),
    parser = ydm
  ),
  mdy_text = list(
    patterns = c(
      paste0(abbr_mo, "(0?[1-9]|[12][0-9]|3[01])(20)?\\d{2}(?!\\d)"),
      paste0(abbr_mo, "[-_ .]?(0?[1-9]|[12][0-9]|3[01])[-_ .]?(20)?\\d{2}(?!\\d)")
    ),
    parser = mdy
  ),
  dmy_text = list(
    patterns = c(
      paste0("(?<!\\d)(0?[1-9]|[12][0-9]|3[01])", abbr_mo, "(20)?\\d{2}(?!\\d)"),
      paste0("(?<!\\d)(0?[1-9]|[12][0-9]|3[01])[-_ .]?", abbr_mo, "[-_ .]?(20)?\\d{2}(?!\\d)")
    ),
    parser = dmy
  ),
  # Then looks for formats with numeric month
  ymd_num = list(
    patterns = c(
      "(?<!\\d)20\\d{2}(1[0-2]|0?[1-9])(0?[1-9]|[12][0-9]|3[01])(?!\\d)",
      "(?<!\\d)20\\d{2}[-_ .](1[0-2]|0[1-9])[-_ .]?(0[1-9]|[12][0-9]|3[01])(?!\\d)"
    ),
    parser = ymd
  ),
  ydm_num = list(
    patterns = c(
      "(?<!\\d)20\\d{2}(0?[1-9]|[12][0-9]|3[01])(1[0-2]|0?[1-9])(?!\\d)",
      "(?<!\\d)20\\d{2}[-_ .](0?[1-9]|[12][0-9]|3[01])[-_ .](1[0-2]|0?[1-9])(?!\\d)"
    ),
    parser = ydm
  ),
  mdy_num = list(
    patterns = c(
      "(?<!\\d)(1[0-2]|0?[1-9])(0?[1-9]|[12][0-9]|3[01])(20)?\\d{2}(?!\\d)",
      "(?<!\\d)(1[0-2]|0?[1-9])[-_ \\.](0?[1-9]|[12][0-9]|3[01])[-_ \\.](20)?\\d{2}(?!\\d)"
    ),
    parser = mdy
  ),
  dmy_num = list(
    patterns = c(
      "(?<!\\d)(0?[1-9]|[12][0-9]|3[01])(1[0-2]|0?[1-9])(20)?\\d{2}(?!\\d)",
      "(?<!\\d)(0?[1-9]|[12][0-9]|3[01])[-_ \\.](1[0-2]|0?[1-9])[-_ \\.](20)?\\d{2}(?!\\d)"
    ),
    parser = dmy
  )
)

filepaths[, c("date_chr", "mtg_date") := .(NA_character_, as.Date(NA))]

for (dp in date_patterns) {
  for (pat in dp$patterns) {
    filepaths[is.na(date_chr) & str_detect(filename, pat), date_chr := str_extract(filename, pat)]
  }
  filepaths[!is.na(date_chr) & is.na(mtg_date), mtg_date := dp$parser(date_chr)]
}

# If no full date exists, extract year and month separately
rgx_my <- "(?<!\\d)(1[0-2]|0?[1-9])[-_ .]?20\\d{2}(?!\\d)"
rgx_y <- "(?<!\\d)20\\d{2}(?!\\d)"

filepaths[is.na(date_chr) & str_detect(filepath, rgx_my), date_chr := str_extract(filepath, rgx_my)]
filepaths[!is.na(date_chr) & is.na(mtg_date), mtg_date := my(date_chr)]
filepaths[is.na(mtg_date) & str_detect(filename, rgx_y), mtg_year := str_extract(filename, rgx_y)]
filepaths[is.na(mtg_date) & str_detect(filename, abbr_mo), mtg_month := str_extract(filename, abbr_mo)]
filepaths[is.na(mtg_date) & str_detect(filepath, rgx_y), mtg_year := str_extract(filepath, rgx_y)]
filepaths[is.na(mtg_date) & str_detect(filepath, abbr_mo), mtg_month := str_extract(filepath, abbr_mo)]
filepaths[!is.na(mtg_year) & !is.na(mtg_month) & is.na(mtg_date), mtg_date := ym(paste(mtg_year, mtg_month, sep = "-"))]

# Finalize columns
filepaths[!is.na(mtg_date), `:=`(
  mtg_year = year(mtg_date),
  mtg_month = month(mtg_date, label = TRUE),
  fullpath = paste0(filepath, filename)
)]

# Save remainders and process valid ones
date_found_no <- filepaths[is.na(mtg_year) | is.na(mtg_month)]
date_found_yes <- filepaths[!is.na(mtg_year) & !is.na(mtg_month)]

# Pivot for matching agenda-minutes-transcript trios
file_trios <- dcast(
  date_found_yes,
  board + mtg_year + mtg_month ~ filetype,
  value.var = "fullpath",
  fun.aggregate = function(x) {
    # Where multiple files exist for a date/board combo,
    #   prioritize those with full dates, i.e. day provided
    idx <- match(x, date_found_yes$fullpath)
    valid_idx <- idx[!is.na(date_found_yes$date_chr[idx])]

    # If no valid date_chr, select the shortest filename
    if(length(valid_idx) == 0) valid_idx <- idx
    filenames <- date_found_yes$filename[valid_idx]
    min_idx <- valid_idx[which.min(nchar(filenames))]
    shortest_idx <- date_found_yes$fullpath[min_idx]

    # Tiebreaker - first available
    selected_idx <- shortest_idx[1]

    # Return the corresponding fullpath
    date_found_yes$fullpath[selected_idx]
  })

# Separate missing sets
no_transcript        <- file_trios[is.na(transcript)]
no_agenda_or_minutes <- file_trios[(is.na(agenda) | is.na(minutes)) & !is.na(transcript)]
file_trios           <- file_trios[!is.na(agenda) & !is.na(minutes) & !is.na(transcript)]
