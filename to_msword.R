library(officer)
library(stringr)

dotx_path <- "minutes_msword.dotx"
title_pattern <- paste0("MINUTES$|(BOARD|COMMITTEE) MEETING$|",
                        "^\\w{3,5}DAY|^PUGET SOUND REGIONAL|\\d{1,2}:\\d{2}")

# Helper functions ----------------------------------------

# Function to extract month abbreviation from minutes text
extract_month_abbrev <- function(minutes_text) {
  first_line <- toupper(strsplit(minutes_text, "\n")[[1]][1])
  month_mapping <- setNames(tolower(month.abb), toupper(month.name))

  # Find which month appears in the first line
  for (month_name in names(month_mapping)) {
    if (grepl(month_name, first_line)) {
      return(month_mapping[month_name])
    }
  }

  # Fallback: return "unk" if no month found
  return("unk")
}

# Safely add paragraph with style fallback
safe_add_par <- function(doc, text, preferred_style, fallback_style = "Normal") {
  # Get available styles in the document
  available_styles <- doc$styles$style_name

  # Use preferred style if available, otherwise use fallback
  style_to_use <- if (preferred_style %in% available_styles) {
    preferred_style
  } else {
    fallback_style
  }

  return(body_add_par(doc, text, style = style_to_use))
}

# Text to MS Word conversion ------------------------------

# Convert minutes to MS Word document via template
minutes_to_word_template <- function(minutes_text,
                                     output_dir = ".",
                                     template_file = dotx_path) {

  # Extract month abbreviation and create filename
  month_abbrev <- extract_month_abbrev(minutes_text)
  filename <- paste0("03a-minutes-", month_abbrev, ".docx")

  # Create full output path
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
  }

  output_path <- file.path(output_dir, filename)

  # Create or load Word document
  if (!is.null(template_file) && file.exists(template_file)) {
    doc <- read_docx(template_file)
  } else {
    doc <- read_docx()
  }

  # Split and clean text
  lines <- strsplit(minutes_text, "\n")[[1]]
  lines <- trimws(lines)
  lines <- lines[lines != ""]

  # Helper functions (same as above)
  is_main_heading <- function(line) {
    grepl("^\\d+[a-z]?\\.", line) &&
      (grepl("Call to Order|Communications|Consent|(Action|Discussion|Information) Item", line) ||
         nchar(line) < 100)
  }

  is_sub_heading <- function(line) {
    grepl("^[a-z]\\.", line) && nchar(line) < 100
  }

  is_action_line <- function(line) {
    grepl("^Action:|It was moved", line)
  }

  # Process title section
  title_lines <- c()
  i <- 1
  while (i <= length(lines) && !is_main_heading(lines[i])) {
    title_lines <- c(title_lines, lines[i])
    i <- i + 1
  }

  # Format title section with custom styling
  for (line in title_lines) {
    if (grepl("MINUTES", line)) {
      doc <- doc %>%
        safe_add_par(line, "Title")
    } else if (grepl(title_pattern, line)) {
      doc <- doc %>%
        safe_add_par(line, "Subtitle", "heading 1")
    } else if (grepl("\\d+:\\d+", line)) {  # Time
      doc <- doc %>%
        safe_add_par(line, "Subtitle", "heading 2")
    }
  }

  doc <- doc %>% body_add_par("", style = "Normal")  # Add space

  # Process main content
  current_line <- i

  while (current_line <= length(lines)) {
    line <- lines[current_line]

    if (is_main_heading(line)) {
      doc <- doc %>%
        body_add_par("", style = "Normal") %>%  # Space before heading
        safe_add_par(line, "heading 2")

    } else if (is_sub_heading(line)) {
      doc <- doc %>%
        safe_add_par(line, "heading 3")

    } else if (is_action_line(line)) {
      doc <- doc %>%
        body_add_par("", style = "Normal") %>%
        safe_add_par(line, "List Paragraph", "Normal") %>%
        body_add_par("", style = "Normal")

    } else if (grepl("^Guests and staff present|^Attachments:", line)) {
      # Special formatting for certain sections
      doc <- doc %>%
        safe_add_par(line, "List Paragraph", "Normal")

    } else {
      doc <- doc %>%
        body_add_par(line, style = "Normal")
    }

    current_line <- current_line + 1
  }

  # Save document
  print(doc, target = output_path)

  cat("Advanced formatted document saved as:", output_path, "\n")
  return(invisible(doc))
}

# Example usage:
# Assuming your minutes_text variable is already loaded:

# Save to specific directory
# minutes_to_word(minutes_text, output_dir = "minutes/2023")

# Conversion with template
# minutes_to_word_template(minutes_text,
#                          output_dir = "formatted_minutes",
#                          template_file = "minutes_template.docx")
