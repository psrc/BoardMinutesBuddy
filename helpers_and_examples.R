library(data.table)
library(officer)
library(stringr)

# Helper Functions ----

#' Extract clean text from Word document using officer package
extract_word_text <- function(filepath) {
  if (!file.exists(filepath)) {
    warning(paste("File not found:", filepath))
    return("")
  }

  tryCatch({
    doc <- read_docx(filepath)
    # Extract all text content, preserving basic structure
    content <- docx_summary(doc)
    # Focus on paragraph content, preserve headings
    text_content <- content[content$content_type == "paragraph", ]

    # Combine text while preserving structure
    if (nrow(text_content) > 0) {
      # Add extra spacing for style breaks that might indicate headings
      combined_text <- paste(text_content$text, collapse = "\n")
      # Clean up excessive whitespace
      combined_text <- str_replace_all(combined_text, "\\s+", " ")
      combined_text <- str_replace_all(combined_text, "\\n\\s*\\n", "\n\n")
      return(str_trim(combined_text))
    }
    return("")
  }, error = function(e) {
    warning(paste("Error reading Word doc:", filepath, "-", e$message))
    return("")
  })
}

#' Extract and clean transcript from VTT file
extract_vtt_text <- function(filepath) {
  if (!file.exists(filepath)) {
    warning(paste("File not found:", filepath))
    return("")
  }

  tryCatch({
    lines <- readLines(filepath, warn = FALSE)

    # Remove VTT header and timestamp lines
    # VTT format: WEBVTT, timestamps (00:00:00.000 --> 00:00:00.000), then speaker: text
    content_lines <- lines[!grepl("^WEBVTT|^NOTE|^\\d{2}:\\d{2}:\\d{2}\\.\\d{3}", lines)]

    # Combine speaker attributions and dialogue
    transcript_text <- paste(content_lines, collapse = " ")

    # Clean up excessive whitespace
    transcript_text <- str_replace_all(transcript_text, "\\s+", " ")
    return(str_trim(transcript_text))
  }, error = function(e) {
    warning(paste("Error reading VTT file:", filepath, "-", e$message))
    return("")
  })
}

# Data Preparation ----

#' Process all file trios and create training dataset
prepare_example_trios <- function(file_trios_dt, num_examples = 3) {
  cat("Processing", nrow(file_trios_dt), "file trios...\n")

  # First pass: calculate combined lengths for all valid examples
  valid_examples <- list()
  failed_count <- 0

  for (i in 1:nrow(file_trios_dt)) {
    row <- file_trios_dt[i]
    cat(sprintf("Processing %d/%d: %s %s %d-%s\n",
                i, nrow(file_trios_dt), row$board, row$mtg_id, row$mtg_year, row$mtg_month))

    # Extract text from each file type
    agenda_text <- extract_word_text(row$agenda)
    minutes_text <- extract_word_text(row$minutes)
    transcript_text <- extract_vtt_text(row$transcript)

    # Validate that all files were processed successfully
    if (nchar(agenda_text) < 50 || nchar(minutes_text) < 50 || nchar(transcript_text) < 100) {
      warning(sprintf("Insufficient content for meeting %s (agenda: %d, minutes: %d, transcript: %d chars)",
                      row$mtg_id, nchar(agenda_text), nchar(minutes_text), nchar(transcript_text)))
      failed_count <- failed_count + 1
      next
    }

    # Calculate combined length and store with texts
    combined_length <- nchar(agenda_text) + nchar(minutes_text)

    valid_examples[[length(valid_examples) + 1]] <- list(
      agenda_text = agenda_text,
      minutes_text = minutes_text,
      transcript_text = transcript_text,
      combined_length = combined_length,
      row_info = row
    )
  }

  cat(sprintf("Found %d valid examples, %d failed\n", length(valid_examples), failed_count))

  # Sort by combined length and take top 50
  if (length(valid_examples) > 0) {
    # Sort by combined_length (ascending)
    valid_examples <- valid_examples[order(sapply(valid_examples, function(x) x$combined_length))]

    # Take only the first 50 (shortest combined lengths)
    n_to_keep <- min(num_examples, length(valid_examples))
    valid_examples <- valid_examples[1:n_to_keep]

    cat(sprintf("Selected %d examples with shortest combined lengths\n", n_to_keep))

    # Create training examples from the selected data
    training_examples <- list()
    for (example_data in valid_examples) {
      example <- create_training_example(
        example_data$agenda_text,
        example_data$transcript_text,
        example_data$minutes_text
      )
      training_examples[[length(training_examples) + 1]] <- example
    }

    cat(sprintf("Successfully created %d training examples\n", length(training_examples)))
    return(training_examples)
  } else {
    cat("No valid examples found\n")
    return(list())
  }
}
