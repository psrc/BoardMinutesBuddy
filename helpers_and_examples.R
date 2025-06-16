library(data.table)
library(officer)
library(stringr)
library(purrr)
library(httr2)
library(jsonlite)

# Text conversion to and from markdown ----------

#' Extract clean text from Word document using officer package
extract_word_text <- function(filepath) {
  if (!file.exists(filepath)) {
    warning(paste("File not found:", filepath))
    return("")
  }

  tryCatch({
    doc <- read_docx(filepath)
    content <- docx_summary(doc)
    text_content <- content[content$content_type == "paragraph", ]

    if (nrow(text_content) > 0) {
      combined_text <- paste(text_content$text, collapse = "\n")
      # Only replace multiple spaces/tabs, but preserve newlines
      combined_text <- str_replace_all(combined_text, "[ \\t]+", " ")
      # Clean up excessive newlines (3+ becomes 2)
      combined_text <- str_replace_all(combined_text, "\\n{3,}", "\n\n")
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

# Extract Word document to markdown with proper formatting detection
extract_word_to_markdown <- function(filepath) {
  filepath <- normalizePath(filepath, winslash = "/", mustWork = FALSE)

  if (!file.exists(filepath)) {
    warning(paste("File not found:", filepath))
    return("")
  }

  tryCatch({
    doc <- read_docx(filepath)
    content <- docx_summary(doc)

    markdown_lines <- c()

    for (i in 1:nrow(content)) {
      row <- content[i, ]

      if (row$content_type == "paragraph") {
        text <- row$text

        # Skip empty paragraphs
        if (is.na(text) || str_trim(text) == "") {
          markdown_lines <- c(markdown_lines, "")
          next
        }

        # Handle headings based on style name
        if (!is.na(row$style_name) && grepl("heading", row$style_name, ignore.case = TRUE)) {
          heading_match <- str_extract(row$style_name, "\\d+")
          level <- if (!is.na(heading_match)) min(as.numeric(heading_match), 6) else 1
          markdown_lines <- c(markdown_lines, paste0(strrep("#", level), " ", text))

        } else {
          # Handle regular paragraphs with formatting
          formatted_text <- text

          # Detect block quotes (like ACTION items) - check if paragraph has special formatting
          is_block_quote <- (!is.na(row$style_name) &&
                               (grepl("quote|block", row$style_name, ignore.case = TRUE) ||
                                  row$style_name == "List Paragraph")) ||
            grepl("^(ACTION|MOTION):", text, ignore.case = TRUE)

          # Apply bold formatting if entire paragraph is bold
          if (!is.na(row$is_bold) && isTRUE(row$is_bold)) {
            formatted_text <- paste0("**", formatted_text, "**")
          }

          # Apply italic formatting if entire paragraph is italic
          if (!is.na(row$is_italic) && isTRUE(row$is_italic)) {
            formatted_text <- paste0("*", formatted_text, "*")
          }

          # Handle block quotes
          if (is_block_quote) {
            formatted_text <- paste0("> ", formatted_text)
          }

          # Handle indentation based on paragraph level
          if (!is.na(row$level) && row$level > 0 && !is_block_quote) {
            indent <- strrep("  ", row$level)
            formatted_text <- paste0(indent, formatted_text)
          }

          markdown_lines <- c(markdown_lines, formatted_text)
        }
      }
    }

    # Clean up and return
    combined_text <- paste(markdown_lines, collapse = "\n")
    combined_text <- str_replace_all(combined_text, "\\n{3,}", "\n\n")
    return(str_trim(combined_text))

  }, error = function(e) {
    warning(paste("Error reading Word doc:", filepath, "-", e$message))
    return("")
  })
}

# Convert markdown to Word using template styles
markdown_to_word <- function(markdown_text, template_filepath, output_filepath) {
  template_filepath <- normalizePath(template_filepath, winslash = "/", mustWork = FALSE)

  if (!file.exists(template_filepath)) {
    stop(paste("Template file not found:", template_filepath))
  }

  # Handle output filepath
  if (output_filepath == "." || dir.exists(output_filepath)) {
    output_dir <- if (output_filepath == ".") getwd() else output_filepath
    output_filepath <- file.path(output_dir, paste0("converted_document_",
                                                    format(Sys.time(), "%Y%m%d_%H%M%S"),
                                                    ".docx"))
  } else if (!str_detect(output_filepath, "\\.docx$")) {
    output_filepath <- paste0(output_filepath, ".docx")
  }

  output_filepath <- normalizePath(output_filepath, winslash = "/", mustWork = FALSE)

  tryCatch({
    # Read template and get available styles
    doc <- read_docx(template_filepath)
    template_styles <- styles_info(doc)

    # Split markdown into lines
    lines <- strsplit(markdown_text, "\n")[[1]]

    for (line in lines) {
      # Handle empty lines
      if (str_trim(line) == "") {
        doc <- body_add_par(doc, "")
        next
      }

      # Detect indentation level from leading spaces
      indent_match <- str_match(line, "^(\\s*)(.*)")
      leading_spaces <- ifelse(is.na(indent_match[2]), "", indent_match[2])
      content <- ifelse(is.na(indent_match[3]), line, indent_match[3])
      indent_level <- nchar(leading_spaces) / 2

      # Handle headings
      if (str_detect(content, "^#{1,6}\\s+")) {
        heading_match <- str_match(content, "^(#{1,6})\\s+(.*)")
        level <- nchar(heading_match[2])
        heading_text <- heading_match[3]

        # Use exact style names from your template
        style_map <- c(
          "1" = "heading 1",
          "2" = "heading 2",
          "3" = "heading 3"
        )

        style_to_use <- if (as.character(level) %in% names(style_map)) {
          style_map[[as.character(level)]]
        } else {
          "Normal"
        }

        doc <- body_add_par(doc, heading_text, style = style_to_use)

        # Handle block quotes (ACTION items)
      } else if (str_detect(content, "^>\\s+")) {
        quote_text <- str_replace(content, "^>\\s+", "")

        # Create block quote formatting - bold text in a distinct paragraph
        if (str_detect(quote_text, "^\\*\\*.*\\*\\*$")) {
          # Remove markdown bold markers and apply actual bold formatting
          clean_text <- str_replace_all(quote_text, "\\*\\*", "")
          doc <- body_add_fpar(doc, fpar(ftext(clean_text, prop = fp_text(bold = TRUE))))
        } else {
          # Process any inline formatting in the quote
          doc <- add_formatted_paragraph(doc, quote_text, 0, template_styles, is_quote = TRUE)
        }

      } else {
        # Handle regular paragraphs with formatting and indentation
        doc <- add_formatted_paragraph(doc, content, indent_level, template_styles)
      }
    }

    print(doc, target = output_filepath)
    message(paste("Document saved to:", output_filepath))

  }, error = function(e) {
    stop(paste("Error converting markdown to Word:", e$message))
  })
}

# Helper function to add formatted paragraphs
add_formatted_paragraph <- function(doc, text, indent_level, template_styles, is_quote = FALSE) {
  # Handle simple cases first
  if (!str_detect(text, "\\*\\*|\\*")) {
    # Plain text
    if (is_quote) {
      # For quotes, use bold formatting and possibly different alignment
      doc <- body_add_fpar(doc, fpar(ftext(text, prop = fp_text(bold = TRUE))))
    } else if (indent_level > 0) {
      # Use List Paragraph style for indented content
      doc <- body_add_par(doc, text, style = "List Paragraph")
    } else {
      doc <- body_add_par(doc, text, style = "Normal")
    }
    return(doc)
  }

  # Handle formatted text
  parts <- parse_formatting(text)

  # Create ftext objects for each part
  ftext_objects <- lapply(parts, function(part) {
    # For quotes, make everything bold
    if (is_quote) {
      ftext(part$text, prop = fp_text(bold = TRUE))
    } else if (part$bold && part$italic) {
      ftext(part$text, prop = fp_text(bold = TRUE, italic = TRUE))
    } else if (part$bold) {
      ftext(part$text, prop = fp_text(bold = TRUE))
    } else if (part$italic) {
      ftext(part$text, prop = fp_text(italic = TRUE))
    } else {
      ftext(part$text)
    }
  })

  # Create paragraph properties
  para_props <- if (indent_level > 0 && !is_quote) {
    fp_par(padding.left = indent_level * 0.5)
  } else {
    fp_par()
  }

  # Add formatted paragraph
  doc <- body_add_fpar(doc, do.call(fpar, c(ftext_objects, list(fp_p = para_props))))

  return(doc)
}

# Helper function to parse markdown formatting
parse_formatting <- function(text) {
  parts <- list()
  pos <- 1

  # Simple regex approach for basic bold/italic
  while (pos <= nchar(text)) {
    # Look for **bold**
    bold_match <- str_locate(substr(text, pos, nchar(text)), "\\*\\*[^*]+\\*\\*")
    # Look for *italic* (but not **bold**)
    italic_match <- str_locate(substr(text, pos, nchar(text)), "(?<!\\*)\\*[^*]+\\*(?!\\*)")

    if (!is.na(bold_match[1])) {
      # Add text before bold
      if (bold_match[1] > 1) {
        before_text <- substr(text, pos, pos + bold_match[1] - 2)
        if (before_text != "") {
          parts[[length(parts) + 1]] <- list(text = before_text, bold = FALSE, italic = FALSE)
        }
      }

      # Add bold text
      bold_text <- substr(text, pos + bold_match[1] + 1, pos + bold_match[2] - 2)
      parts[[length(parts) + 1]] <- list(text = bold_text, bold = TRUE, italic = FALSE)

      pos <- pos + bold_match[2]

    } else if (!is.na(italic_match[1])) {
      # Add text before italic
      if (italic_match[1] > 1) {
        before_text <- substr(text, pos, pos + italic_match[1] - 2)
        if (before_text != "") {
          parts[[length(parts) + 1]] <- list(text = before_text, bold = FALSE, italic = FALSE)
        }
      }

      # Add italic text
      italic_text <- substr(text, pos + italic_match[1], pos + italic_match[2] - 2)
      parts[[length(parts) + 1]] <- list(text = italic_text, bold = FALSE, italic = TRUE)

      pos <- pos + italic_match[2]

    } else {
      # No more formatting, add remaining text
      remaining <- substr(text, pos, nchar(text))
      if (remaining != "") {
        parts[[length(parts) + 1]] <- list(text = remaining, bold = FALSE, italic = FALSE)
      }
      break
    }
  }

  # If no parts were found, return the original text
  if (length(parts) == 0) {
    parts[[1]] <- list(text = text, bold = FALSE, italic = FALSE)
  }

  return(parts)
}


# Data Preparation ------------------------------

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
    agenda_text <- extract_word_with_styles(row$agenda) #extract_word_text(row$agenda)
    minutes_text <- extract_word_with_styles(row$minutes) #extract_word_text(row$minutes)
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

convert_examples_to_jsonl <- function(example_trios, output_path="./example.jsonl") {
  if (length(example_trios) == 0) {
    stop("No examples provided in example_trios")
  }

  # Validate input structure
  for (i in seq_along(example_trios)) {
    example <- example_trios[[i]]
    if (!("messages" %in% names(example))) {
      stop("Invalid example structure at index ", i)
    }
  }

  # Convert each example to JSON and write to file
  jsonl_lines <- map_chr(example_trios, ~ toJSON(.x, auto_unbox = TRUE))

  # Write to JSONL file
  writeLines(jsonl_lines, output_path)

  cat("Converted", length(example_trios), "examples to JSONL format at:", output_path, "\n")

  return(invisible(output_path))
}
