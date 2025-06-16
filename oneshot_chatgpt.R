# OpenAI API One-shot Script with JSONL Example Loading
# Purpose: Use single example with full gpt-4.1 model

library(data.table)
library(httr2)
library(jsonlite)
library(officer)
library(stringr)
library(purrr)

gpt_model <- "gpt-4.1"

# Load examples from JSONL file ----
load_examples_from_jsonl <- function(jsonl_path) {
  if (!file.exists(jsonl_path)) {
    stop("JSONL file not found: ", jsonl_path)
  }

  lines <- readLines(jsonl_path, warn = FALSE)
  examples <- map(lines, ~ fromJSON(.x, simplifyVector = FALSE))

  # Validate that each example has the expected structure
  for (i in seq_along(examples)) {
    example <- examples[[i]]
    if (!("messages" %in% names(example)) || length(example$messages) < 3) {
      stop("Invalid example structure at line ", i, ": expected 'messages' with at least 3 entries")
    }
  }

  return(examples)
}

# Create single-example messages ----
create_oneshot_chatgpt <- function(agenda_text, transcript_text, example) {

  # System message
  system_msg <- list(role = "system",
                     content = paste("You are an expert at creating structured meeting minutes for public agency board meetings.",
                                     "Given an agenda and meeting transcript, you produce concise, professional minutes that follow",
                                     "the agenda structure, summarize key discussions without quotes, and highlight decisions and action items."))

  # Extract user and assistant content from the example
  user_content <- example$messages[[2]]$content
  assistant_content <- example$messages[[3]]$content

  # Create the message chain
  messages <- list(
    system_msg,
    list(role = "user", content = user_content),
    list(role = "assistant", content = assistant_content),
    list(role = "user", content = sprintf("AGENDA:\n%s\n\nTRANSCRIPT:\n%s", agenda_text, transcript_text))
  )

  return(messages)
}

# Generate minutes with single example ----
generate_minutes_chatgpt <- function(agenda_text, transcript_text, examples_jsonl_path,
                                     model_id = gpt_model, example_index = NULL) {

  # Load examples from JSONL file
  examples <- load_examples_from_jsonl(examples_jsonl_path)

  # Select example (random if not specified)
  if (is.null(example_index)) {
    selected_example <- sample(examples, 1)[[1]]
  } else {
    if (example_index > length(examples) || example_index < 1) {
      stop("Example index out of range: ", example_index, " (available: 1-", length(examples), ")")
    }
    selected_example <- examples[[example_index]]
  }

  # Create messages with the selected example
  messages <- create_oneshot_chatgpt(agenda_text, transcript_text, selected_example)

  message_data <- list(
    model = model_id,
    max_tokens = 6500,
    temperature = 0.3,
    messages = messages
  )

  response <- request(paste0(BASE_URL, "/chat/completions")) |>
    req_method("POST") |>
    req_headers(
      "Authorization" = paste("Bearer", OPENAI_API_KEY),
      "Content-Type" = "application/json"
    ) |>
    req_timeout(300) |>
    req_body_json(message_data) |>
    req_retry(max_tries = 5, backoff = ~ 2 ^ .x) |>
    req_perform()

  if (resp_status(response) != 200) {
    stop("Failed to generate minutes: ", resp_body_string(response))
  }

  result <- resp_body_json(response)
  minutes_rawtext <- result$choices[[1]]$message$content
  #minutes_to_word_template(minutes_rawtext, "minutes_msword.dotx")
  return(minutes_rawtext)
}
