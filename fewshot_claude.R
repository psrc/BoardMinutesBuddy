# Anthropic API Few-shot Script

library(data.table)
library(httr2)
library(purrr)

# Anthropic API configuration
claude_model <- "claude-sonnet-4-0"
ANTHROPIC_BASE_URL <- "https://api.anthropic.com"
ANTHROPIC_API_KEY <- Sys.getenv("ANTHROPIC_API_KEY")

# Create per-example messages ----
create_fewshot_claude <- function(agenda_text,
                                  transcript_text,
                                  example_trios,
                                  model_id = claude_model,
                                  max_tokens = 6500) {

  # System message for Anthropic
  system_msg <- paste("You are an expert at creating structured meeting minutes for public agency board meetings.",
                      "Given an agenda and meeting transcript, you produce concise, professional minutes that follow",
                      "the agenda structure, summarize key discussions without quotes, and highlight decisions and action items.")

  # Process examples for Anthropic format
  messages_list <- list()

  for (i in seq_along(example_trios)) {
    example <- example_trios[[i]]

    # Extract user and assistant content from the example
    user_content <- example$messages[[2]]$content
    assistant_content <- example$messages[[3]]$content

    # Store as user/assistant pair
    messages_list[[i]] <- list(
      user = user_content,
      assistant = assistant_content
    )
  }

  # Create the final task
  final_task <- sprintf("AGENDA:\n%s\n\nTRANSCRIPT:\n%s", agenda_text, transcript_text)

  return(list(
    system = system_msg,
    examples = messages_list,
    final_task = final_task
  ))
}

# Send prompt in chunks ----
generate_minutes_claude <- function(agenda_path, transcript_path, example_trios,
                                     model_id = claude_model, max_examples = 3) {

  agenda_text <- extract_word_text(agenda_path)
  transcript_text <- extract_vtt_text(transcript_path)

  prompt_data <- create_fewshot_claude(agenda_text, transcript_text, example_trios, model_id)

  # Randomly select up to N examples
  selected_examples <- sample(prompt_data$examples, min(max_examples, length(prompt_data$examples)))

  # Build messages array for Anthropic API
  messages <- list()

  # Add few-shot examples as alternating user/assistant messages
  for (ex in selected_examples) {
    messages <- append(messages, list(list(role = "user", content = ex$user)))
    messages <- append(messages, list(list(role = "assistant", content = ex$assistant)))
  }

  # Add the final user task
  messages <- append(messages, list(list(role = "user", content = prompt_data$final_task)))

  # Prepare request body for Anthropic API
  request_body <- list(
    model = model_id,
    max_tokens = max_tokens,
    temperature = 0.3,
    system = prompt_data$system,
    messages = messages
  )

  # Make API request to Anthropic
  response <- request(paste0(ANTHROPIC_BASE_URL, "/v1/messages")) |>
    req_method("POST") |>
    req_headers(
      "x-api-key" = ANTHROPIC_API_KEY,
      "Content-Type" = "application/json",
      "anthropic-version" = "2023-06-01"
    ) |>
    req_timeout(300) |>
    req_body_json(request_body) |>
    req_retry(max_tries = 5, backoff = ~ 2 ^ .x) |>
    req_perform()

  # Check response status
  if (resp_status(response) != 200) {
    stop("Failed to generate minutes: ", resp_body_string(response))
  }

  # Parse response
  result <- resp_body_json(response)

  # Extract content from Anthropic response format
  return(result$content[[1]]$text)
}

# Usage example:
# minutes <- generate_minutes_anthropic(agenda_path, transcript_path, example_trios)
