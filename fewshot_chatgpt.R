# OpenAI API Few-shot Script with Prompt Chunking
# Purpose: Split few-shot examples to avoid exceeding token limits

library(data.table)
library(httr2)
library(jsonlite)
library(officer)
library(stringr)
library(purrr)

gpt_model <- "gpt-4.1-mini-2025-04-14"

# Create per-example messages ----
create_fewshot_chatgpt <- function(agenda_path, transcript_path, example_trios) {

  agenda_text <- extract_word_text(agenda_path)
  transcript_text <- extract_vtt_text(transcript_path)

  # System message
  system_msg <- list(role = "system",
                     content = paste0("You are an expert at creating structured meeting minutes for public agency board meetings.",
                                      "Given an agenda and meeting transcript, you produce concise, professional minutes that follow",
                                      "the agenda structure, summarize key discussions without quotes, and highlight decisions and action items."))

  # Chunk user examples by one example at a time
  messages_list <- list()

  for (i in seq_along(example_trios)) {
    example <- example_trios[[i]]
    user_content <- example$messages[[2]]$content
    assistant_content <- example$messages[[3]]$content

    messages <- list(
      system_msg,
      list(role = "user", content = user_content),
      list(role = "assistant", content = assistant_content)
    )

    messages_list[[i]] <- messages
  }

  # Add the new input at the end
  final_task <- list(
    role = "user",
    content = sprintf("AGENDA:\n%s\n\nTRANSCRIPT:\n%s", agenda_text, transcript_text)
  )

  return(list(system = system_msg, examples = messages_list, final_task = final_task))
}

# Send prompt in chunks ----
generate_minutes_chatgpt <- function(agenda_text, transcript_text, example_trios,
                                     model_id = gpt_model, max_examples = 3) {

  prompt_data <- create_fewshot_chatgpt(agenda_text, transcript_text, example_trios)

  # Randomly select up to N examples
  selected_examples <- sample(prompt_data$examples, min(max_examples, length(prompt_data$examples)))

  # Flatten message chain: system + selected examples + final user task
  combined_messages <- list(prompt_data$system)
  for (ex in selected_examples) {
    combined_messages <- c(combined_messages, ex[2:3])  # only user/assistant from example
  }
  combined_messages <- c(combined_messages, list(prompt_data$final_task))

  message_data <- list(
    model = model_id,
    max_tokens = 6500,
    temperature = 0.3,
    messages = combined_messages
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
  return(result$choices[[1]]$message$content)
}
