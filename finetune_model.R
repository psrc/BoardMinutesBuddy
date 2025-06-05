# OpenAI API Fine-tuning Script for Board Meeting Minutes
# Purpose: Train ChatGPT model to generate structured meeting minutes from agendas and transcripts

library(data.table)
library(httr2)
library(jsonlite)
library(purrr)

# Configuration
OPENAI_API_KEY <- Sys.getenv("OPENAI_API_KEY")
if (OPENAI_API_KEY == "") stop("Please set OPENAI_API_KEY environment variable")

BASE_URL <- "https://api.openai.com/v1"
gpt_model <- "gpt-4.1-nano-2025-04-14" # Consider parameterizing e.g. Sys.getenv("BOARD_MINUTES_MODEL")

#' Create training example in OpenAI format
create_training_example <- function(agenda_text, transcript_text, minutes_text) {
  # System message
  system_msg <- paste("You are an expert at creating structured meeting minutes for public agency board meetings.",
                      "Given an agenda and meeting transcript, you produce concise, professional minutes that follow",
                      "the agenda structure, summarize key discussions without quotes, and highlight decisions and action items.")
  # User message with inputs
  user_msg <- sprintf("Please create meeting minutes based on this agenda and transcript.\n\nAGENDA:\n%s\n\nTRANSCRIPT:\n%s",
                      agenda_text, transcript_text)

  # Assistant response (target output)
  assistant_msg <- minutes_text

  # Format for OpenAI fine-tuning (chat completions format)
  list(
    messages = list(
      list(role = "system", content = system_msg),
      list(role = "user", content = user_msg),
      list(role = "assistant", content = assistant_msg)
    )
  )
}

#' Save training data as JSONL file
save_training_jsonl <- function(training_examples, filepath = "training_data.jsonl") {
  # Convert each example to JSON and write line by line
  jsonl_lines <- map_chr(training_examples, ~ toJSON(.x, auto_unbox = TRUE))
  writeLines(jsonl_lines, filepath)
  cat("Training data saved to:", filepath, "\n")
  cat("File size:", file.size(filepath), "bytes\n")
  return(filepath)
}

# OpenAI API Functions ----

#' Upload training file to OpenAI
upload_training_file <- function(jsonl_filepath) {
  cat("Uploading training file to OpenAI...\n")

  response <- request(paste0(BASE_URL, "/files")) |>
    req_method("POST") |>
    req_headers(
      "Authorization" = paste("Bearer", OPENAI_API_KEY)
    ) |>
    req_body_multipart(
      file = curl::form_file(jsonl_filepath),
      purpose = "fine-tune"
    ) |>
    req_retry(max_tries = 5, backoff = ~ 2 ^ .x) |>
    req_perform()

  if (resp_status(response) != 200) {
    stop("Failed to upload file: ", resp_body_string(response))
  }

  result <- tryCatch({
    resp_body_json(response)
  }, error = function(e) {
    warning("Failed to parse JSON response from upload_training_file(): ", e$message)
    cat("Raw response:\n", resp_body_string(response), "\n")
    return(NULL)
  })
  if (is.null(result)) stop("Upload failed due to invalid JSON response.")

  cat("File uploaded successfully. File ID:", result$id, "\n")
  return(result$id)
}

#' Create fine-tuning job
create_finetuning_job <- function(file_id, model = gpt_model) {
  cat("Creating fine-tuning job...\n")

  job_data <- list(
    training_file = file_id,
    model = model,
    suffix = paste0("board_minutes_", format(Sys.Date(), "%Y%m%d")),
    hyperparameters = list(
      n_epochs = "auto",  # Let OpenAI determine optimal epochs
      batch_size = "auto",
      learning_rate_multiplier = "auto"
    )
  )

  response <- request(paste0(BASE_URL, "/fine_tuning/jobs")) |>
    req_method("POST") |>
    req_headers(
      "Authorization" = paste("Bearer", OPENAI_API_KEY),
      "Content-Type" = "application/json"
    ) |>
    req_body_json(job_data) |>
    req_retry(max_tries = 5, backoff = ~ 2 ^ .x) |>
    req_perform()

  if (resp_status(response) != 200) {
    stop("Failed to create fine-tuning job: ", resp_body_string(response))
  }

  result <- tryCatch({
    resp_body_json(response)
  }, error = function(e) {
    warning("Failed to parse JSON in create_finetuning_job(): ", e$message)
    cat("Raw response:\n", resp_body_string(response), "\n")
    return(NULL)
  })
  if (is.null(result)) stop("Job creation failed due to invalid JSON response.")

  cat("Fine-tuning job created. Job ID:", result$id, "\n")
  cat("Status:", result$status, "\n")
  cat("Model:", result$model, "\n")
  return(result)
}

#' Check fine-tuning job status
check_job_status <- function(job_id) {
  response <- request(paste0(BASE_URL, "/fine_tuning/jobs/", job_id)) |>
    req_headers(
      "Authorization" = paste("Bearer", OPENAI_API_KEY)
    ) |>
    req_retry(max_tries = 5, backoff = ~ 2 ^ .x) |>
    req_perform()

  if (resp_status(response) != 200) {
    stop("Failed to check job status in check_job_status(): ", resp_body_string(response))
  }

  result <- tryCatch({
    resp_body_json(response)
  }, error = function(e) {
    warning("Failed to parse JSON response: ", e$message)
    cat("Raw response:\n", resp_body_string(response), "\n")
    return(NULL)
  })
  if (is.null(result)) stop("Check job status failed due to invalid JSON response.")

  cat("Job", job_id, "status:", result$status, "\n")
  if (!is.null(result$fine_tuned_model)) {
    cat("Fine-tuned model ID:", result$fine_tuned_model, "\n")
  }
  if (!is.null(result$trained_tokens)) {
    cat("Trained tokens:", result$trained_tokens, "\n")
  }
  return(result)
}

#' List all fine-tuning jobs
list_finetuning_jobs <- function() {
  response <- request(paste0(BASE_URL, "/fine_tuning/jobs")) |>
    req_method("POST") |>
    req_headers(
      "Authorization" = paste("Bearer", OPENAI_API_KEY)
    ) |>
    req_retry(max_tries = 5, backoff = ~ 2 ^ .x) |>
    req_perform()

  if (resp_status(response) != 200) {
    stop("Failed to list jobs: ", resp_body_string(response))
  }

  result <- tryCatch({
    resp_body_json(response)
  }, error = function(e) {
    warning("Failed to parse JSON response in list_finetuning_jobs(): ", e$message)
    cat("Raw response:\n", resp_body_string(response), "\n")
    return(NULL)
  })
  if (is.null(result)) stop("List finetuning jobs failed due to invalid JSON response.")

  return(result$data)
}

#' Use fine-tuned model to generate minutes
generate_minutes <- function(agenda_text, transcript_text, model_id) {
  system_msg <- paste("You are an expert at creating structured meeting minutes for public agency board meetings.",
                      "Given an agenda and meeting transcript, you produce concise, professional minutes that follow",
                      "the agenda structure, summarize key discussions without quotes, and highlight decisions and action items.")
  user_msg <- sprintf("Please create meeting minutes based on this agenda and transcript.\n\nAGENDA:\n%s\n\nTRANSCRIPT:\n%s",
                      agenda_text, transcript_text)

  message_data <- list(
    model = model_id,
    max_tokens = 4000,
    temperature = 0.3,  # Lower temperature for more consistent output
    messages = list(
      list(role = "system", content = system_msg),
      list(role = "user", content = user_msg)
    )
  )

  response <- request(paste0(BASE_URL, "/chat/completions")) |>
    req_method("POST") |>
    req_headers(
      "Authorization" = paste("Bearer", OPENAI_API_KEY),
      "Content-Type" = "application/json"
    ) |>
    req_body_json(message_data) |>
    req_retry(max_tries = 5, backoff = ~ 2 ^ .x) |>
    req_perform()

  if (resp_status(response) != 200) {
    stop("Failed to generate minutes: ", resp_body_string(response))
  }

  result <- tryCatch({
    resp_body_json(response)
  }, error = function(e) {
    warning("Failed to parse JSON response in generate_minutes(): ", e$message)
    cat("Raw response:\n", resp_body_string(response), "\n")
    return(NULL)
  })
  if (is.null(result)) stop("Minutes generation failed due to invalid JSON response.")

  return(result$choices[[1]]$message$content)
}

# Validation Functions ----

#' Validate training data format
validate_training_data <- function(training_examples) {
  cat("Validating training data format...\n")

  if (length(training_examples) < 10) {
    warning("OpenAI recommends at least 10 training examples, but 50-100+ is better")
  }

  # Check token limits (rough estimate: 1 token ≈ 4 characters)
  token_counts <- map_dbl(training_examples, function(example) {
    total_chars <- sum(nchar(unlist(example$messages)))
    return(total_chars / 4)  # Rough token estimate
  })

  max_tokens <- max(token_counts)
  cat("Estimated max tokens per example:", round(max_tokens), "\n")

  if (max_tokens > 16000) {
    warning("Some examples may exceed token limits. Consider truncating long transcripts.")
  }

  cat("Validation complete. Examples ready for training.\n")
  return(TRUE)
}

# Main Execution Functions ----

#' Complete fine-tuning pipeline
run_finetuning_pipeline <- function(file_trios_dt) {
  cat("=== Starting OpenAI Fine-tuning Pipeline ===\n")

  # Step 1: Prepare training data
  training_examples <- prepare_example_trios(file_trios_dt, 50)

  if (length(training_examples) < 10) {
    stop("Insufficient training examples. Need at least 10, got ", length(training_examples))
  }

  # Step 2: Validate training data
  validate_training_data(training_examples)

  # Step 3: Save as JSONL
  jsonl_file <- save_training_jsonl(training_examples)

  # Step 4: Upload to OpenAI
  file_id <- upload_training_file(jsonl_file)

  # Step 5: Create fine-tuning job
  job_info <- create_finetuning_job(file_id)

  cat("\n=== Fine-tuning Job Created ===\n")
  cat("Job ID:", job_info$id, "\n")
  cat("Check status with: check_job_status('", job_info$id, "')\n", sep = "")
  cat("List all jobs with: list_finetuning_jobs()\n")
  cat("Fine-tuning typically takes 10-30 minutes for this size dataset.\n")

  return(list(
    job_id = job_info$id,
    file_id = file_id,
    training_examples_count = length(training_examples)
  ))
}

# Example usage:
# Assuming file_trios data.table is loaded:

# Run the complete pipeline
pipeline_result <- run_finetuning_pipeline(file_trios)

# Check job status periodically
job_status <- check_job_status(pipeline_result$job_id)

# Once complete, use the fine-tuned model
if (job_status$status == "succeeded") {
  minutes <- generate_minutes(agenda_text, transcript_text, job_status$fine_tuned_model)
  cat(minutes)
}

cat("Fine-tuning script loaded. Set OPENAI_API_KEY and run: run_finetuning_pipeline(file_trios)\n")
