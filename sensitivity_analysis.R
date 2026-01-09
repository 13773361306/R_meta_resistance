library(meta)
library(readxl)
library(openxlsx)
library(dplyr)
library(purrr)

config <- list(
  file_path = "D:/Desktop/meta/E.coli.xlsx",
  required_cols = c("Events", "Total", "study"),
  results_dir = "D:/Desktop/meta/",
  min_studies = 10,
  robustness_thresholds = c(0.05, 0.10)
)

read_and_clean_data <- function(file_path, sheet_name, required_cols) {
  tryCatch({
    data <- read_excel(file_path, sheet = sheet_name)

    missing_cols <- setdiff(required_cols, names(data))
    if (length(missing_cols) > 0) return(NULL)

    cleaned <- data %>%
      mutate(
        Events = as.numeric(Events),
        Total = as.numeric(Total)
      ) %>%
      filter(
        complete.cases(Events, Total),
        Total >= 10,
        Events >= 0,
        Events <= Total
      )

    if (nrow(cleaned) == 0) return(NULL)

    cleaned
  }, error = function(e) {
    NULL
  })
}

perform_main_meta <- function(data) {
  if (nrow(data) < 2) return(NULL)

  metaprop(
    Events, Total,
    data = data,
    studlab = study,
    sm = "PFT",
    overall = TRUE
  )
}

perform_leave_one_out <- function(data, main_result) {
  if (nrow(data) < 3) return(NULL)

  original_estimate <- main_result$TE.random

  map_dfr(seq_len(nrow(data)), function(i) {
    subset_data <- data[-i, ]

    if (nrow(subset_data) < 2) return(NULL)

    meta_subset <- metaprop(
      Events, Total,
      data = subset_data,
      studlab = study,
      sm = "PFT",
      overall = TRUE
    )

    pooled_prop <- exp(meta_subset$TE.random) /
      (1 + exp(meta_subset$TE.random))

    original_prop <- exp(original_estimate) /
      (1 + exp(original_estimate))

    data.frame(
      Excluded_Study = data$study[i],
      Proportion = round(pooled_prop, 4),
      CI_Lower = round(
        exp(meta_subset$lower.random) /
          (1 + exp(meta_subset$lower.random)), 4
      ),
      CI_Upper = round(
        exp(meta_subset$upper.random) /
          (1 + exp(meta_subset$upper.random)), 4
      ),
      I2 = round(meta_subset$I2, 2),
      Change = round(pooled_prop - original_prop, 4)
    )
  })
}

assess_robustness <- function(results, thresholds) {
  if (is.null(results) || nrow(results) == 0) {
    return(list(robustness = "Not assessable", max_change = NA))
  }

  max_abs_change <- max(abs(results$Change))

  robustness <- if (max_abs_change < thresholds[1]) {
    "Highly robust"
  } else if (max_abs_change < thresholds[2]) {
    "Moderately robust"
  } else {
    "Sensitive"
  }

  list(robustness = robustness, max_change = max_abs_change)
}

run_sensitivity_analysis <- function(config) {

  sheet_names <- excel_sheets(config$file_path)
  all_results <- list()

  for (sheet_name in sheet_names) {

    data <- read_and_clean_data(
      config$file_path,
      sheet_name,
      config$required_cols
    )
    if (is.null(data)) next

    main_result <- perform_main_meta(data)
    if (is.null(main_result)) next

    if (nrow(data) < config$min_studies) next

    sensitivity_results <- perform_leave_one_out(data, main_result)
    if (is.null(sensitivity_results)) next

    robustness_info <- assess_robustness(
      sensitivity_results,
      config$robustness_thresholds
    )

    top_influential <- sensitivity_results %>%
      mutate(Absolute_Change = abs(Change)) %>%
      arrange(desc(Absolute_Change)) %>%
      slice_head(n = 3)

    all_results[[sheet_name]] <- list(
      data = data,
      main_result = main_result,
      sensitivity_results = sensitivity_results,
      robustness = robustness_info$robustness,
      max_change = robustness_info$max_change,
      top_influential = top_influential
    )
  }

  if (length(all_results) > 0) {
    save_results(all_results, config)
  }
}

save_results <- function(all_results, config) {

  summary_data <- map_dfr(names(all_results), function(sheet_name) {
    result <- all_results[[sheet_name]]

    pooled_prop <- exp(result$main_result$TE.random) /
      (1 + exp(result$main_result$TE.random))

    data.frame(
      Sheet = sheet_name,
      Studies = nrow(result$data),
      Original_Proportion = round(pooled_prop, 4),
      Max_Change = result$max_change,
      Robustness = result$robustness,
      Most_Influential_Study =
        ifelse(nrow(result$top_influential) > 0,
               result$top_influential$Excluded_Study[1], "N/A"),
      Influence_Amount =
        ifelse(nrow(result$top_influential) > 0,
               result$top_influential$Change[1], NA)
    )
  })

  detailed_data <- map_dfr(names(all_results), function(sheet_name) {
    all_results[[sheet_name]]$sensitivity_results %>%
      mutate(Sheet = sheet_name, .before = 1)
  })

  wb <- createWorkbook()

  addWorksheet(wb, "Summary")
  writeData(wb, "Summary", summary_data)

  addWorksheet(wb, "Leave_One_Out")
  writeData(wb, "Leave_One_Out", detailed_data)

  timestamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
  output_file <- file.path(
    config$results_dir,
    sprintf("Sensitivity_Analysis_%s.xlsx", timestamp)
  )

  saveWorkbook(wb, output_file, overwrite = TRUE)
}

run_sensitivity_analysis(config)
