
library(meta)
library(readxl)
library(writexl)
library(openxlsx)
library(dplyr)

file_path <- "D:/Desktop/meta/E.coli.xlsx"

sheet_names <- excel_sheets(file_path)
cat("Found sheets:", paste(sheet_names, collapse = ", "), "\n\n")

all_results <- list()

shorten_sheet_name <- function(name, max_length = 31) {
  if (nchar(name) <= max_length) {
    name
  } else {
    paste0(substr(name, 1, 25), "...")
  }
}

for (sheet_name in sheet_names) {
  cat("Analyzing:", sheet_name, "\n")
  
  ecoli <- read_excel(file_path, sheet = sheet_name)
  
  if (!all(c("Events", "Total", "study") %in% names(ecoli))) {
    cat("Warning:", sheet_name, "missing required columns, skipping\n\n")
    next
  }
  
  ecoli <- ecoli %>%
    mutate(
      Events = as.numeric(Events),
      Total  = as.numeric(Total)
    )
  
  ecoli_clean <- ecoli %>%
    filter(
      !is.na(Events),
      !is.na(Total),
      Total >= 10,
      Events >= 0,
      Events <= Total
    )
  
  if (nrow(ecoli_clean) == 0) {
    cat("Warning:", sheet_name, "no valid data after cleaning, skipping\n\n")
    next
  }
  
  removed_na <- nrow(ecoli) - nrow(ecoli %>% filter(!is.na(Events), !is.na(Total)))
  removed_small <- nrow(ecoli %>% filter(!is.na(Events), !is.na(Total))) - nrow(ecoli_clean)
  total_removed <- removed_na + removed_small
  
  tryCatch({
    meta_overall <- metaprop(
      Events, Total,
      data = ecoli_clean,
      studlab = study,
      sm = "PFT",
      overall = TRUE
    )
    
    pdf_file <- paste0("D:/Desktop/meta/", sheet_name, "_forest.pdf")
    pdf(pdf_file, height = 20, width = 12, family = "Times")
    
    forest(
      meta_overall,
      layout = "Revman",
      leftcols = c("studlab", "event", "n", "effect.ci"),
      rightcols = FALSE,
      fontsize = 9.5,
      overall = TRUE,
      overall.hetstat = TRUE,
      weight.study = "same",
      squaresize = 0.8,
      height = 0.5,
      plotwidth = "12cm",
      colgap.forest.left = "1.5cm",
      colgap.left = "1.5cm",
      common = FALSE,
      xlim = c(0, 1),
      smlab = "Proportion"
    )
    
    dev.off()
    
    results_df <- data.frame(
      Sheet_Name = sheet_name,
      Studies_Original = nrow(ecoli),
      Studies_Analyzed = meta_overall$k,
      Studies_Removed_NA = removed_na,
      Studies_Removed_Small = removed_small,
      Total_Removed = total_removed,
      Total_Participants = sum(meta_overall$n),
      Total_Events = sum(meta_overall$event),
      Overall_Proportion = round(meta_overall$TE.random, 4),
      CI_95_Lower = round(meta_overall$lower.random, 4),
      CI_95_Upper = round(meta_overall$upper.random, 4),
      I2 = round(meta_overall$I2 * 100, 2),
      P_value_heterogeneity = round(meta_overall$pval.Q, 4),
      Q_statistic = round(meta_overall$Q, 4),
      df = meta_overall$df.Q,
      Status = "Success"
    )
    
    all_results[[sheet_name]] <- list(
      summary = results_df,
      meta_object = meta_overall,
      cleaned_data = ecoli_clean
    )
    
  }, error = function(e) {
    results_df <- data.frame(
      Sheet_Name = sheet_name,
      Studies_Original = nrow(ecoli),
      Studies_Analyzed = nrow(ecoli_clean),
      Studies_Removed_NA = removed_na,
      Studies_Removed_Small = removed_small,
      Total_Removed = total_removed,
      Total_Participants = sum(ecoli_clean$Total, na.rm = TRUE),
      Total_Events = sum(ecoli_clean$Events, na.rm = TRUE),
      Overall_Proportion = NA,
      CI_95_Lower = NA,
      CI_95_Upper = NA,
      I2 = NA,
      P_value_heterogeneity = NA,
      Q_statistic = NA,
      df = NA,
      Status = paste("Error:", e$message)
    )
    
    all_results[[sheet_name]] <- list(
      summary = results_df,
      meta_object = NULL,
      cleaned_data = ecoli_clean
    )
  })
}

if (length(all_results) > 0) {
  all_summaries <- do.call(rbind, lapply(all_results, function(x) x$summary))
  results_file <- "D:/Desktop/meta/All_Sheets_Meta_Analysis_Results.xlsx"
  
  wb <- createWorkbook()
  addWorksheet(wb, "Summary_All_Sheets")
  writeData(wb, "Summary_All_Sheets", all_summaries)
  
  for (sheet_name in names(all_results)) {
    short_name <- shorten_sheet_name(sheet_name)
    
    if (!is.null(all_results[[sheet_name]]$meta_object)) {
      study_details <- data.frame(
        Study = all_results[[sheet_name]]$meta_object$studlab,
        Events = all_results[[sheet_name]]$meta_object$event,
        Total = all_results[[sheet_name]]$meta_object$n,
        Proportion = round(all_results[[sheet_name]]$meta_object$TE, 4),
        Weight_Random = round(all_results[[sheet_name]]$meta_object$w.random, 2)
      )
      
      details_sheet_name <- substr(paste0("Det_", short_name), 1, 31)
      addWorksheet(wb, details_sheet_name)
      writeData(wb, details_sheet_name, study_details)
    }
    
    cleaned_sheet_name <- substr(paste0("Data_", short_name), 1, 31)
    addWorksheet(wb, cleaned_sheet_name)
    writeData(wb, cleaned_sheet_name, all_results[[sheet_name]]$cleaned_data)
  }
  
  saveWorkbook(wb, results_file, overwrite = TRUE)
}


