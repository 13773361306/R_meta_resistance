library(meta)
library(readxl)
library(writexl)
library(openxlsx)
library(dplyr)

file_path <- "D:/Desktop/meta/E.coli.xlsx"

sheet_names <- excel_sheets(file_path)
cat("Sheets found:", paste(sheet_names, collapse = ", "), "\n\n")

all_results <- list()

for(sheet_name in sheet_names) {
  cat("Analyzing:", sheet_name, "\n")
  
  ecoli <- read_excel(file_path, sheet = sheet_name)
  
  required_cols <- c("Events", "Total", "study", "Region")
  if(!all(required_cols %in% names(ecoli))) {
    missing_cols <- setdiff(required_cols, names(ecoli))
    cat("Warning:", sheet_name, "missing columns:", paste(missing_cols, collapse = ", "), "- skipping\n\n")
    next
  }
  
  ecoli <- ecoli %>%
    mutate(
      Events = as.numeric(Events),
      Total = as.numeric(Total)
    )
  
  ecoli_clean <- ecoli %>%
    filter(
      complete.cases(across(c("Events", "Total", "Region"))),
      Total >= 10,
      Events >= 0,
      Events <= Total
    )
  
  if(nrow(ecoli_clean) == 0) {
    cat("Warning:", sheet_name, "no valid data after cleaning - skipping\n\n")
    next
  }
  
  removed_na <- nrow(ecoli) - nrow(ecoli %>% filter(complete.cases(across(c("Events", "Total", "Region")))))
  removed_small <- nrow(ecoli %>% filter(complete.cases(across(c("Events", "Total", "Region"))))) - nrow(ecoli_clean)
  
  cat("  Data cleaning summary:\n")
  cat("  - Original studies:", nrow(ecoli), "\n")
  if(removed_na > 0) {
    cat("  - Studies with missing values removed:", removed_na, "\n")
  }
  if(removed_small > 0) {
    cat("  - Studies with Total < 10 removed:", removed_small, "\n")
  }
  cat("  - Studies included in analysis:", nrow(ecoli_clean), "\n")
  
  if(removed_small > 0) {
    small_studies <- ecoli %>% 
      filter(complete.cases(across(c("Events", "Total", "Region"))), Total < 10)
    if(nrow(small_studies) > 0) {
      cat("  - Removed small sample studies:", paste(small_studies$study, collapse = ", "), "\n")
    }
  }
  
  cat("  Events data type:", class(ecoli_clean$Events), "\n")
  cat("  Total data type:", class(ecoli_clean$Total), "\n")
  cat("  Events range:", min(ecoli_clean$Events), "to", max(ecoli_clean$Events), "\n")
  cat("  Total range:", min(ecoli_clean$Total), "to", max(ecoli_clean$Total), "\n")
  
  tryCatch({
    meta.region <- metaprop(Events, Total, data = ecoli_clean,
                        studlab = study,
                        sm = "PFT",
                        subgroup = Region,
                        overall = TRUE)
    
    pdf_file <- paste0("D:/Desktop/meta/", sheet_name, "_Region_forest.pdf")
    pdf(pdf_file, height = 35, width = 22, family = "Times")
    
    forest(meta.region, 
           layout = "Revman",
           leftcols = c("studlab", "event", "n", "effect.ci"),
           rightcols = FALSE,
           family = "Times",
           fontsize = 19,
           
           col.vertical.line = "transparent",
           col.diamond.common = "lightslategray",
           col.diamond.lines.common = "lightslategray",
           col.diamond.random = "#669999",
           col.diamond.lines.random = "#669999",
           col.square = "#336699",
           col.square.lines = "#336699",
           col.study = "#006666",
           col.inside = "black",
           col.label.right = "black",
           col.random = "#660099",
           col.subgroup = "#003399",
           
           overall = TRUE,
           overall.hetstat = TRUE,
           squaresize = 1.0,
           height = 0.9,
           
           plotwidth = "18cm",
           colgap.forest.left = "4cm",
           colgap.left = "3.5cm",
           
           spacing = 1.5,
           spacing.subgroup = 1.8,
           
           common = FALSE,
           xlab = "",
           xlim = c(0, 1),
           smlab = "Proportion",
           
           fs.hetstat = 19,
           fs.test.subgroup = 19,
           fs.test.overall = 19,
           fs.heading = 19,
           fs.study = 19,
           fs.axis = 19,
           fs.common = 19,
           fs.random = 19,
           fs.label = 19,
           fs.xlab = 19,
           fs.smlab = 19,
           fs.subgroup = 19,
           fs.subgroup.labels = 19,
           fs.addline = 19,
           
           addrow = TRUE,
           addrow.overall = TRUE,
           addrow.subgroups = TRUE,
           addrow.after.subgroups = TRUE,
           
           top = 4,
           bottom = 4,
           
           lwd = 1.2,
           lwd.square = 1.5,
           lwd.diamond = 1.8)
    
    dev.off()
    
    cat("  Region subgroup forest plot saved:", pdf_file, "\n")
    
    subgroup_test <- data.frame(
      Sheet_Name = sheet_name,
      Test = c("Test for subgroup differences (Common effect)", 
               "Test for subgroup differences (Random effects)"),
      Chi_squared = c(meta.region$Q.b.common, meta.region$Q.b.random),
      df = c(meta.region$df.Q.b.common, meta.region$df.Q.b.random),
      P_value = c(meta.region$pval.Q.b.common, meta.region$pval.Q.b.random),
      Status = "Success"
    )
    
    subgroup_details <- data.frame(
      Sheet_Name = sheet_name,
      Region = c(meta.region$bylevs, "Overall"),
      Studies = c(meta.region$k.w, meta.region$k),
      Events = c(meta.region$event.w, sum(meta.region$event)),
      Total = c(meta.region$n.w, sum(meta.region$n)),
      Proportion = c(meta.region$TE.w, meta.region$TE.random),
      CI_lower = c(meta.region$lower.w, meta.region$lower.random),
      CI_upper = c(meta.region$upper.w, meta.region$upper.random),
      I2 = c(meta.region$I2.w, meta.region$I2),
      Status = "Success"
    )
    
    all_results[[sheet_name]] <- list(
      meta_object = meta.region,
      subgroup_test = subgroup_test,
      subgroup_details = subgroup_details,
      cleaned_data = ecoli_clean,
      removed_na = removed_na,
      removed_small = removed_small,
      analysis_success = TRUE
    )
    
    cat("  Region subgroup analysis: included", meta.region$k, "studies,", 
        length(meta.region$bylevs), "regions\n")
    cat("  Test for subgroup differences p-value (random effects):", round(meta.region$pval.Q.b.random, 4), "\n\n")
    
  }, error = function(e) {
    cat("  Meta-analysis error:", e$message, "\n")
    
    subgroup_test <- data.frame(
      Sheet_Name = sheet_name,
      Test = c("Test for subgroup differences (Common effect)", 
               "Test for subgroup differences (Random effects)"),
      Chi_squared = NA,
      df = NA,
      P_value = NA,
      Status = paste("Error:", e$message)
    )
    
    if(nrow(ecoli_clean) > 0) {
      total_events <- sum(ecoli_clean$Events, na.rm = TRUE)
      total_participants <- sum(ecoli_clean$Total, na.rm = TRUE)
      region_counts <- table(ecoli_clean$Region)
      
      subgroup_details <- data.frame(
        Sheet_Name = sheet_name,
        Region = c(names(region_counts), "Overall"),
        Studies = c(as.numeric(region_counts), nrow(ecoli_clean)),
        Events = c(rep(NA, length(region_counts)), total_events),
        Total = c(rep(NA, length(region_counts)), total_participants),
        Proportion = NA,
        CI_lower = NA,
        CI_upper = NA,
        I2 = NA,
        Status = paste("Error:", e$message)
      )
    } else {
      subgroup_details <- data.frame(
        Sheet_Name = sheet_name,
        Region = "No data",
        Studies = 0,
        Events = 0,
        Total = 0,
        Proportion = NA,
        CI_lower = NA,
        CI_upper = NA,
        I2 = NA,
        Status = paste("Error:", e$message)
      )
    }
    
    all_results[[sheet_name]] <- list(
      meta_object = NULL,
      subgroup_test = subgroup_test,
      subgroup_details = subgroup_details,
      cleaned_data = ecoli_clean,
      removed_na = removed_na,
      removed_small = removed_small,
      analysis_success = FALSE
    )
    
    cat("  Error recorded, continuing to next sheet\n\n")
  })
}

if(length(all_results) > 0) {
  
  all_subgroup_tests <- do.call(rbind, lapply(all_results, function(x) x$subgroup_test))
  
  all_subgroup_details <- do.call(rbind, lapply(all_results, function(x) x$subgroup_details))
  
  wb <- createWorkbook()
  
  addWorksheet(wb, "Region_Subgroup_Tests")
  writeData(wb, "Region_Subgroup_Tests", all_subgroup_tests)
  
  addWorksheet(wb, "Region_Subgroup_Results")
  writeData(wb, "Region_Subgroup_Results", all_subgroup_details)
  
  for(sheet_name in names(all_results)) {
    
    if(!is.null(all_results[[sheet_name]]$meta_object) && all_results[[sheet_name]]$analysis_success) {
      
      short_sheet_name <- substr(sheet_name, 1, min(20, nchar(sheet_name)))
      data_sheet_name <- paste0("Reg_", short_sheet_name)
      
      if(nchar(data_sheet_name) > 31) {
        data_sheet_name <- substr(data_sheet_name, 1, 31)
      }
      
      study_level_data <- data.frame(
        Study = all_results[[sheet_name]]$meta_object$studlab,
        Region = all_results[[sheet_name]]$meta_object$byvar,
        Events = all_results[[sheet_name]]$meta_object$event,
        Total = all_results[[sheet_name]]$meta_object$n,
        Proportion = round(all_results[[sheet_name]]$meta_object$TE, 4)
      )
      
      addWorksheet(wb, data_sheet_name)
      writeData(wb, data_sheet_name, study_level_data)
    }
    
    short_sheet_name <- substr(sheet_name, 1, min(20, nchar(sheet_name)))
    cleaned_sheet_name <- paste0("Clean_", short_sheet_name)
    
    if(nchar(cleaned_sheet_name) > 31) {
      cleaned_sheet_name <- substr(cleaned_sheet_name, 1, 31)
    }
    
    addWorksheet(wb, cleaned_sheet_name)
    writeData(wb, cleaned_sheet_name, all_results[[sheet_name]]$cleaned_data)
  }
  
  results_file <- "D:/Desktop/meta/All_Sheets_Region_Analysis.xlsx"
  saveWorkbook(wb, results_file, overwrite = TRUE)
  
  cat("=== Region subgroup analysis completed ===\n")
  cat("Processed", length(all_results), "sheets with Region subgroup data\n")
  
  success_count <- sum(sapply(all_results, function(x) x$analysis_success))
  error_count <- length(all_results) - success_count
  
  cat("Sheets successfully analyzed:", success_count, "\n")
  cat("Sheets with errors:", error_count, "\n")
  cat("Forest plots saved in: D:/Desktop/meta/ (ending with _Region_forest.pdf)\n")
  cat("Summary results saved in:", results_file, "\n")
  
  cat("\n=== Sheets with significant Region subgroup differences ===\n")
  significant_sheets <- unique(all_subgroup_tests$Sheet_Name[all_subgroup_tests$P_value < 0.05 & 
                                                             grepl("Random", all_subgroup_tests$Test) &
                                                             !is.na(all_subgroup_tests$P_value)])
  if(length(significant_sheets) > 0) {
    cat(paste(significant_sheets, collapse = ", "), "\n")
  } else {
    cat("No significant Region subgroup differences\n")
  }
  
  cat("\n=== Region distribution by sheet ===\n")
  for(sheet_name in names(all_results)) {
    if(all_results[[sheet_name]]$analysis_success) {
      regions <- unique(all_results[[sheet_name]]$subgroup_details$Region)
      regions <- regions[regions != "Overall"]
      cat(sheet_name, ":", length(regions), "regions -", paste(regions, collapse = ", "), "\n")
    } else {
      cat(sheet_name, ": analysis failed -", all_results[[sheet_name]]$subgroup_details$Status[1], "\n")
    }
  }
  
  cat("\n=== Data cleaning summary ===\n")
  for(sheet_name in names(all_results)) {
    original_count <- nrow(read_excel(file_path, sheet = sheet_name))
    cleaned_count <- nrow(all_results[[sheet_name]]$cleaned_data)
    removed_na <- all_results[[sheet_name]]$removed_na
    removed_small <- all_results[[sheet_name]]$removed_small
    
    cat(sheet_name, ": ", original_count, "original studies -> ", cleaned_count, 
        "analyzed studies\n")
    if(removed_na > 0 | removed_small > 0) {
      cat("     Reasons: missing values(", removed_na, "), ",
          "small sample studies(", removed_small, ")\n")
    }
  }
  
} else {
  cat("=== Warning ===\n")
  cat("No sheets found containing required columns (Events, Total, study, Region)\n")
}
