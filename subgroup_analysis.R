library(meta)
library(readxl)
library(writexl)
library(openxlsx)
library(dplyr)

file_path <- "D:/Desktop/meta/E.coli.xlsx"

sheet_names <- excel_sheets(file_path)
all_results <- list()

shorten_sheet_name <- function(name, max_length = 31) {
  if (nchar(name) <= max_length) return(name)
  paste0(substr(name, 1, 25), "...")
}

for (sheet_name in sheet_names) {

  data_raw <- read_excel(file_path, sheet = sheet_name)

  if (!all(c("Events", "Total", "study") %in% names(data_raw))) next

  data_clean <- data_raw %>%
    mutate(
      Events = as.numeric(Events),
      Total  = as.numeric(Total)
    ) %>%
    filter(
      !is.na(Events),
      !is.na(Total),
      Total >= 10,
      Events >= 0,
      Events <= Total
    )

  if (nrow(data_clean) == 0) next

  meta_res <- tryCatch(
    metaprop(
      Events, Total,
      data = data_clean,
      studlab = study,
      sm = "PFT",
      overall = TRUE
    ),
    error = function(e) NULL
  )

  if (is.null(meta_res)) next

  pdf(
    paste0("D:/Desktop/meta/", sheet_name, "_forest.pdf"),
    height = 20,
    width = 12,
    family = "Times"
  )

  forest(
    meta_res,
    layout = "Revman",
    leftcols = c("studlab", "event", "n", "effect.ci"),
    rightcols = FALSE,
    fontsize = 9.5,
    col.vertical.line = "transparent",
    col.diamond.common = "lightslategray",
    col.diamond.random = "#669999",
    col.square = "#336699",
    overall = TRUE,
    overall.hetstat = TRUE,
    common = FALSE,
    xlim = c(0, 1),
    smlab = "Proportion"
  )

  dev.off()

  summary_df <- data.frame(
    Sheet = sheet_name,
    Studies = meta_res$k,
    Total_Participants = sum(meta_res$n),
    Total_Events = sum(meta_res$event),
    Proportion = round(meta_res$TE.random, 4),
    CI_Lower = round(meta_res$lower.random, 4),
    CI_Upper = round(meta_res$upper.random, 4),
    I2 = round(meta_res$I2 * 100, 2),
    Q = round(meta_res$Q, 4),
    df = meta_res$df.Q,
    P_Q = round(meta_res$pval.Q, 4)
  )

  all_results[[sheet_name]] <- list(
    summary = summary_df,
    meta = meta_res,
    data = data_clean
  )
}

if (length(all_results) > 0) {

  all_summary <- do.call(rbind, lapply(all_results, function(x) x$summary))

  wb <- createWorkbook()
  addWorksheet(wb, "Summary")
  writeData(wb, "Summary", all_summary)

  for (nm in names(all_results)) {

    short_nm <- shorten_sheet_name(nm)

    addWorksheet(wb, substr(paste0("Detail_", short_nm), 1, 31))
    writeData(
      wb,
      substr(paste0("Detail_", short_nm), 1, 31),
      data.frame(
        Study = all_results[[nm]]$meta$studlab,
        Events = all_results[[nm]]$meta$event,
        Total = all_results[[nm]]$meta$n,
        Proportion = round(all_results[[nm]]$meta$TE, 4),
        Weight_Random = round(all_results[[nm]]$meta$w.random, 2)
      )
    )

    addWorksheet(wb, substr(paste0("Data_", short_nm), 1, 31))
    writeData(
      wb,
      substr(paste0("Data_", short_nm), 1, 31),
      all_results[[nm]]$data
    )
  }

  saveWorkbook(
    wb,
    "D:/Desktop/meta/All_Sheets_Meta_Analysis_Results.xlsx",
    overwrite = TRUE
  )
}
