library(meta)
library(ggplot2)
library(readxl)
library(openxlsx)

file_path <- "D:/Desktop/meta/E.coli.xlsx"
sheet_names <- getSheetNames(file_path)

all_results <- data.frame(
  SheetName = character(),
  StudyCount = integer(),
  Slope = numeric(),
  Slope_SE = numeric(),
  Slope_CI_Lower = numeric(),
  Slope_CI_Upper = numeric(),
  P_Value = numeric(),
  R_Squared = numeric(),
  Intercept = numeric(),
  Mean_Resistance = numeric(),
  SD_Resistance = numeric(),
  Min_Resistance = numeric(),
  Max_Resistance = numeric(),
  stringsAsFactors = FALSE
)

detailed_results <- data.frame(
  SheetName = character(),
  Study = character(),
  Year = integer(),
  Events = integer(),
  Total = integer(),
  ResistanceRate = numeric(),
  stringsAsFactors = FALSE
)

for (sheet_name in sheet_names) {

  ecoli <- read_excel(file_path, sheet = sheet_name)

  required_cols <- c("Events", "Total", "study", "year")
  if (!all(required_cols %in% colnames(ecoli))) next

  ecoli_clean <- ecoli[complete.cases(ecoli[, c("Events", "Total", "year")]), ]
  if (nrow(ecoli_clean) == 0) next

  ecoli_clean$Events <- as.numeric(ecoli_clean$Events)
  ecoli_clean$Total <- as.numeric(ecoli_clean$Total)
  ecoli_clean$year <- as.numeric(ecoli_clean$year)

  pft_result <- metaprop(
    Events, Total, study,
    data = ecoli_clean,
    sm = "PFT",
    method.tau = "REML"
  )

  plot_data <- data.frame(
    study = ecoli_clean$study,
    year = ecoli_clean$year,
    yi = pft_result$TE,
    vi = pft_result$seTE^2,
    Events = ecoli_clean$Events,
    Total = ecoli_clean$Total,
    percent = ecoli_clean$Events / ecoli_clean$Total * 100
  )

  if (nrow(plot_data) < 10) {

    all_results <- rbind(all_results, data.frame(
      SheetName = sheet_name,
      StudyCount = nrow(plot_data),
      Slope = NA,
      Slope_SE = NA,
      Slope_CI_Lower = NA,
      Slope_CI_Upper = NA,
      P_Value = NA,
      R_Squared = NA,
      Intercept = NA,
      Mean_Resistance = round(mean(plot_data$percent, na.rm = TRUE), 2),
      SD_Resistance = round(sd(plot_data$percent, na.rm = TRUE), 2),
      Min_Resistance = round(min(plot_data$percent, na.rm = TRUE), 2),
      Max_Resistance = round(max(plot_data$percent, na.rm = TRUE), 2)
    ))

    detailed_results <- rbind(
      detailed_results,
      transform(
        plot_data,
        SheetName = sheet_name,
        ResistanceRate = round(percent, 2)
      )[, c("SheetName", "study", "year", "Events", "Total", "ResistanceRate")]
    )

    next
  }

  if (length(unique(plot_data$year)) < 2) next

  weighted_lm <- lm(percent ~ year, data = plot_data, weights = Total)

  slope <- coef(weighted_lm)[2]
  slope_se <- summary(weighted_lm)$coefficients[2, 2]
  p_value <- summary(weighted_lm)$coefficients[2, 4]
  r_squared <- summary(weighted_lm)$r.squared * 100
  intercept <- coef(weighted_lm)[1]

  slope_ci <- confint(weighted_lm)[2, ]

  all_results <- rbind(all_results, data.frame(
    SheetName = sheet_name,
    StudyCount = nrow(plot_data),
    Slope = round(slope, 4),
    Slope_SE = round(slope_se, 4),
    Slope_CI_Lower = round(slope_ci[1], 4),
    Slope_CI_Upper = round(slope_ci[2], 4),
    P_Value = round(p_value, 4),
    R_Squared = round(r_squared, 1),
    Intercept = round(intercept, 2),
    Mean_Resistance = round(mean(plot_data$percent), 2),
    SD_Resistance = round(sd(plot_data$percent), 2),
    Min_Resistance = round(min(plot_data$percent), 2),
    Max_Resistance = round(max(plot_data$percent), 2)
  ))

  detailed_results <- rbind(
    detailed_results,
    transform(
      plot_data,
      SheetName = sheet_name,
      ResistanceRate = round(percent, 2)
    )[, c("SheetName", "study", "year", "Events", "Total", "ResistanceRate")]
  )

  year_seq <- seq(min(plot_data$year), max(plot_data$year), length.out = 100)
  trend_data <- data.frame(year = year_seq)
  pred <- predict(weighted_lm, newdata = trend_data, interval = "confidence")
  trend_data$predicted <- pred[, 1]

  bubble_plot <- ggplot() +
    geom_ribbon(
      data = trend_data,
      aes(x = year, ymin = pred[, 2], ymax = pred[, 3]),
      fill = "lightblue", alpha = 0.3
    ) +
    geom_line(
      data = trend_data,
      aes(x = year, y = predicted),
      linewidth = 1.5
    ) +
    geom_point(
      data = plot_data,
      aes(x = year, y = percent, size = Total),
      alpha = 0.7
    ) +
    scale_size_continuous(range = c(4, 15)) +
    scale_y_continuous(labels = scales::percent_format(scale = 1)) +
    labs(
      title = paste("Meta-regression:", sheet_name),
      x = "Publication Year",
      y = "Resistance Rate (%)",
      caption = paste0(
        "Slope = ", round(slope, 4),
        " (95% CI ", round(slope_ci[1], 4), "–", round(slope_ci[2], 4), "), ",
        "p = ", ifelse(p_value < 0.001, "<0.001", round(p_value, 3)),
        ", R² = ", round(r_squared, 1), "%"
      )
    ) +
    theme_minimal(base_size = 18)

  pdf(
    paste0("D:/Desktop/meta/Meta_Regression_", gsub("[^A-Za-z0-9]", "_", sheet_name), ".pdf"),
    width = 14,
    height = 10
  )
  print(bubble_plot)
  dev.off()
}

wb <- createWorkbook()
addWorksheet(wb, "Summary_Results")
writeData(wb, "Summary_Results", all_results)

addWorksheet(wb, "Detailed_Results")
writeData(wb, "Detailed_Results", detailed_results)

saveWorkbook(
  wb,
  "D:/Desktop/meta/Meta_Regression_Results.xlsx",
  overwrite = TRUE
)
