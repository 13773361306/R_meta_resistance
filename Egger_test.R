library(readxl)
library(metafor)

file_path <- "D:/Desktop/meta/E.coli.xlsx"
sheet_names <- excel_sheets(file_path)

results <- data.frame()

for(sheet_name in sheet_names) {
  cat("Analyzing:", sheet_name, "... ")
  
  tryCatch({
    data <- read_excel(file_path, sheet = sheet_name)
    
    df <- data.frame(
      events = as.numeric(data$Events),
      total = as.numeric(data$Total)
    )
    
    df <- na.omit(df)
    
    if(nrow(df) >= 3) {
      df <- escalc(measure = "PFT", xi = df$events, ni = df$total, data = df)
      
      res <- rma(yi, vi, data = df, method = "REML")
      
      egger <- regtest(res)
      
      results <- rbind(results, data.frame(
        Sheet = sheet_name,
        Studies = nrow(df),
        CI_Lower = round(res$ci.lb, 4),
        CI_Upper = round(res$ci.ub, 4),
        I2 = round(res$I2, 2),
        Tau2 = round(res$tau2, 4),
        Egger_z = round(egger$zval, 4),
        Egger_p = round(egger$pval, 4),
        Bias = ifelse(egger$pval < 0.05, "Yes", "No")
      ))
      
      pdf_file <- paste0("D:/Desktop/meta/", sheet_name, "_PFT_funnel.pdf")
      pdf(pdf_file, width = 7, height = 6)
      
      funnel(res, 
             xlab = "Freeman-Tukey Transformed Proportion",
             ylab = "Standard Error",
             main = paste("Funnel Plot:", sheet_name),
             level = c(0.9, 0.95),
             shade = c("gray90", "gray75"),
             refline = res$beta[1])
      
      legend("topright", 
             legend = c(
               paste("Studies:", nrow(df)),
               paste("Egger test: p =", round(egger$pval, 4)),
               paste("I² =", round(res$I2, 1), "%"),
               paste("τ² =", round(res$tau2, 4))
             ),
             bty = "n",
             cex = 0.9)
      
      dev.off()
      
      tiff_file <- paste0("D:/Desktop/meta/", sheet_name, "_PFT_funnel.tiff")
      tiff(tiff_file, width = 2400, height = 2000, res = 300, compression = "lzw")
      
      par(mar = c(4.5, 4.5, 3, 2))
      
      funnel(res, 
             xlab = "Freeman-Tukey Transformed Proportion",
             ylab = "Standard Error",
             main = paste("Funnel Plot:", sheet_name),
             level = c(0.9, 0.95),
             shade = c("gray90", "gray75"),
             refline = res$beta[1],
             cex.lab = 1.2,
             cex.axis = 1.1,
             cex.main = 1.3,
             pch = 19,
             cex = 1.2)
      
      legend("topright", 
             legend = c(
               paste("Studies:", nrow(df)),
               paste("Egger test: p =", round(egger$pval, 4)),
               paste("I² =", round(res$I2, 1), "%")
             ),
             bty = "n",
             cex = 1.1)
      
      dev.off()
      
      cat("completed (p =", round(egger$pval, 4), ")\n")
    } else {
      cat("skipped (insufficient studies)\n")
    }
    
  }, error = function(e) {
    cat("error:", e$message, "\n")
  })
}

cat("\n")
cat(paste(rep("=", 50), collapse = ""))
cat("\nAnalysis completed!\n\n")

if(nrow(results) > 0) {
  print(results)
  
  write.csv(results, "D:/Desktop/meta/Egger_PFT_Results.csv", row.names = FALSE)
  cat("\n✓ Results saved: D:/Desktop/meta/Egger_PFT_Results.csv\n")
  
  if(requireNamespace("writexl", quietly = TRUE)) {
    library(writexl)
    write_xlsx(results, "D:/Desktop/meta/Egger_PFT_Results.xlsx")
    cat("✓ Results saved: D:/Desktop/meta/Egger_PFT_Results.xlsx\n")
  }
  
  txt_file <- "D:/Desktop/meta/Egger_PFT_Results.txt"
  sink(txt_file)
  cat("Egger Test Results (PFT Transformation)\n")
  cat("Generated:", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "\n")
  cat(paste(rep("-", 60), collapse = ""), "\n")
  print(results)
  sink()
  cat("✓ Results saved: D:/Desktop/meta/Egger_PFT_Results.txt\n")
  
  cat("\nSummary statistics:\n")
  cat("- Total sheets analyzed:", nrow(results), "\n")
  cat("- Publication bias present (p < 0.05):", sum(results$Bias == "Yes"), "\n")
  cat("- No publication bias (p ≥ 0.05):", sum(results$Bias == "No"), "\n")
  cat("- Average number of studies:", round(mean(results$Studies), 1), "\n")
  cat("- Average I²:", round(mean(results$I2, na.rm = TRUE), 1), "%\n")
  cat("- Average τ²:", round(mean(results$Tau2, na.rm = TRUE), 4), "\n")
  
  p_vals <- results$Egger_p
  cat("\nEgger test p-value distribution:\n")
  cat("- p < 0.001:", sum(p_vals < 0.001, na.rm = TRUE), "\n")
  cat("- 0.001 ≤ p < 0.01:", sum(p_vals >= 0.001 & p_vals < 0.01, na.rm = TRUE), "\n")
  cat("- 0.01 ≤ p < 0.05:", sum(p_vals >= 0.01 & p_vals < 0.05, na.rm = TRUE), "\n")
  cat("- 0.05 ≤ p < 0.1:", sum(p_vals >= 0.05 & p_vals < 0.1, na.rm = TRUE), "\n")
  cat("- p ≥ 0.1:", sum(p_vals >= 0.1, na.rm = TRUE), "\n")
  
  cat("\nImage file statistics:\n")
  pdf_files <- list.files("D:/Desktop/meta/", pattern = "_PFT_funnel\\.pdf$", full.names = TRUE)
  tiff_files <- list.files("D:/Desktop/meta/", pattern = "_PFT_funnel\\.tiff$", full.names = TRUE)
  
  cat("- PDF files:", length(pdf_files), "\n")
  cat("- TIFF files:", length(tiff_files), "\n")
  
  cat("\nGenerating summary plot...\n")
  
  summary_pdf <- "D:/Desktop/meta/Egger_PFT_Summary.pdf"
  pdf(summary_pdf, width = 10, height = 8)
  
  par(mfrow = c(2, 2))
  
  hist(p_vals, breaks = 20, 
       col = "#669999", border = "white",
       main = "Distribution of Egger Test p-values",
       xlab = "p-value", ylab = "Frequency",
       xlim = c(0, 1))
  abline(v = 0.05, col = "red", lwd = 2, lty = 2)
  abline(v = 0.1, col = "orange", lwd = 2, lty = 2)
  legend("topright", 
         legend = c("p < 0.05: Significant bias",
                   "0.05 ≤ p < 0.1: Marginal",
                   "p ≥ 0.1: No bias"),
         col = c("red", "orange", "gray"),
         lty = 2, lwd = 2, bty = "n", cex = 0.9)
  
  hist(results$I2, breaks = 20,
       col = "#336699", border = "white",
       main = "Distribution of Heterogeneity (I²)",
       xlab = "I² (%)", ylab = "Frequency",
       xlim = c(0, 100))
  abline(v = 50, col = "red", lwd = 2, lty = 2)
  abline(v = 75, col = "orange", lwd = 2, lty = 2)
  legend("topright",
         legend = c("I² ≥ 75%: High heterogeneity",
                   "50% ≤ I² < 75%: Moderate",
                   "I² < 50%: Low heterogeneity"),
         col = c("orange", "red", "gray"),
         lty = 2, lwd = 2, bty = "n", cex = 0.9)
  
  hist(results$Studies, breaks = 20,
       col = "#99CCCC", border = "white",
       main = "Distribution of Study Numbers",
       xlab = "Number of Studies", ylab = "Frequency")
  
  plot(results$I2, -log10(p_vals),
       pch = 19, cex = 1.2,
       col = ifelse(p_vals < 0.05, "red", 
                   ifelse(p_vals < 0.1, "orange", "#336699")),
       main = "Egger p-value vs Heterogeneity",
       xlab = "I² (%)", ylab = "-log10(p-value)")
  abline(h = -log10(0.05), col = "red", lty = 2, lwd = 1.5)
  abline(h = -log10(0.1), col = "orange", lty = 2, lwd = 1.5)
  legend("topright",
         legend = c("p < 0.05", "0.05 ≤ p < 0.1", "p ≥ 0.1"),
         col = c("red", "orange", "#336699"),
         pch = 19, bty = "n")
  
  dev.off()
  
  cat("✓ Summary plot saved:", summary_pdf, "\n")
  
} else {
  cat("No sheets were successfully analyzed\n")
}

cat(paste(rep("=", 50), collapse = ""))
cat("\n")