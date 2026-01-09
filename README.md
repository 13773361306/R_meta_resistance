# R_meta_resistance

[README.md](https://github.com/user-attachments/files/24520899/README.md)

# Meta-analysis Project: Intra-abdominal Infection Antibiotic Resistance Rates

## Overview
This project performs comprehensive meta-analyses of antibiotic resistance rates for *Escherichia coli* isolated from intra-abdominal infections. The code automates the analysis of multiple Excel sheets containing study data and generates forest plots and summary statistics.

## Project Description
This is a meta-analysis project specifically designed to analyze the resistance rates of common antibiotics against pathogens causing intra-abdominal infections.

## Features
- **Batch Processing**: Automatically analyzes multiple Excel sheets in a single run  
- **Data Cleaning**: Filters studies with missing data and small sample sizes (Total < 10)  
- **Meta-analysis**: Calculates pooled resistance proportions using Freeman-Tukey double arcsine transformation  
- **Forest Plots**: Generates publication-quality forest plots for each analysis  
- **Comprehensive Output**: Produces detailed Excel reports with summary statistics and cleaned data  

## Prerequisites

### Required R Packages
Ensure the following packages are installed:

```r
install.packages(c("meta", "readxl", "writexl", "openxlsx", "dplyr"))
```

### Data Requirements
The input Excel file should:
- Contain multiple sheets (one for each antibiotic/analysis)  
- Each sheet must include the following columns:  
  - `study`: Study identifier  
  - `Events`: Number of resistant isolates  
  - `Total`: Total number of isolates tested  

### File Structure
```
Desktop/meta/
├── antibiotic_resistance_raw_data/
│   └── E.coli.xlsx                 # Input data file
├── [antibiotic_name]_forest.pdf    # Generated forest plots
└── All_Sheets_Meta_Analysis_Results.xlsx  # Summary results
```

## Usage

### 1. Prepare Input Data
- Place your Excel file at: `D:/Desktop/meta/antibiotic_resistance_raw_data/E.coli.xlsx`  
- Organize data into separate sheets (one per antibiotic or subgroup analysis)  
- Ensure each sheet has the required columns: `study`, `Events`, `Total`  

### 2. Run the Analysis
- Execute the entire R script. The code will:
  - Automatically detect all sheets in the Excel file  
  - Process each sheet individually  
  - Generate forest plots as PDF files  
  - Create a comprehensive Excel summary report  

### 3. Output Files

#### Forest Plots
- **Location**: `D:/Desktop/meta/`  
- **Format**: PDF files named `[sheet_name]_forest.pdf`  
- **Content**: Forest plots showing individual study proportions and pooled estimates  

#### Summary Report
- **File**: `All_Sheets_Meta_Analysis_Results.xlsx`  
- **Contains**:
  - `Summary_All_Sheets`: Overall results for all analyses  
  - `Det_[short_name]`: Detailed study-level data for each analysis  
  - `Data_[short_name]`: Cleaned data used for each analysis  

## Analysis Details

### Data Cleaning Process
- **Type Conversion**: Ensures `Events` and `Total` are numeric  
- **NA Removal**: Excludes studies with missing `Events` or `Total`  
- **Sample Size Filter**: Removes studies with `Total < 10`  
- **Logic Check**: Ensures `Events ≤ Total` and `Events ≥ 0`  

### Statistical Methods
- **Effect Measure**: Proportion (resistance rate)  
- **Transformation**: Freeman-Tukey double arcsine (PFT)  
- **Pooling Method**: Random-effects model  
- **Heterogeneity**: I² statistic and Q-test  

### Key Output Statistics
- Number of studies included  
- Pooled resistance proportion with 95% CI  
- Heterogeneity measures (I², Q-statistic, p-value)  
- Data cleaning statistics  

## Customization
- **Modify File Paths**:  
```r
file_path <- "D:/Desktop/meta/antibiotic_resistance_raw_data/E.coli.xlsx"
```  
- **Adjust Sample Size Threshold**:  
```r
Total >= 10  # Change this value to adjust the threshold
```  
- **Customize Forest Plot Appearance**: Modify the `forest()` function parameters for different visual styles.  

## Error Handling
- Skips sheets with missing required columns  
- Continues processing even if individual analyses fail  
- Records error messages for troubleshooting  

### Troubleshooting
- **Missing required columns**: Ensure each Excel sheet contains `study`, `Events`, `Total`  
- **No valid data after cleaning**: Check for missing values or small sample sizes  
- **Meta-analysis errors**: Ensure sufficient studies are available (≥2) and check for extreme proportions  

### Console Output Messages
- Sheet-by-sheet analysis status  
- Data cleaning statistics  
- Analysis results summary  
- Error messages with context  

## Results Interpretation

### Forest Plots
- **Squares**: Individual study proportions (size indicates weight)  
- **Diamonds**: Pooled estimates with 95% confidence intervals  
- **I²**: Heterogeneity statistic (higher values indicate greater heterogeneity)  

### Excel Report
- **Status**: `"Success"` or error message  
- **I²**: Percentage of variability due to heterogeneity  
- **95% CI**: Confidence interval for pooled proportion  

## Version Information
- **R Version**: 4.0 or higher recommended  
- **Package Versions**:  
  - `meta`: ≥ 5.2-0  
  - `readxl`: ≥ 1.3.1  
  - `dplyr`: ≥ 1.0.0  

## Author & Contact
- **Project**: Meta-analysis of Intra-abdominal Infection Resistance  
- **Purpose**: Research and academic analysis  
- **Last Updated**: 2026-01-08  

## License
For academic and research use.  

## Citation
If using this code for publication, please cite the `meta` R package and provide appropriate attribution.

## Support
For questions or issues, check:
- Console error messages  
- Excel file structure requirements  
- R package installation status
