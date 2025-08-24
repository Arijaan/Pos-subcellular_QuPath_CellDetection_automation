library(ggplot2)
library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)
library(ggsignif)

# Get command line arguments
args <- commandArgs(trailingOnly = TRUE)

if (length(args) >= 1) {
  # Use provided file path from measurement analyzer
  path1 <- args[1]
  cat("Using provided file:", path1, "\n")
} else {
  # Fallback to file chooser for manual use
  cat("Select the Combined_Summary.xlsx file...\n")
  path1 <- file.choose()
}

# Set output folder based on input file location when called from measurement analyzer
if (length(args) >= 1) {
  # When called from measurement analyzer, use the input directory
  input_dir <- dirname(path1)
  folder_name <- "Weighted_BarPlot_Analysis"
  base_output_dir <- input_dir
} else {
  # Manual mode - use default settings
  folder_name <- "test"
  base_output_dir <- "C:/Users/psoor/OneDrive/Desktop/Bachelorarbeit/Figures/plots and stats"
}

var1 <- "Control"
var2 <- "T1D"
var3 <- "T2D"
var4 <- "Aab"
fill_cust <- ""
title_cust <- ""

# Clean the folder name (remove invalid characters)
folder_name <- gsub("[^a-zA-Z0-9_-]", "_", folder_name)
if (folder_name == "" || is.na(folder_name)) {
  folder_name <- paste0("analysis_", format(Sys.time(), "%Y%m%d_%H%M%S"))
}

output_dir <- file.path(base_output_dir, folder_name)

# Create output directory
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
  cat("Created output directory:", output_dir, "\n")
} else {
  cat("Using existing directory:", output_dir, "\n")
}

# Function to read summary data from the Summary sheet
read_summary_data <- function(file_path) {
  tryCatch({
    # Check if Summary sheet exists
    existing_sheets <- excel_sheets(file_path)
    if (!"Summary" %in% existing_sheets) {
      stop("Summary sheet not found in the Excel file!")
    }
    
    data <- read_xlsx(file_path, sheet = "Summary")
    
    if (nrow(data) < 6 || ncol(data) < 2) {
      stop("Summary sheet does not have expected structure!")
    }
    
    # Extract the data - assuming structure is:
    # Row 1: Num Cells, Row 2: Num Positive, Row 3: Num 1+, Row 4: Num 2+, Row 5: Num 3+, Row 6: Num Negative
    
    # Get column names (should be: Classification, Control, T1D, T2D, Aab, Grand Total)
    col_names <- names(data)
    condition_cols <- col_names[2:(length(col_names)-1)]  # Exclude first (Classification) and last (Grand Total)
    
    # Extract cell counts and positive counts for each condition
    num_cells_row <- data[data$Classification == "Num Cells", ]
    num_1plus_row <- data[data$Classification == "Num 1+", ]
    num_2plus_row <- data[data$Classification == "Num 2+", ]
    num_3plus_row <- data[data$Classification == "Num 3+", ]
    
    # Create the result data frame
    result_data <- data.frame()
    
    for (condition in condition_cols) {
      if (condition %in% names(num_cells_row) && 
          !is.na(num_cells_row[[condition]]) && 
          num_cells_row[[condition]] > 0) {
        
        total_cells <- as.numeric(num_cells_row[[condition]])
        cells_1plus <- as.numeric(num_1plus_row[[condition]])
        cells_2plus <- as.numeric(num_2plus_row[[condition]])
        cells_3plus <- as.numeric(num_3plus_row[[condition]])
        
        # Calculate percentages
        perc_1plus <- (cells_1plus / total_cells) * 100
        perc_2plus <- (cells_2plus / total_cells) * 100
        perc_3plus <- (cells_3plus / total_cells) * 100
        
        # Add to result
        condition_data <- data.frame(
          Group = rep(condition, 3),
          Classification = c("1+", "2+", "3+"),
          Percentage = c(perc_1plus, perc_2plus, perc_3plus),
          Total_Cells = rep(total_cells, 3),
          Count = c(cells_1plus, cells_2plus, cells_3plus)
        )
        
        result_data <- rbind(result_data, condition_data)
      }
    }
    
    return(result_data)
    
  }, error = function(e) {
    stop(paste("Error reading Summary sheet:", e$message))
  })
}

# Function to read individual sample data from each sheet for error bars
read_individual_data <- function(file_path) {
  tryCatch({
    existing_sheets <- excel_sheets(file_path)
    available_sheets <- existing_sheets[existing_sheets != "Summary"]
    
    if (length(available_sheets) == 0) {
      return(NULL)
    }
    
    all_individual_data <- data.frame()
    
    for (sheet in available_sheets) {
      data <- tryCatch({
        read_xlsx(file_path, sheet = sheet)
      }, error = function(e) NULL)
      
      if (!is.null(data) && nrow(data) >= 9 && ncol(data) >= 2) {
        # Extract proportion data (rows 7-9 correspond to Prop_1+, Prop_2+, Prop_3+)
        prop_1_plus <- as.numeric(data[7, -1])
        prop_2_plus <- as.numeric(data[8, -1])
        prop_3_plus <- as.numeric(data[9, -1])
        
        # Remove NA values
        prop_1_plus <- prop_1_plus[!is.na(prop_1_plus)] * 100  # Convert to percentage
        prop_2_plus <- prop_2_plus[!is.na(prop_2_plus)] * 100
        prop_3_plus <- prop_3_plus[!is.na(prop_3_plus)] * 100
        
        n_samples <- length(prop_1_plus)
        
        if (n_samples > 0) {
          sheet_data <- data.frame(
            Group = rep(sheet, n_samples * 3),
            Classification = rep(c("1+", "2+", "3+"), each = n_samples),
            Percentage = c(prop_1_plus, prop_2_plus, prop_3_plus),
            Sample_ID = rep(1:n_samples, 3)
          )
          all_individual_data <- rbind(all_individual_data, sheet_data)
        }
      }
    }
    
    return(all_individual_data)
    
  }, error = function(e) {
    cat("Warning: Could not read individual sample data for error bars:", e$message, "\n")
    return(NULL)
  })
}

# Function to calculate maximum values for error bars
calculate_error_bars <- function(summary_data, individual_data) {
  if (is.null(individual_data)) {
    # If no individual data, set error bars to the summary value (no error bars)
    summary_data$Max_Value <- summary_data$Percentage
    return(summary_data)
  }
  
  # Calculate maximum percentage for each group and classification
  max_values <- individual_data %>%
    group_by(Group, Classification) %>%
    summarise(Max_Value = max(Percentage, na.rm = TRUE), .groups = 'drop')
  
  # Merge with summary data
  summary_with_errors <- merge(summary_data, max_values, 
                              by = c("Group", "Classification"), 
                              all.x = TRUE)
  
  # If no max value found, use the summary percentage
  summary_with_errors$Max_Value[is.na(summary_with_errors$Max_Value)] <- 
    summary_with_errors$Percentage[is.na(summary_with_errors$Max_Value)]
  
  return(summary_with_errors)
}

# Read summary data
summary_data <- read_summary_data(path1)

if (nrow(summary_data) == 0) {
  stop("No valid data found in Summary sheet!")
}

# Read individual sample data for error bars
individual_data <- read_individual_data(path1)

# Calculate error bars (maximum values)
summary_data <- calculate_error_bars(summary_data, individual_data)

# Ensure proper factor levels
available_groups <- unique(summary_data$Group)
summary_data$Group <- factor(summary_data$Group, levels = available_groups)
summary_data$Classification <- factor(summary_data$Classification, levels = c("1+", "2+", "3+"))

# Perform statistical tests
perform_statistical_tests <- function(data) {
  groups <- unique(data$Group)
  classifications <- unique(data$Classification)
  
  proportion_test_results <- data.frame()
  chisquare_results <- data.frame()
  
  # Define comparison pairs
  comparison_pairs <- list()
  
  # Control vs other groups
  for (group in groups) {
    if (group != "Control") {
      comparison_pairs[[paste("Control vs", group)]] <- c("Control", group)
    }
  }
  
  # T1D vs other non-Control groups
  if ("T1D" %in% groups && "T2D" %in% groups) {
    comparison_pairs[["T1D vs T2D"]] <- c("T1D", "T2D")
  }
  if ("T1D" %in% groups && "Aab" %in% groups) {
    comparison_pairs[["T1D vs Aab"]] <- c("T1D", "Aab")
  }
  
  # Perform Chi-square tests for each classification if there are 2 or more groups
  if (length(groups) >= 2) {
    cat("Performing Chi-square tests...\n")
    
    for (class in classifications) {
      class_data <- data[data$Classification == class, ]
      
      # Check if we have data for at least 2 groups
      groups_with_data <- unique(class_data$Group)
      if (length(groups_with_data) >= 2) {
        
        # Create contingency table for chi-square test
        # Rows: Groups, Columns: Positive vs Negative cells
        contingency_table <- matrix(0, nrow = length(groups_with_data), ncol = 2)
        rownames(contingency_table) <- groups_with_data
        colnames(contingency_table) <- c("Positive", "Negative")
        
        for (i in seq_along(groups_with_data)) {
          group <- groups_with_data[i]
          group_data <- class_data[class_data$Group == group, ]
          
          if (nrow(group_data) > 0) {
            positive_count <- group_data$Count[1]
            total_count <- group_data$Total_Cells[1]
            negative_count <- total_count - positive_count
            
            contingency_table[i, 1] <- positive_count
            contingency_table[i, 2] <- negative_count
          }
        }
        
        # Perform chi-square test
        chisquare_result <- tryCatch({
          chi_test <- chisq.test(contingency_table)
          p_value <- chi_test$p.value
          chi_statistic <- chi_test$statistic
          df <- chi_test$parameter
          
          significance <- if (p_value < 0.001) "***" else if (p_value < 0.01) "**" else if (p_value < 0.05) "*" else ""
          
          data.frame(
            Classification = class,
            Comparison = paste(groups_with_data, collapse = " vs "),
            Test = "Chi-square test",
            Proportion_1 = NA,
            Proportion_2 = NA,
            Chi_statistic = chi_statistic,
            df = df,
            p_value = p_value,
            significance = significance,
            Sample_1 = "",
            Sample_2 = "",
            Groups_tested = paste(groups_with_data, collapse = ", "),
            stringsAsFactors = FALSE
          )
        }, error = function(e) {
          data.frame(
            Classification = class,
            Comparison = paste(groups_with_data, collapse = " vs "),
            Test = "Chi-square test",
            Proportion_1 = NA,
            Proportion_2 = NA,
            Chi_statistic = NA,
            df = NA,
            p_value = NA,
            significance = "",
            Sample_1 = "",
            Sample_2 = "",
            Groups_tested = paste(groups_with_data, collapse = ", "),
            stringsAsFactors = FALSE
          )
        })
        
        chisquare_results <- rbind(chisquare_results, chisquare_result)
      }
    }
  }
  
  # Perform pairwise two-proportion z-tests
  for (class in classifications) {
    class_data <- data[data$Classification == class, ]
    
    for (comparison_name in names(comparison_pairs)) {
      group1 <- comparison_pairs[[comparison_name]][1]
      group2 <- comparison_pairs[[comparison_name]][2]
      
      group1_data <- class_data[class_data$Group == group1, ]
      group2_data <- class_data[class_data$Group == group2, ]
      
      if (nrow(group1_data) == 0 || nrow(group2_data) == 0) {
        next
      }
      
      # Extract counts for proportion test
      x1 <- group1_data$Count[1]  # Positive cells in group 1
      n1 <- group1_data$Total_Cells[1]  # Total cells in group 1
      x2 <- group2_data$Count[1]  # Positive cells in group 2
      n2 <- group2_data$Total_Cells[1]  # Total cells in group 2
      
      # Perform two-proportion z-test
      proportion_result <- tryCatch({
        test <- prop.test(c(x1, x2), c(n1, n2), alternative = "two.sided", correct = FALSE)
        p_value <- test$p.value
        chi_statistic <- test$statistic
        
        # Calculate proportions
        prop1 <- x1 / n1
        prop2 <- x2 / n2
        
        significance <- if (p_value < 0.001) "***" else if (p_value < 0.01) "**" else if (p_value < 0.05) "*" else ""
        
        data.frame(
          Classification = class,
          Comparison = comparison_name,
          Test = "Two-proportion z-test",
          Proportion_1 = prop1,
          Proportion_2 = prop2,
          Chi_statistic = chi_statistic,
          df = 1,  # Two-proportion test always has 1 degree of freedom
          p_value = p_value,
          significance = significance,
          Sample_1 = paste0(x1, "/", n1),
          Sample_2 = paste0(x2, "/", n2),
          Groups_tested = paste(c(group1, group2), collapse = ", "),
          stringsAsFactors = FALSE
        )
      }, error = function(e) {
        data.frame(
          Classification = class,
          Comparison = comparison_name,
          Test = "Two-proportion z-test",
          Proportion_1 = NA,
          Proportion_2 = NA,
          Chi_statistic = NA,
          df = 1,
          p_value = NA,
          significance = "",
          Sample_1 = paste0(x1, "/", n1),
          Sample_2 = paste0(x2, "/", n2),
          Groups_tested = paste(c(group1, group2), collapse = ", "),
          stringsAsFactors = FALSE
        )
      })
      
      proportion_test_results <- rbind(proportion_test_results, proportion_result)
    }
  }
  
  return(list(
    proportion = proportion_test_results,
    chisquare = chisquare_results,
    combined = rbind(proportion_test_results, chisquare_results)
  ))
}

# Perform statistical tests
stat_results <- perform_statistical_tests(summary_data)

# Create bar plot with total counts and percentages
p <- ggplot(summary_data, aes(x = Group, y = Percentage, fill = Group)) +
  geom_col(alpha = 0.7, width = 0.7, color = "black", size = 0.5) +
  geom_errorbar(aes(ymin = Percentage, ymax = Max_Value), 
                width = 0.3, 
                color = "black", 
                size = 0.8,
                alpha = 0.8) +
  facet_wrap(~ Classification, scales = "free_y") +
  scale_fill_manual(values = c("Control" = "lightblue", 
                               "T1D" = "lightcoral", 
                               "T2D" = "khaki3", 
                               "Aab" = "lightgreen")) +
  scale_x_discrete(labels = c("Control" = var1,
                              "T1D" = var2, 
                              "T2D" = var3,
                              "Aab" = var4)) +
  labs(
    x = "",
    y = "miR-155 expression (%)",
    fill = fill_cust,
    title = title_cust
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    plot.title = element_text(hjust = 0.5),
    legend.position = "bottom"
  )

# Add significance annotations for each classification using proportion test results
for (class in c("1+", "2+", "3+")) {
  # Check if proportion results exist and have the expected structure
  if (is.null(stat_results$proportion) || nrow(stat_results$proportion) == 0) {
    next
  }
  
  class_proportion <- stat_results$proportion[stat_results$proportion$Classification == class & 
                                             stat_results$proportion$significance != "", ]
  
  if (is.null(class_proportion) || nrow(class_proportion) == 0) {
    next
  }
  
  comparisons_list <- list()
  annotations_list <- c()
  
  for (i in 1:nrow(class_proportion)) {
    comparison <- class_proportion$Comparison[i]
    significance <- class_proportion$significance[i]
    
    if (grepl("Control vs", comparison)) {
      group1 <- "Control"
      group2 <- sub("Control vs ", "", comparison)
    } else if (grepl("T1D vs", comparison)) {
      parts <- strsplit(comparison, " vs ")[[1]]
      group1 <- parts[1]
      group2 <- parts[2]
    } else {
      next
    }
    
    if (group1 %in% available_groups && group2 %in% available_groups) {
      comparisons_list[[length(comparisons_list) + 1]] <- c(group1, group2)
      annotations_list <- c(annotations_list, significance)
    }
  }
  
  if (length(comparisons_list) > 0) {
    p <- p + 
      geom_signif(
        data = subset(summary_data, Classification == class),
        comparisons = comparisons_list,
        annotations = annotations_list,
        step_increase = 0.1,
        tip_length = 0.02,
        textsize = 3
      )
  }
}

print(p)


# Save results
output_png_path <- file.path(output_dir, "summary_barplot.png")
ggsave(output_png_path, device = "png", width = 1200, height = 800,
       units = c("px"), create.dir = FALSE, plot = p, bg = "white")

# Create Excel workbook with summary results
wb <- createWorkbook()
addWorksheet(wb, "Summary_Data")
addWorksheet(wb, "Proportion_Test_Results")
addWorksheet(wb, "ChiSquare_Results")
addWorksheet(wb, "Statistical_Tests")

# Write data to sheets
writeData(wb, "Summary_Data", summary_data)

# Write statistical results to sheets
if (!is.null(stat_results$proportion) && nrow(stat_results$proportion) > 0) {
  writeData(wb, "Proportion_Test_Results", stat_results$proportion)
} else {
  writeData(wb, "Proportion_Test_Results", data.frame(Note = "No two-proportion z-test results available"))
}

if (!is.null(stat_results$chisquare) && nrow(stat_results$chisquare) > 0) {
  writeData(wb, "ChiSquare_Results", stat_results$chisquare)
} else {
  writeData(wb, "ChiSquare_Results", data.frame(Note = "No chi-square test results available"))
}

# Combine all statistical tests into one sheet
if (!is.null(stat_results$combined) && nrow(stat_results$combined) > 0) {
  writeData(wb, "Statistical_Tests", stat_results$combined)
} else {
  writeData(wb, "Statistical_Tests", data.frame(Note = "No statistical test results available"))
}

output_xlsx_path <- file.path(output_dir, "Summary_BarPlot_Results.xlsx")
saveWorkbook(wb, output_xlsx_path, overwrite = TRUE)

cat("\nSummary bar plot results saved to:", output_dir, "\n")
cat("Files created:\n")
cat("- summary_barplot.png\n")
cat("- Summary_BarPlot_Results.xlsx\n")

cat("\n", rep("=", 60), "\n", sep = "")
cat("STATISTICAL RESULTS SUMMARY")
cat("\n", rep("=", 60), "\n", sep = "")

cat("\nTWO-PROPORTION Z-TEST RESULTS:\n")
if (!is.null(stat_results$proportion) && nrow(stat_results$proportion) > 0) {
  print(stat_results$proportion)
} else {
  cat("No two-proportion z-test results available\n")
}

cat("\nCHI-SQUARE TEST RESULTS:\n")
if (!is.null(stat_results$chisquare) && nrow(stat_results$chisquare) > 0) {
  print(stat_results$chisquare)
} else {
  cat("No chi-square test results available\n")
}

cat("\n", rep("=", 60), "\n", sep = "")
