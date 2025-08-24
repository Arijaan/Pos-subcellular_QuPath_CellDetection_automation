library(ggplot2)
library(readxl)
library(dplyr)
library(tidyr)

cat("Select the Combined_Summary.xlsx file...\n")
path1 <- file.choose()

# Function to read and process data from each sheet
read_sheet_data <- function(file_path, sheet_name) {
  tryCatch({
    # Read the sheet
    data <- read_xlsx(file_path, sheet = sheet_name)
    
    # Check if data exists and has the right structure
    if (nrow(data) < 7 || ncol(data) < 2) {
      cat(paste("Warning: Sheet", sheet_name, "doesn't have enough data\n"))
      return(NULL)
    }
    
    # Extract proportion data (rows 7-9 correspond to Prop_1+, Prop_2+, Prop_3+)
    prop_1_plus <- as.numeric(data[7, -1])  # Row 7, excluding first column (Count)
    prop_2_plus <- as.numeric(data[8, -1])  # Row 8
    prop_3_plus <- as.numeric(data[9, -1])  # Row 9
    
    # Remove any NA values
    prop_1_plus <- prop_1_plus[!is.na(prop_1_plus)]
    prop_2_plus <- prop_2_plus[!is.na(prop_2_plus)]
    prop_3_plus <- prop_3_plus[!is.na(prop_3_plus)]
    
    # Create a long format data frame
    n_samples <- length(prop_1_plus)
    
    if (n_samples > 0) {
      result <- data.frame(
        Group = rep(sheet_name, n_samples * 3),
        Classification = rep(c("1+", "2+", "3+"), each = n_samples),
        Percentage = c(prop_1_plus, prop_2_plus, prop_3_plus),
        Sample_ID = rep(1:n_samples, 3)
      )
      
      cat(paste("Successfully read", n_samples, "samples from", sheet_name, "sheet\n"))
      return(result)
    } else {
      cat(paste("Warning: No valid data found in", sheet_name, "sheet\n"))
      return(NULL)
    }
    
  }, error = function(e) {
    cat(paste("Error reading sheet", sheet_name, ":", e$message, "\n"))
    return(NULL)
  })
}

# Read data from all available sheets
available_sheets <- c("Control", "T1D", "T2D", "Aab")
all_data <- list()

for (sheet in available_sheets) {
  data <- read_sheet_data(path1, sheet)
  if (!is.null(data)) {
    all_data[[sheet]] <- data
  }
}

# Check if we have any data
if (length(all_data) == 0) {
  stop("No valid data found in any sheets!")
}

# Combine all data
combined_data <- do.call(rbind, all_data)

# Convert Group to factor with proper levels
combined_data$Group <- factor(combined_data$Group, levels = available_sheets)

# Convert Classification to factor with proper order
combined_data$Classification <- factor(combined_data$Classification, 
                                     levels = c("1+", "2+", "3+"))

cat("\nData summary:\n")
print(table(combined_data$Group, combined_data$Classification))

# Perform statistical tests (both Wilcoxon and t-tests for multiple comparisons)
perform_statistical_tests <- function(data) {
  groups <- unique(data$Group)
  classifications <- unique(data$Classification)
  
  wilcox_results <- data.frame()
  ttest_results <- data.frame()
  
  # Define comparison pairs
  comparison_pairs <- list()
  
  # Control vs other groups
  for (group in groups) {
    if (group != "Control") {
      comparison_pairs[[paste("Control vs", group)]] <- c("Control", group)
    }
  }
  
  # T1D vs T2D and T1D vs Aab
  if ("T1D" %in% groups && "T2D" %in% groups) {
    comparison_pairs[["T1D vs T2D"]] <- c("T1D", "T2D")
  }
  if ("T1D" %in% groups && "Aab" %in% groups) {
    comparison_pairs[["T1D vs Aab"]] <- c("T1D", "Aab")
  }
  
  for (class in classifications) {
    class_data <- data[data$Classification == class, ]
    
    for (comparison_name in names(comparison_pairs)) {
      group1 <- comparison_pairs[[comparison_name]][1]
      group2 <- comparison_pairs[[comparison_name]][2]
      
      group1_data <- class_data[class_data$Group == group1, "Percentage"]
      group2_data <- class_data[class_data$Group == group2, "Percentage"]
      
      if (length(group1_data) == 0 || length(group2_data) == 0) {
        cat(paste("Warning: Insufficient data for", comparison_name, "in", class, "\n"))
        next
      }
      
      # Perform Wilcoxon test
      wilcox_result <- tryCatch({
        test <- wilcox.test(group1_data, group2_data, alternative = "two.sided")
        p_value <- test$p.value
        
        # Determine significance
        if (p_value < 0.001) {
          significance <- "***"
        } else if (p_value < 0.01) {
          significance <- "**"
        } else if (p_value < 0.05) {
          significance <- "*"
        } else {
          significance <- ""
        }
        
        data.frame(
          Classification = class,
          Comparison = comparison_name,
          Test = "Wilcoxon",
          p_value = p_value,
          significance = significance,
          stringsAsFactors = FALSE
        )
      }, error = function(e) {
        data.frame(
          Classification = class,
          Comparison = comparison_name,
          Test = "Wilcoxon",
          p_value = NA,
          significance = "",
          stringsAsFactors = FALSE
        )
      })
      
      # Perform Two-Sample t-test
      ttest_result <- tryCatch({
        test <- t.test(group1_data, group2_data, alternative = "two.sided", var.equal = FALSE)
        p_value <- test$p.value
        
        # Determine significance
        if (p_value < 0.001) {
          significance <- "***"
        } else if (p_value < 0.01) {
          significance <- "**"
        } else if (p_value < 0.05) {
          significance <- "*"
        } else {
          significance <- ""
        }
        
        data.frame(
          Classification = class,
          Comparison = comparison_name,
          Test = "t-test",
          p_value = p_value,
          significance = significance,
          stringsAsFactors = FALSE
        )
      }, error = function(e) {
        data.frame(
          Classification = class,
          Comparison = comparison_name,
          Test = "t-test",
          p_value = NA,
          significance = "",
          stringsAsFactors = FALSE
        )
      })
      
      wilcox_results <- rbind(wilcox_results, wilcox_result)
      ttest_results <- rbind(ttest_results, ttest_result)
    }
  }
  
  # Combine results
  all_results <- rbind(wilcox_results, ttest_results)
  
  return(list(
    wilcox = wilcox_results,
    ttest = ttest_results,
    combined = all_results
  ))
}

# Perform statistical tests
stat_results <- perform_statistical_tests(combined_data)


# Prepare significance annotations for plot (using Wilcoxon results for Control vs other comparisons)
if (nrow(stat_results$wilcox) > 0) {
  # Filter for Control vs other comparisons only for plot annotations
  control_comparisons <- stat_results$wilcox[grepl("Control vs", stat_results$wilcox$Comparison), ]
  
  if (nrow(control_comparisons) > 0) {
    # Extract group names from comparison strings
    control_comparisons$Group <- sub("Control vs ", "", control_comparisons$Comparison)
    
    # Calculate y positions for significance labels
    max_y <- combined_data %>%
      group_by(Classification, Group) %>%
      summarise(max_val = max(Percentage, na.rm = TRUE), .groups = "drop")
    
    significance_data <- control_comparisons %>%
      filter(significance != "") %>%
      left_join(max_y, by = c("Classification", "Group")) %>%
      mutate(y_pos = max_val * 1.1)
  } else {
    significance_data <- data.frame()
  }
} else {
  significance_data <- data.frame()
}

# Create the boxplot
p <- ggplot(combined_data, aes(x = Classification, y = Percentage, fill = Group)) +
  geom_boxplot(position = position_dodge(width = 0.8), alpha = 0.7) +
  scale_fill_manual(values = c("Control" = "lightblue", 
                               "T1D" = "lightcoral", 
                               "T2D" = "khaki3", 
                               "Aab" = "lightgreen")) +
  labs(
    x = "Cell Classification",
    y = "miR-155 expression(%)",
    fill = "Condition"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    plot.title = element_text(hjust = 0.5),
    legend.position = "bottom"
  )


print(p)

# Print data summary again

print(table(combined_data$Group, combined_data$Classification))

# Print statistical results
if (nrow(stat_results$combined) > 0) {
  cat("STATISTICAL TEST RESULTS (All Comparisons)\n")
  cat(rep("=", 60), "\n")
  cat("* p < 0.05, ** p < 0.01, *** p < 0.001\n\n")
  
  cat("WILCOXON RANK-SUM TEST RESULTS:\n")
  cat(rep("-", 40), "\n")
  if (nrow(stat_results$wilcox) > 0) {
    print(stat_results$wilcox)
  } else {
    cat("No Wilcoxon test results available.\n")
  }
  
  cat("\nTWO-SAMPLE T-TEST RESULTS:\n")
  cat(rep("-", 40), "\n")
  if (nrow(stat_results$ttest) > 0) {
    print(stat_results$ttest)
  } else {
    cat("No t-test results available.\n")
  }
  
  cat("\n", rep("=", 60), "\n")
}
