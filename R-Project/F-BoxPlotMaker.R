library(ggplot2)
library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)
library(ggsignif)

cat("Select summary file containing cell counts...\n")
path1 <- file.choose()

# Ask for folder name
cat("Enter the name for the output folder (will be created under 'plots and stats'):\n")
folder_name <-"Ductal_donors"

# enter plot xlab names and legend name

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

base_output_dir <- "C:/Users/psoor/OneDrive/Desktop/Bachelorarbeit/Figures/plots and stats"
output_dir <- file.path(base_output_dir, folder_name)


# Create output directory
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
  cat("Created output directory:", output_dir, "\n")
} else {
  cat("Using existing directory:", output_dir, "\n")
}

# Function to read and process data from each sheet
read_sheet_data <- function(file_path, sheet_name) {
  tryCatch({
    data <- read_xlsx(file_path, sheet = sheet_name)
    
    if (nrow(data) < 7 || ncol(data) < 2) {
      return(NULL)
    }
    
    # Extract proportion data (rows 7-9 correspond to Prop_1+, Prop_2+, Prop_3+)
    prop_1_plus <- as.numeric(data[7, -1])
    prop_2_plus <- as.numeric(data[8, -1])
    prop_3_plus <- as.numeric(data[9, -1])
    
    # Remove NA values
    prop_1_plus <- prop_1_plus[!is.na(prop_1_plus)]
    prop_2_plus <- prop_2_plus[!is.na(prop_2_plus)]
    prop_3_plus <- prop_3_plus[!is.na(prop_3_plus)]
    
    n_samples <- length(prop_1_plus)
    
    if (n_samples > 0) {
      return(data.frame(
        Group = rep(sheet_name, n_samples * 3),
        Classification = rep(c("1+", "2+", "3+"), each = n_samples),
        Percentage = c(prop_1_plus, prop_2_plus, prop_3_plus),
        Sample_ID = rep(1:n_samples, 3)
      ))
    } else {
      return(NULL)
    }
  }, error = function(e) {
    return(NULL)
  })
}

# Read data from available sheets
all_possible_sheets <- c("Control", "T1D", "T2D", "Aab")
existing_sheets <- tryCatch(excel_sheets(path1), error = function(e) character(0))
available_sheets <- intersect(all_possible_sheets, existing_sheets)

if (!"Control" %in% available_sheets) {
  stop("Control sheet is required but not found in the Excel file!")
}

all_data <- list()
for (sheet in available_sheets) {
  data <- read_sheet_data(path1, sheet)
  if (!is.null(data)) {
    all_data[[sheet]] <- data
  }
}

if (length(all_data) == 0) {
  stop("No valid data found in any sheets!")
}

# Combine and prepare data
combined_data <- do.call(rbind, all_data)
combined_data$Group <- factor(combined_data$Group, levels = available_sheets)
combined_data$Classification <- factor(combined_data$Classification, levels = c("1+", "2+", "3+"))

# Perform statistical tests
perform_statistical_tests <- function(data) {
  groups <- unique(data$Group)
  classifications <- unique(data$Classification)
  
  wilcox_results <- data.frame()
  ttest_results <- data.frame()
  anova_results <- data.frame()
  
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
  
  # Perform ANOVA if there are 3 or more groups
  if (length(groups) >= 3) {
    cat("Performing ANOVA tests (3+ groups detected)...\n")
    
    for (class in classifications) {
      class_data <- data[data$Classification == class, ]
      
      # Check if we have data for at least 3 groups
      groups_with_data <- unique(class_data$Group)
      if (length(groups_with_data) >= 3) {
        
        anova_result <- tryCatch({
          # Perform one-way ANOVA
          aov_test <- aov(Percentage ~ Group, data = class_data)
          aov_summary <- summary(aov_test)
          p_value <- aov_summary[[1]][["Pr(>F)"]][1]
          
          significance <- if (p_value < 0.001) "***" else if (p_value < 0.01) "**" else if (p_value < 0.05) "*" else ""
          
          data.frame(
            Classification = class,
            Test = "One-way ANOVA",
            F_statistic = aov_summary[[1]][["F value"]][1],
            df_between = aov_summary[[1]][["Df"]][1],
            df_within = aov_summary[[1]][["Df"]][2],
            p_value = p_value,
            significance = significance,
            Groups_tested = paste(groups_with_data, collapse = ", "),
            stringsAsFactors = FALSE
          )
        }, error = function(e) {
          data.frame(
            Classification = class,
            Test = "One-way ANOVA",
            F_statistic = NA,
            df_between = NA,
            df_within = NA,
            p_value = NA,
            significance = "",
            Groups_tested = paste(groups_with_data, collapse = ", "),
            stringsAsFactors = FALSE
          )
        })
        
        anova_results <- rbind(anova_results, anova_result)
      }
    }
  } else {
    cat("ANOVA not performed: Less than 3 groups detected\n")
  }
  
  for (class in classifications) {
    class_data <- data[data$Classification == class, ]
    
    for (comparison_name in names(comparison_pairs)) {
      group1 <- comparison_pairs[[comparison_name]][1]
      group2 <- comparison_pairs[[comparison_name]][2]
      
      group1_data <- class_data[class_data$Group == group1, "Percentage"]
      group2_data <- class_data[class_data$Group == group2, "Percentage"]
      
      if (length(group1_data) == 0 || length(group2_data) == 0) {
        next
      }
      
      # Perform Wilcoxon test
      wilcox_result <- tryCatch({
        test <- wilcox.test(group1_data, group2_data, alternative = "two.sided")
        p_value <- test$p.value
        
        significance <- if (p_value < 0.001) "***" else if (p_value < 0.01) "**" else if (p_value < 0.05) "*" else ""
        
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
        
        significance <- if (p_value < 0.001) "***" else if (p_value < 0.01) "**" else if (p_value < 0.05) "*" else ""
        
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
  
  return(list(
    wilcox = wilcox_results,
    ttest = ttest_results,
    anova = anova_results,
    combined = rbind(wilcox_results, ttest_results)
  ))
}

stat_results <- perform_statistical_tests(combined_data)

# Create boxplot with significance annotations
groups <- levels(combined_data$Group)

p <- ggplot(combined_data, aes(x = Group, y = Percentage, fill = Group)) +
  geom_boxplot(alpha = 0.7) +
  facet_wrap(~ Classification, scales = "free_y", strip.position = "bottom") +
  scale_fill_manual(values = c("Control" = "lightblue2", 
                               "T1D" = "lightcoral", 
                               "T2D" = "khaki3", 
                               "Aab" = "lightgreen"),
                    labels = c("Control" = var1,
                               "T1D" = var2, 
                               "T2D" = var3,
                               "Aab" = var4)) +
  labs(
    x = "",
    y = "miR-155 expression(%)",
    fill = fill_cust,
    title = title_cust
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    plot.title = element_text(hjust = 0.5),
    legend.position = "bottom",
    strip.placement = "outside",
    strip.text = element_text(size = 10)
  )

# Add significance annotations for each classification
for (class in c("1+", "2+", "3+")) {
  class_ttest <- stat_results$ttest[stat_results$ttest$Classification == class & 
                                   stat_results$ttest$significance != "", ]
  
  if (nrow(class_ttest) > 0) {
    comparisons_list <- list()
    
    for (i in 1:nrow(class_ttest)) {
      comparison <- class_ttest$Comparison[i]
      
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
      
      if (group1 %in% groups && group2 %in% groups) {
        comparisons_list[[length(comparisons_list) + 1]] <- c(group1, group2)
      }
    }
    
    if (length(comparisons_list) > 0) {
      p <- p + 
        geom_signif(
          data = subset(combined_data, Classification == class),
          comparisons = comparisons_list,
          test = "t.test",
          map_signif_level = c("***" = 0.001, "**" = 0.01, "*" = 0.05),
          step_increase = 0.1,
          tip_length = 0.02
        )
    }
  }
}

# Add sample size annotations
sample_counts <- table(combined_data$Group, combined_data$Classification)
count_data <- data.frame()

for (group in rownames(sample_counts)) {
  for (class in colnames(sample_counts)) {
    if (sample_counts[group, class] > 0) {
      count_data <- rbind(count_data, data.frame(
        Group = group,
        Classification = class,
        count = sample_counts[group, class],
        stringsAsFactors = FALSE
      ))
    }
  }
}

p <- p + 
  geom_text(
    data = count_data,
    aes(x = Group, y = -Inf, label = paste0("n=", count)),
    vjust = -0.5,
    size = 3,
    color = "black",
    inherit.aes = FALSE
  )

print(p)


# Save results
output_png_path <- file.path(output_dir, "boxplot.png")
ggsave(output_png_path, device = "png", width = 1200, height = 1200,
       units = c("px"), create.dir = FALSE, plot = p, bg = "white")

# Create Excel workbook with results
wb <- createWorkbook()
addWorksheet(wb, "Wilcoxon_Results")
addWorksheet(wb, "T_test_Results") 
addWorksheet(wb, "ANOVA_Results")
addWorksheet(wb, "Data_Summary")
addWorksheet(wb, "Raw_Data")

# Write results to sheets
if (nrow(stat_results$wilcox) > 0) {
  writeData(wb, "Wilcoxon_Results", stat_results$wilcox)
} else {
  writeData(wb, "Wilcoxon_Results", data.frame(Note = "No Wilcoxon test results available"))
}

if (nrow(stat_results$ttest) > 0) {
  writeData(wb, "T_test_Results", stat_results$ttest)
} else {
  writeData(wb, "T_test_Results", data.frame(Note = "No t-test results available"))
}

if (nrow(stat_results$anova) > 0) {
  writeData(wb, "ANOVA_Results", stat_results$anova)
} else {
  writeData(wb, "ANOVA_Results", data.frame(Note = "No ANOVA results available (requires 3+ groups)"))
}

data_summary_table <- as.data.frame.matrix(table(combined_data$Group, combined_data$Classification))
data_summary_table$Group <- rownames(data_summary_table)
data_summary_table <- data_summary_table[, c("Group", "1+", "2+", "3+")]
writeData(wb, "Data_Summary", data_summary_table)
writeData(wb, "Raw_Data", combined_data)

output_xlsx_path <- file.path(output_dir, "Statistical_Results.xlsx")
saveWorkbook(wb, output_xlsx_path, overwrite = TRUE)

# Copy the input Combined_Summary.xlsx file to the output directory
input_copy_path <- file.path(output_dir, paste0("INPUT_", basename(path1)))
file.copy(path1, input_copy_path, overwrite = TRUE)

# POWER ANALYSIS
cat("\n", rep("=", 60), "\n", sep = "")
cat("POWER ANALYSIS - SAMPLE SIZE CALCULATION")
cat("\n", rep("=", 60), "\n", sep = "")

# Install pwr package if not available
if (!require(pwr, quietly = TRUE)) {
  install.packages("pwr")
  library(pwr)
}

# Function to perform power analysis for t-test
perform_power_analysis <- function(control_data, comparison_data, group_name) {
  # Calculate effect size (Cohen's d)
  control_mean <- mean(control_data, na.rm = TRUE)
  control_sd <- sd(control_data, na.rm = TRUE)
  comparison_mean <- mean(comparison_data, na.rm = TRUE)
  comparison_sd <- sd(comparison_data, na.rm = TRUE)
  
  # Pooled standard deviation
  pooled_sd <- sqrt(((length(control_data) - 1) * control_sd^2 + 
                     (length(comparison_data) - 1) * comparison_sd^2) / 
                    (length(control_data) + length(comparison_data) - 2))
  
  # Cohen's d effect size
  cohens_d <- abs(control_mean - comparison_mean) / pooled_sd
  
  # Power analysis for 80% power (0.80)
  power_result <- pwr.t.test(d = cohens_d, 
                            power = 0.80, 
                            sig.level = 0.05, 
                            type = "two.sample")
  
  # Current sample sizes
  current_n_control <- length(control_data)
  current_n_comparison <- length(comparison_data)
  
  # Calculate current power with existing sample sizes
  current_power <- pwr.t.test(n = min(current_n_control, current_n_comparison),
                             d = cohens_d,
                             sig.level = 0.05,
                             type = "two.sample")$power
  
  return(list(
    group = group_name,
    cohens_d = cohens_d,
    effect_size_interpretation = case_when(
      cohens_d < 0.2 ~ "Very Small",
      cohens_d < 0.5 ~ "Small", 
      cohens_d < 0.8 ~ "Medium",
      TRUE ~ "Large"
    ),
    required_n_per_group = ceiling(power_result$n),
    current_n_control = current_n_control,
    current_n_comparison = current_n_comparison,
    current_power = round(current_power, 3),
    control_mean = round(control_mean, 2),
    comparison_mean = round(comparison_mean, 2),
    control_sd = round(control_sd, 2),
    comparison_sd = round(comparison_sd, 2)
  ))
}

# Perform power analysis for each comparison group
power_results <- list()

# Get control data
control_data <- combined_data[combined_data$Group == "Control", "Percentage"]

cat("\nUsing Control group as baseline for power analysis:")
cat("\n- Control mean:", round(mean(control_data, na.rm = TRUE), 2), "%")
cat("\n- Control SD:", round(sd(control_data, na.rm = TRUE), 2), "%")
cat("\n- Control n:", length(control_data))

cat("\n\nPOWER ANALYSIS RESULTS (for 80% power, α = 0.05):\n")
cat(rep("-", 80), "\n", sep = "")

for (group in names(all_data)) {
  if (group != "Control") {
    comparison_data <- combined_data[combined_data$Group == group, "Percentage"]
    
    if (length(comparison_data) > 0) {
      power_analysis <- perform_power_analysis(control_data, comparison_data, group)
      power_results[[group]] <- power_analysis
      
      cat("\n", group, " vs Control:\n", sep = "")
      cat("  Effect Size (Cohen's d):", sprintf("%.3f", power_analysis$cohens_d), 
          "(", power_analysis$effect_size_interpretation, ")\n")
      cat("  Required sample size per group:", power_analysis$required_n_per_group, "\n")
      cat("  Current sample sizes: Control =", power_analysis$current_n_control, 
          ",", group, "=", power_analysis$current_n_comparison, "\n")
      cat("  Current statistical power:", sprintf("%.1f%%", power_analysis$current_power * 100), "\n")
      cat("  Means: Control =", power_analysis$control_mean, "%, ", 
          group, " =", power_analysis$comparison_mean, "%\n")
      
      # Additional sample size recommendations
      additional_needed_control <- max(0, power_analysis$required_n_per_group - power_analysis$current_n_control)
      additional_needed_comparison <- max(0, power_analysis$required_n_per_group - power_analysis$current_n_comparison)
      
      if (additional_needed_control > 0 || additional_needed_comparison > 0) {
        cat("  Additional samples needed: Control =", additional_needed_control, 
            ",", group, "=", additional_needed_comparison, "\n")
      } else {
        cat("  ✓ Adequate sample size for 80% power\n")
      }
    }
  }
}

# Create power analysis summary table
if (length(power_results) > 0) {
  power_summary <- data.frame(
    Comparison = paste(names(power_results), "vs Control"),
    Effect_Size_Cohens_d = sapply(power_results, function(x) round(x$cohens_d, 3)),
    Effect_Size_Category = sapply(power_results, function(x) x$effect_size_interpretation),
    Required_N_per_Group = sapply(power_results, function(x) x$required_n_per_group),
    Current_N_Control = sapply(power_results, function(x) x$current_n_control),
    Current_N_Comparison = sapply(power_results, function(x) x$current_n_comparison),
    Current_Power = sapply(power_results, function(x) paste0(round(x$current_power * 100, 1), "%")),
    Adequate_Power = sapply(power_results, function(x) ifelse(x$current_power >= 0.80, "Yes", "No")),
    stringsAsFactors = FALSE
  )
  
  cat("\n\nPOWER ANALYSIS SUMMARY TABLE:\n")
  print(power_summary)
  
  # Add power analysis to Excel file
  addWorksheet(wb, "Power_Analysis")
  writeData(wb, "Power_Analysis", power_summary)
  
  # Add interpretation guide
  interpretation_guide <- data.frame(
    Effect_Size_d = c("< 0.2", "0.2 - 0.5", "0.5 - 0.8", "> 0.8"),
    Interpretation = c("Very Small", "Small", "Medium", "Large"),
    stringsAsFactors = FALSE
  )
  
  writeData(wb, "Power_Analysis", interpretation_guide, startRow = nrow(power_summary) + 4)
  writeData(wb, "Power_Analysis", "Cohen's d Effect Size Interpretation:", 
            startRow = nrow(power_summary) + 3)
  
  # Save updated workbook
  saveWorkbook(wb, output_xlsx_path, overwrite = TRUE)
}


cat("\nResults saved to folder:", folder_name, "\n")
cat("Full path:", output_dir, "\n")
cat("Files created:\n")
cat("- boxplot.png\n")
cat("- Statistical_Results.xlsx (now includes Power Analysis sheet)\n")
cat("- INPUT_", basename(path1), "\n", sep = "")

cat("\n", rep("=", 60), "\n", sep = "")
cat("STATISTICAL RESULTS SUMMARY")
cat("\n", rep("=", 60), "\n", sep = "")

cat("\nWILCOXON TEST RESULTS:\n")
if (nrow(stat_results$wilcox) > 0) {
  print(stat_results$wilcox)
} else {
  cat("No Wilcoxon test results available\n")
}

cat("\nT-TEST RESULTS:\n")
if (nrow(stat_results$ttest) > 0) {
  print(stat_results$ttest)
} else {
  cat("No t-test results available\n")
}

cat("\nANOVA RESULTS:\n")
if (nrow(stat_results$anova) > 0) {
  print(stat_results$anova)
} else {
  cat("No ANOVA results available (requires 3+ groups)\n")
}

cat("\n", rep("=", 60), "\n", sep = "")

# ADDITIONAL POWER ANALYSIS FOR EACH COMPARISON
cat("\n", rep("=", 60), "\n", sep = "")
cat("DETAILED POWER ANALYSIS FOR EACH COMPARISON AND CLASSIFICATION")
cat("\n", rep("=", 60), "\n", sep = "")

# Install effsize package if not available for Cohen's d calculation
if (!require(effsize, quietly = TRUE)) {
  install.packages("effsize")
  library(effsize)
}

# Function to perform and display power analysis for specific classification
perform_classification_power_analysis <- function(control_data, treatment_data, comparison_name, classification) {
  cat("\n", rep("-", 60), "\n", sep = "")
  cat("POWER ANALYSIS:", comparison_name, "- Classification:", classification, "\n")
  cat(rep("-", 60), "\n", sep = "")
  
  if (length(control_data) == 0 || length(treatment_data) == 0) {
    cat("Insufficient data for this classification\n")
    return(NULL)
  }
  
  # Calculate Cohen's d using effsize package
  d_value <- tryCatch({
    cohen.d(control_data, treatment_data)
  }, error = function(e) {
    # Fallback calculation if cohen.d fails
    control_mean <- mean(control_data, na.rm = TRUE)
    control_sd <- sd(control_data, na.rm = TRUE)
    treatment_mean <- mean(treatment_data, na.rm = TRUE)
    treatment_sd <- sd(treatment_data, na.rm = TRUE)
    
    # Pooled standard deviation
    pooled_sd <- sqrt(((length(control_data) - 1) * control_sd^2 + 
                       (length(treatment_data) - 1) * treatment_sd^2) / 
                      (length(control_data) + length(treatment_data) - 2))
    
    cohens_d <- (treatment_mean - control_mean) / pooled_sd
    
    # Create similar structure to cohen.d output
    list(estimate = cohens_d, 
         magnitude = case_when(
           abs(cohens_d) < 0.2 ~ "negligible",
           abs(cohens_d) < 0.5 ~ "small", 
           abs(cohens_d) < 0.8 ~ "medium",
           TRUE ~ "large"
         ))
  })
  
  cat("Cohen's d effect size:", round(d_value$estimate, 4), "\n")
  cat("Effect size magnitude:", d_value$magnitude, "\n")
  
  # Calculate descriptive statistics
  cat("\nDescriptive Statistics:")
  cat("\n- Control: Mean =", round(mean(control_data, na.rm = TRUE), 2), 
      "%, SD =", round(sd(control_data, na.rm = TRUE), 2), "%, n =", length(control_data))
  cat("\n- Treatment: Mean =", round(mean(treatment_data, na.rm = TRUE), 2), 
      "%, SD =", round(sd(treatment_data, na.rm = TRUE), 2), "%, n =", length(treatment_data))
  
  # Perform power analysis with error handling
  power_result <- tryCatch({
    # Check if effect size is extremely large
    effect_size <- abs(d_value$estimate)
    if (effect_size > 5) {
      cat("\n\nNote: Effect size is extremely large (d =", round(effect_size, 2), ")")
      cat("\nWith such a large effect, very small sample sizes are adequate for 80% power.")
      cat("\nRequired sample size per group: ≤ 3")
      
      # Calculate current power manually for very large effects
      current_power <- pwr.t.test(n = min(length(control_data), length(treatment_data)),
                                 d = effect_size,
                                 sig.level = 0.05,
                                 type = "two.sample")$power
      
      list(n = 3, d = effect_size, note = "Very large effect")
    } else {
      # Normal power analysis
      pwr.t.test(d = effect_size, 
                sig.level = 0.05, 
                power = 0.80, 
                type = "two.sample", 
                alternative = "two.sided")
    }
  }, error = function(e) {
  cat("\n\nError in power calculation:", e$message)
    cat("\nThis typically occurs with extremely large or small effect sizes.")
    
    # Provide alternative calculation
    effect_size <- abs(d_value$estimate)
    if (effect_size > 3) {
      cat("\nWith effect size d =", round(effect_size, 2), ", very small samples are adequate.")
      list(n = 3, d = effect_size, note = "Error - likely very large effect")
    } else {
      cat("\nUnable to calculate required sample size.")
      list(n = NA, d = effect_size, note = "Calculation error")
    }
  })
  
  if (!is.null(power_result) && !is.na(power_result$n)) {
    if (is.null(power_result$note)) {
      cat("\n\nPower Analysis Results:\n")
      print(power_result)
    }
    
    required_n <- ceiling(power_result$n)
  } else {
    required_n <- NA
  }
  
  # Current sample sizes and power
  current_power <- tryCatch({
    pwr.t.test(n = min(length(control_data), length(treatment_data)),
               d = abs(d_value$estimate),
               sig.level = 0.05,
               type = "two.sample")$power
  }, error = function(e) {
    if (abs(d_value$estimate) > 3) {
      0.99  # Very high power for large effects
    } else {
      NA
    }
  })
  
  if (!is.na(current_power)) {
    cat("\nCurrent power with existing sample sizes:", round(current_power * 100, 1), "%")
  } else {
    cat("\nCurrent power: Unable to calculate")
  }
  
  # Sample size recommendations
  if (!is.na(required_n)) {
    additional_control <- max(0, required_n - length(control_data))
    additional_treatment <- max(0, required_n - length(treatment_data))
    
    if (additional_control > 0 || additional_treatment > 0) {
      cat("\nAdditional samples needed:")
      cat("\n- Control: +", additional_control, "samples")
      cat("\n- Treatment: +", additional_treatment, "samples")
    } else {
      cat("\n✓ Adequate sample size for 80% power")
    }
  } else {
    cat("\nSample size recommendations: Unable to calculate")
  }
  
  cat("\n")
  
  return(list(
    comparison = comparison_name,
    classification = classification,
    cohens_d = d_value$estimate,
    effect_magnitude = d_value$magnitude,
    required_n = ifelse(is.na(required_n), "Unable to calculate", required_n),
    current_power = ifelse(is.na(current_power), NA, current_power),
    control_n = length(control_data),
    treatment_n = length(treatment_data),
    control_mean = mean(control_data, na.rm = TRUE),
    treatment_mean = mean(treatment_data, na.rm = TRUE)
  ))
}

# Perform power analysis for each comparison and classification
all_power_results <- list()
classifications <- c("1+", "2+", "3+")

# Control vs T1D
if ("T1D" %in% names(all_data)) {
  cat("\n", rep("=", 80), "\n", sep = "")
  cat("T1D vs Control")
  cat("\n", rep("=", 80), "\n", sep = "")
  
  for (class in classifications) {
    control_data <- combined_data[combined_data$Group == "Control" & combined_data$Classification == class, "Percentage"]
    t1d_data <- combined_data[combined_data$Group == "T1D" & combined_data$Classification == class, "Percentage"]
    
    result <- perform_classification_power_analysis(control_data, t1d_data, "T1D vs Control", class)
    if (!is.null(result)) {
      all_power_results[[paste("T1D_vs_Control", class, sep = "_")]] <- result
    }
  }
}

# Control vs T2D
if ("T2D" %in% names(all_data)) {
  cat("\n", rep("=", 80), "\n", sep = "")
  cat("T2D vs Control")
  cat("\n", rep("=", 80), "\n", sep = "")
  
  for (class in classifications) {
    control_data <- combined_data[combined_data$Group == "Control" & combined_data$Classification == class, "Percentage"]
    t2d_data <- combined_data[combined_data$Group == "T2D" & combined_data$Classification == class, "Percentage"]
    
    result <- perform_classification_power_analysis(control_data, t2d_data, "T2D vs Control", class)
    if (!is.null(result)) {
      all_power_results[[paste("T2D_vs_Control", class, sep = "_")]] <- result
    }
  }
}

# Control vs Aab
if ("Aab" %in% names(all_data)) {
  cat("\n", rep("=", 80), "\n", sep = "")
  cat("Aab vs Control")
  cat("\n", rep("=", 80), "\n", sep = "")
  
  for (class in classifications) {
    control_data <- combined_data[combined_data$Group == "Control" & combined_data$Classification == class, "Percentage"]
    aab_data <- combined_data[combined_data$Group == "Aab" & combined_data$Classification == class, "Percentage"]
    
    result <- perform_classification_power_analysis(control_data, aab_data, "Aab vs Control", class)
    if (!is.null(result)) {
      all_power_results[[paste("Aab_vs_Control", class, sep = "_")]] <- result
    }
  }
}

# Create comprehensive summary table
if (length(all_power_results) > 0) {
  cat("\n", rep("=", 100), "\n", sep = "")
  cat("COMPREHENSIVE POWER ANALYSIS SUMMARY")
  cat("\n", rep("=", 100), "\n", sep = "")
  
  summary_df <- data.frame(
    Comparison = sapply(all_power_results, function(x) x$comparison),
    Classification = sapply(all_power_results, function(x) x$classification),
    Cohens_d = sapply(all_power_results, function(x) round(x$cohens_d, 3)),
    Effect_Size = sapply(all_power_results, function(x) x$effect_magnitude),
    Required_N_per_Group = sapply(all_power_results, function(x) x$required_n),
    Current_N_Control = sapply(all_power_results, function(x) x$control_n),
    Current_N_Treatment = sapply(all_power_results, function(x) x$treatment_n),
    Current_Power_Percent = sapply(all_power_results, function(x) round(x$current_power * 100, 1)),
    Control_Mean = sapply(all_power_results, function(x) round(x$control_mean, 2)),
    Treatment_Mean = sapply(all_power_results, function(x) round(x$treatment_mean, 2)),
    Adequate_Power = sapply(all_power_results, function(x) ifelse(x$current_power >= 0.80, "Yes", "No")),
    stringsAsFactors = FALSE
  )
  
  print(summary_df)
  
  # Add detailed power analysis to Excel file
  addWorksheet(wb, "Detailed_Power_Analysis")
  writeData(wb, "Detailed_Power_Analysis", summary_df)
  
  # Add classification-specific interpretation
  classification_note <- data.frame(
    Note = c("1+ = Low expression cells", 
             "2+ = Medium expression cells", 
             "3+ = High expression cells",
             "",
             "Power Analysis Target: 80% power, α = 0.05"),
    stringsAsFactors = FALSE
  )
  
  writeData(wb, "Detailed_Power_Analysis", classification_note, startRow = nrow(summary_df) + 3)
  
  # Save updated workbook
  saveWorkbook(wb, output_xlsx_path, overwrite = TRUE)
}












