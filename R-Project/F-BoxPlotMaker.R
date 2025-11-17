library(ggplot2)
library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)
library(ggsignif)
library(svDialogs)
library(pwr)

cat("Select summary file containing cell counts...\n")
path1 <- file.choose()

# Ask for folder name
cat("Enter the name for the output folder (will be created under 'plots and stats'):\n")
folder_name <- dlg_input("Name of the output folder")$res
if (is.null(folder_name) || folder_name == "") {
  folder_name <- paste0("analysis_", format(Sys.time(), "%Y%m%d_%H%M%S"))
}

# enter plot xlab names and legend name

var1 <- "Small ducs"
var2 <- "Big ducts"
var3 <- "T2D"
var4 <- "Aab"
fill_cust <- ""
title_cust <- ""

# Clean the folder name (remove invalid characters)
folder_name <- gsub("[^a-zA-Z0-9_-]", "_", folder_name)

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
    prop_rows <- suppressWarnings(lapply(7:9, function(idx) as.numeric(data[idx, -1])))
    prop_rows <- lapply(prop_rows, function(values) values[!is.na(values)])
    n_samples <- min(lengths(prop_rows))
    
    if (n_samples > 0) {
      prop_rows <- lapply(prop_rows, function(values) values[seq_len(n_samples)])
      return(data.frame(
        Group = rep(sheet_name, n_samples * 3),
        Classification = rep(c("1+", "2+", "3+"), each = n_samples),
        Percentage = unlist(prop_rows, use.names = FALSE),
        Sample_ID = rep(seq_len(n_samples), 3)
      ))
    }
    NULL
  }, error = function(e) NULL)
}

load_all_group_data <- function(file_path, sheets) {
  raw_list <- lapply(sheets, function(sheet) read_sheet_data(file_path, sheet))
  names(raw_list) <- sheets
  valid_idx <- !sapply(raw_list, is.null)
  if (!any(valid_idx)) {
    return(list(combined = NULL, raw = list(), available = character(0)))
  }
  filtered_raw <- raw_list[valid_idx]
  available <- names(filtered_raw)
  combined <- bind_rows(filtered_raw)
  combined$Group <- factor(combined$Group, levels = available)
  combined$Classification <- factor(combined$Classification, levels = c("1+", "2+", "3+"))
  list(combined = combined, raw = filtered_raw, available = available)
}

# Read data from available sheets
all_possible_sheets <- c("Control", "T1D", "T2D", "Aab")
cached_sheets <- tryCatch(excel_sheets(path1), error = function(e) character(0))
available_sheets <- intersect(all_possible_sheets, cached_sheets)

if (!"Control" %in% available_sheets) {
  stop("Control sheet is required but not found in the Excel file!")
}

loaded_groups <- load_all_group_data(path1, available_sheets)

if (is.null(loaded_groups$combined)) {
  stop("No valid data found in any sheets!")
}

combined_data <- loaded_groups$combined
all_data <- loaded_groups$raw
groups <- levels(combined_data$Group)

significance_label <- function(p_value) {
  if (is.na(p_value)) {
    ""
  } else if (p_value < 0.001) {
    "***"
  } else if (p_value < 0.01) {
    "**"
  } else if (p_value < 0.05) {
    "*"
  } else {
    ""
  }
}

generate_comparison_pairs <- function(group_levels) {
  pairs <- list()
  non_control <- setdiff(group_levels, "Control")
  for (group in non_control) {
    pairs[[paste("Control vs", group)]] <- c("Control", group)
  }
  if (all(c("T1D", "T2D") %in% group_levels)) {
    pairs[["T1D vs T2D"]] <- c("T1D", "T2D")
  }
  if (all(c("T1D", "Aab") %in% group_levels)) {
    pairs[["T1D vs Aab"]] <- c("T1D", "Aab")
  }
  pairs
}

compute_pairwise_result <- function(test_fun, test_name, group1_data, group2_data, classification, comparison_name) {
  tryCatch({
    test <- test_fun(group1_data, group2_data)
    p_value <- test$p.value
    data.frame(
      Classification = classification,
      Comparison = comparison_name,
      Test = test_name,
      p_value = p_value,
      significance = significance_label(p_value),
      stringsAsFactors = FALSE
    )
  }, error = function(e) {
    data.frame(
      Classification = classification,
      Comparison = comparison_name,
      Test = test_name,
      p_value = NA,
      significance = "",
      stringsAsFactors = FALSE
    )
  })
}

unique_pairs <- function(pairs_list) {
  if (!length(pairs_list)) {
    return(list())
  }
  keys <- vapply(pairs_list, function(pair) paste(pair, collapse = "::"), character(1))
  pairs_list[!duplicated(keys)]
}

build_significance_pairs <- function(ttest_results, comparison_pairs, valid_groups) {
  if (!nrow(ttest_results)) {
    return(list())
  }
  significant <- ttest_results[ttest_results$significance != "", , drop = FALSE]
  if (!nrow(significant)) {
    return(list())
  }
  pair_map <- split(significant, significant$Classification)
  pair_map <- lapply(pair_map, function(df) {
    pairs <- lapply(df$Comparison, function(comparison) comparison_pairs[[comparison]])
    pairs <- Filter(function(pair) !is.null(pair) && all(pair %in% valid_groups), pairs)
    unique_pairs(pairs)
  })
  Filter(length, pair_map)
}

ensure_package <- function(package_name) {
  if (!require(package_name, character.only = TRUE, quietly = TRUE)) {
    install.packages(package_name)
    library(package_name, character.only = TRUE)
  }
}

print_banner <- function(title, width = 60, symbol = "=") {
  line <- paste(rep(symbol, width), collapse = "")
  cat("\n", line, "\n", title, "\n", line, "\n", sep = "")
}

print_subbanner <- function(title, width = 60, symbol = "-") {
  line <- paste(rep(symbol, width), collapse = "")
  cat("\n", line, "\n", title, "\n", line, "\n", sep = "")
}

print_rule <- function(width = 60, symbol = "=") {
  cat("\n", paste(rep(symbol, width), collapse = ""), "\n", sep = "")
}

# Perform statistical tests
perform_statistical_tests <- function(data) {
  group_levels <- levels(droplevels(data$Group))
  classifications <- levels(droplevels(data$Classification))
  comparison_pairs <- generate_comparison_pairs(group_levels)
  
  wilcox_results <- list()
  ttest_results <- list()
  anova_results <- list()
  
  if (length(group_levels) >= 3) {
    cat("Performing ANOVA tests (3+ groups detected)...\n")
    for (class in classifications) {
      class_data <- data[data$Classification == class, ]
      groups_with_data <- levels(droplevels(class_data$Group))
      if (length(groups_with_data) >= 3) {
        anova_entry <- tryCatch({
          aov_test <- aov(Percentage ~ Group, data = class_data)
          aov_summary <- summary(aov_test)
          p_value <- aov_summary[[1]][["Pr(>F)"]][1]
          data.frame(
            Classification = class,
            Test = "One-way ANOVA",
            F_statistic = aov_summary[[1]][["F value"]][1],
            df_between = aov_summary[[1]][["Df"]][1],
            df_within = aov_summary[[1]][["Df"]][2],
            p_value = p_value,
            significance = significance_label(p_value),
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
        anova_results[[length(anova_results) + 1]] <- anova_entry
      }
    }
  } else {
    cat("ANOVA not performed: Less than 3 groups detected\n")
  }
  
  if (length(comparison_pairs) > 0) {
    wilcox_fun <- function(x, y) wilcox.test(x, y, alternative = "two.sided")
    ttest_fun <- function(x, y) t.test(x, y, alternative = "two.sided", var.equal = FALSE)
    for (class in classifications) {
      class_data <- data[data$Classification == class, ]
      for (comparison_name in names(comparison_pairs)) {
        groups_pair <- comparison_pairs[[comparison_name]]
        group1_data <- class_data[class_data$Group == groups_pair[1], "Percentage"]
        group2_data <- class_data[class_data$Group == groups_pair[2], "Percentage"]
        if (length(group1_data) == 0 || length(group2_data) == 0) {
          next
        }
        wilcox_results[[length(wilcox_results) + 1]] <-
          compute_pairwise_result(wilcox_fun, "Wilcoxon", group1_data, group2_data, class, comparison_name)
        ttest_results[[length(ttest_results) + 1]] <-
          compute_pairwise_result(ttest_fun, "t-test", group1_data, group2_data, class, comparison_name)
      }
    }
  }
  
  wilcox_df <- if (length(wilcox_results)) bind_rows(wilcox_results) else data.frame()
  ttest_df <- if (length(ttest_results)) bind_rows(ttest_results) else data.frame()
  anova_df <- if (length(anova_results)) bind_rows(anova_results) else data.frame()
  
  list(
    wilcox = wilcox_df,
    ttest = ttest_df,
    anova = anova_df,
    combined = bind_rows(wilcox_df, ttest_df),
    comparisons = comparison_pairs
  )
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
significance_mapping <- build_significance_pairs(stat_results$ttest, stat_results$comparisons, groups)
if (length(significance_mapping)) {
  signif_thresholds <- c("***" = 0.001, "**" = 0.01, "*" = 0.05)
  for (class in names(significance_mapping)) {
    p <- p +
      geom_signif(
        data = subset(combined_data, Classification == class),
        comparisons = significance_mapping[[class]],
        test = "t.test",
        map_signif_level = signif_thresholds,
        step_increase = 0.1,
        tip_length = 0.02
      )
  }
}

# Add sample size annotations
sample_counts <- as.data.frame(as.table(table(combined_data$Group, combined_data$Classification)))
colnames(sample_counts) <- c("Group", "Classification", "count")
count_data <- subset(sample_counts, count > 0)
count_data$Group <- factor(count_data$Group, levels = groups)
count_data$Classification <- factor(count_data$Classification, levels = levels(combined_data$Classification))

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
print_banner("POWER ANALYSIS - SAMPLE SIZE CALCULATION")


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

print_subbanner("POWER ANALYSIS RESULTS (for 80% power, α = 0.05):", width = 80)

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

print_banner("STATISTICAL RESULTS SUMMARY")

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

print_rule()

# ADDITIONAL POWER ANALYSIS FOR EACH COMPARISON
print_banner("DETAILED POWER ANALYSIS FOR EACH COMPARISON AND CLASSIFICATION")

# Install effsize package if not available for Cohen's d calculation
ensure_package("effsize")

# Function to perform and display power analysis for specific classification
perform_classification_power_analysis <- function(control_data, treatment_data, comparison_name, classification) {
  print_subbanner(paste("POWER ANALYSIS:", comparison_name, "- Classification:", classification))
  
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

analyze_group_vs_control <- function(group_name, classifications, data, available_groups) {
  if (!(group_name %in% available_groups)) {
    return(list())
  }
  print_banner(paste(group_name, "vs Control"), width = 80)
  group_results <- list()
  for (class in classifications) {
    control_values <- data[data$Group == "Control" & data$Classification == class, "Percentage"]
    treatment_values <- data[data$Group == group_name & data$Classification == class, "Percentage"]
    result <- perform_classification_power_analysis(control_values, treatment_values, paste(group_name, "vs Control"), class)
    if (!is.null(result)) {
      group_results[[paste(group_name, class, sep = "_")]] <- result
    }
  }
  group_results
}

# Perform power analysis for each comparison and classification
all_power_results <- list()
classifications <- levels(combined_data$Classification)
comparison_groups <- c("T1D", "T2D", "Aab")

for (group in comparison_groups) {
  group_results <- analyze_group_vs_control(group, classifications, combined_data, groups)
  if (length(group_results)) {
    all_power_results <- c(all_power_results, group_results)
  }
}

# Create comprehensive summary table
if (length(all_power_results) > 0) {
  print_banner("COMPREHENSIVE POWER ANALYSIS SUMMARY", width = 100)
  
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
             "Power Analysis Target: 80% power, a = 0.05"),
    stringsAsFactors = FALSE
  )
  
  writeData(wb, "Detailed_Power_Analysis", classification_note, startRow = nrow(summary_df) + 3)
  
  # Save updated workbook
  saveWorkbook(wb, output_xlsx_path, overwrite = TRUE)
}












