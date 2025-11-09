library(ggplot2)
library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)
library(ggsignif)

cat("Select the Combined_Summary.xlsx file...\n")
path1 <- file.choose()

# Create output directory with the same name as the grandparent directory
input_dir <- dirname(path1)
grandparent_dir <- dirname(input_dir)
grandparent_dir_name <- basename(grandparent_dir)
output_dir <- file.path(input_dir, grandparent_dir_name)

if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
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
        Group = rep(sheet_name, n_samples),
        Prop_1_plus = prop_1_plus,
        Prop_2_plus = prop_2_plus,
        Prop_3_plus = prop_3_plus,
        Sample_ID = 1:n_samples
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
# Calculate weighted expression values for each sample
# 1+ cells = base value (weight = 1)
# 2+ cells = 2 * base value (weight = 2) 
# 3+ cells = 3 * base value (weight = 3)

weighted_data <- combined_data %>%
  mutate(
    # Calculate weighted expression value for each sample
    # This represents the intensity-weighted expression level
    weighted_expression = (1 * Prop_1_plus + 2 * Prop_2_plus + 3 * Prop_3_plus),
    
    # Also calculate individual classification percentages for comparison
    Classification_1plus = Prop_1_plus,
    Classification_2plus = Prop_2_plus,
    Classification_3plus = Prop_3_plus
  ) %>%
  select(Group, Sample_ID, weighted_expression, Classification_1plus, Classification_2plus, Classification_3plus)

# Create long format data for individual classifications
individual_data <- weighted_data %>%
  pivot_longer(
    cols = c(Classification_1plus, Classification_2plus, Classification_3plus),
    names_to = "Classification",
    values_to = "Percentage"
  ) %>%
  mutate(
    Classification = case_when(
      Classification == "Classification_1plus" ~ "1+",
      Classification == "Classification_2plus" ~ "2+",
      Classification == "Classification_3plus" ~ "3+",
      TRUE ~ Classification
    )
  ) %>%
  mutate(Classification = factor(Classification, levels = c("1+", "2+", "3+")))

# Create weighted data for plotting
weighted_plot_data <- weighted_data %>%
  select(Group, Sample_ID, weighted_expression) %>%
  mutate(Classification = "Weighted") %>%
  rename(Percentage = weighted_expression)

# Combine weighted and individual data
all_plot_data <- bind_rows(
  individual_data %>% select(Group, Sample_ID, Classification, Percentage),
  weighted_plot_data
) %>%
  mutate(Classification = factor(Classification, levels = c("1+", "2+", "3+", "Weighted")))
# Perform statistical tests
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
  
  # T1D vs other non-Control groups
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
    combined = rbind(wilcox_results, ttest_results)
  ))
}

stat_results <- perform_statistical_tests(all_plot_data)

# Create boxplot with significance annotations
groups <- levels(all_plot_data$Group)

p <- ggplot(all_plot_data, aes(x = Group, y = Percentage, fill = Group)) +
  geom_boxplot(alpha = 0.7) +
  facet_wrap(~ Classification, scales = "free_y") +
  scale_fill_manual(values = c("Control" = "lightblue", 
                               "T1D" = "lightcoral", 
                               "T2D" = "khaki3", 
                               "Aab" = "lightgreen")) +
  scale_x_discrete(labels = c("Control" = "Small ducts",
                              "T1D" = "Big ducts", 
                              "T2D" = "T2D",
                              "Aab" = "Aab")) +
  labs(
    x = "Condition",
    y = "miR-155 expression(%)",
    fill = "Condition",
    title = "miR-155 Expression: Individual Classifications + Weighted Analysis"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    plot.title = element_text(hjust = 0.5),
    legend.position = "bottom"
  )

# Add significance annotations for each classification
for (class in levels(all_plot_data$Classification)) {
  class_wilcox <- stat_results$wilcox[stat_results$wilcox$Classification == class & 
                                      stat_results$wilcox$significance != "", ]
  
  if (nrow(class_wilcox) > 0) {
    comparisons_list <- list()
    
    for (i in 1:nrow(class_wilcox)) {
      comparison <- class_wilcox$Comparison[i]
      
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
          data = subset(all_plot_data, Classification == class),
          comparisons = comparisons_list,
          test = "wilcox.test",
          map_signif_level = c("***" = 0.001, "**" = 0.01, "*" = 0.05),
          step_increase = 0.1,
          tip_length = 0.02
        )
    }
  }
}

# Add sample size annotations
sample_counts <- all_plot_data %>%
  group_by(Group, Classification) %>%
  summarise(n = n(), .groups = "drop")

p <- p + 
  geom_text(
    data = sample_counts,
    aes(x = Group, y = -Inf, label = paste0("n=", n)),
    vjust = -0.5,
    size = 3,
    color = "black",
    inherit.aes = FALSE
  )

print(p)

# Save results
output_png_path <- file.path(output_dir, "weighted_boxplot.png")
ggsave(output_png_path, device = "png", width = 1400, height = 1000,
       units = c("px"), create.dir = FALSE, plot = p, bg = "white")

# Create Excel workbook with results
wb <- createWorkbook()
addWorksheet(wb, "Wilcoxon_Results")
addWorksheet(wb, "T_test_Results") 
addWorksheet(wb, "Weighted_Data")
addWorksheet(wb, "Individual_Data")
addWorksheet(wb, "Raw_Proportions")

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

# Write weighted data
writeData(wb, "Weighted_Data", weighted_data)

# Write individual classification data
writeData(wb, "Individual_Data", individual_data)

# Write raw proportion data
writeData(wb, "Raw_Proportions", combined_data)

output_xlsx_path <- file.path(output_dir, "Weighted_Statistical_Results.xlsx")
saveWorkbook(wb, output_xlsx_path, overwrite = TRUE)

cat("\nWeighted Expression Analysis Complete!\n")
cat("=====================================\n")
cat("This analysis shows:\n")
cat("- Individual 1+, 2+, 3+ classifications as separate boxplots\n")
cat("- Weighted expression combining all levels (1×1+ + 2×2+ + 3×3+)\n")
cat("- Statistical comparisons for each classification\n")
cat("- Sample sizes for each group\n")
cat("\nResults saved to:", output_dir, "\n")
cat("- Plot: weighted_boxplot.png\n")
cat("- Data: Weighted_Statistical_Results.xlsx\n")
