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
folder_name <-"test"

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



cat("\nResults saved to folder:", folder_name, "\n")
cat("Full path:", output_dir, "\n")
cat("Files created:\n")
cat("- boxplot.png\n")
cat("- Statistical_Results.xlsx\n")
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


