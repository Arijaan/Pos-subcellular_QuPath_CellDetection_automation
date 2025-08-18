library(ggplot2)
library(dplyr)
library(tidyr)
library(tibble)

# This script is meant to be combined with the 
# "save measures" script in QuPath after cell detection
# Interactive file selection
cat("Select the Control CSV file...\n")
path1 <- file.choose()
cat("Select the T1D CSV file...\n")
path2 <- file.choose()

# Read specific columns (column 1 or 2. Either 5,6 or 7,8 or 8,9)
c1 <- 5
c2 <- 6
control <- read.csv(path1)[, c(c1, c2)]
t1d <- read.csv(path2)[, c(c1, c2)]
t2d <- read.csv(path3)[, c(c1, c2)]
# DEBUG: Clean and convert the data to numeric
control[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(control[, 2])))
t1d[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(t1d[, 2])))

# DEBUG: Remove rows with NA values
control <- control[complete.cases(control), ]
t1d <- t1d[complete.cases(t1d), ]

# Add group labels
control$group <- "Control"
t1d$group <- "T1D"

# Combine datasets
combined_data <- rbind(control, t1d)

# Handle data reshaping
if(ncol(combined_data) == 3) {
  colnames(combined_data)[1:2] <- c("classification", "value")
  combined_data_long <- combined_data
} else {
  control_wide <- control %>% select(-group)
  t1d_wide <- t1d %>% select(-group)
  classification_cols <- colnames(control_wide)
  
  control_long <- control_wide %>%
    mutate(sample_id = row_number()) %>%
    pivot_longer(cols = all_of(classification_cols), 
                 names_to = "classification", 
                 values_to = "value") %>%
    mutate(group = "Control") %>%
    select(-sample_id)
  
  t1d_long <- t1d_wide %>%
    mutate(sample_id = row_number()) %>%
    pivot_longer(cols = all_of(classification_cols), 
                 names_to = "classification", 
                 values_to = "value") %>%
    mutate(group = "T1D") %>%
    select(-sample_id)
  
  combined_data_long <- rbind(control_long, t1d_long)
}

# Calculate totals and proportions
all_classifications <- unique(combined_data_long$classification)

if("all" %in% all_classifications) {
  total_counts <- combined_data_long %>%
    filter(classification == "all") %>%
    select(group, value) %>%
    rename(total = value) %>%
    mutate(total = as.numeric(as.character(total)))
  
  proportion_data <- combined_data_long %>%
    filter(classification != "all") %>%
    mutate(value = as.numeric(as.character(value))) %>%
    merge(total_counts, by = "group", all.x = TRUE) %>%
    filter(!is.na(value) & !is.na(total) & total > 0) %>%
    mutate(proportion = value / total)
} else {
  # Read totals from I2 (row 2, column 9)
  control_total_raw <- read.csv(path1)[2, c2]
  t1d_total_raw <- read.csv(path2)[2, c2]
  
  control_total <- as.numeric(gsub("[^0-9.-]", "", as.character(control_total_raw)))
  t1d_total <- as.numeric(gsub("[^0-9.-]", "", as.character(t1d_total_raw)))
  
  total_counts <- data.frame(
    group = c("Control", "T1D"),
    total = c(control_total, t1d_total)
  )
  
  proportion_data <- combined_data_long %>%
    mutate(value = as.numeric(as.character(value))) %>%
    filter(!is.na(value)) %>%
    merge(total_counts, by = "group") %>%
    mutate(proportion = value / total)
}

# Calculate weighted expression values
# 1+ cells = base value (weight = 1)
# 2+ cells = 2 * base value (weight = 2) 
# 3+ cells = 3 * base value (weight = 3)

# Calculate summary statistics with weighted values
weighted_summary_data <- proportion_data %>%
  group_by(group) %>%
  summarise(
    # Extract individual classification counts
    count_1plus = sum(value[classification == "Num 1+"], na.rm = TRUE),
    count_2plus = sum(value[classification == "Num 2+"], na.rm = TRUE),
    count_3plus = sum(value[classification == "Num 3+"], na.rm = TRUE),
    total = first(total),
    
    # Calculate weighted expression value
    weighted_expression = (1 * count_1plus + 2 * count_2plus + 3 * count_3plus),
    
    # Calculate as proportion of total cells
    weighted_proportion = weighted_expression / total,
    
    # Convert to percentage for display
    weighted_percentage = weighted_proportion * 100,
    
    # Calculate standard error for weighted proportion
    se_weighted = sqrt(weighted_proportion * (1 - weighted_proportion) / total) * 100,
    
    .groups = "drop"
  ) %>%
  # Add a classification column for consistency with plotting functions
  mutate(classification = "Weighted Expression")

# Display the weighted calculation breakdown
cat("Weighted Expression Calculation:\n")
cat("===============================\n")
for(i in 1:nrow(weighted_summary_data)) {
  group_name <- weighted_summary_data$group[i]
  cat("\n", group_name, ":\n")
  cat("  1+ cells:", weighted_summary_data$count_1plus[i], "(weight = 1)\n")
  cat("  2+ cells:", weighted_summary_data$count_2plus[i], "(weight = 2)\n") 
  cat("  3+ cells:", weighted_summary_data$count_3plus[i], "(weight = 3)\n")
  cat("  Weighted sum:", weighted_summary_data$weighted_expression[i], "\n")
  cat("  Total cells:", weighted_summary_data$total[i], "\n")
  cat("  Weighted %:", round(weighted_summary_data$weighted_percentage[i], 2), "%\n")
}

# Perform statistical test on weighted expression
control_data <- weighted_summary_data[weighted_summary_data$group == "Control", ]
t1d_data <- weighted_summary_data[weighted_summary_data$group == "T1D", ]

if (nrow(control_data) > 0 && nrow(t1d_data) > 0) {
  # Use weighted expression values directly for comparison
  control_weighted <- control_data$weighted_expression[1]
  control_total <- control_data$total[1]
  t1d_weighted <- t1d_data$weighted_expression[1]
  t1d_total <- t1d_data$total[1]
  
  # Calculate proportions for statistical test
  control_prop <- control_weighted / control_total
  t1d_prop <- t1d_weighted / t1d_total
  
  tryCatch({
    # Test if T1D has higher weighted expression than Control
    # Using a two-sample t-test for proportions converted to normal approximation
    prop_test <- prop.test(c(t1d_weighted, control_weighted), 
                          c(t1d_total, control_total), 
                          alternative = "greater", 
                          correct = TRUE)
    
    # Determine significance
    primary_p_value <- prop_test$p.value
    if (primary_p_value < 0.001) {
      significance <- "***"
    } else if (primary_p_value < 0.01) {
      significance <- "**"
    } else if (primary_p_value < 0.05) {
      significance <- "*"
    } else {
      significance <- ""
    }
    
    weighted_test_results <- data.frame(
      test = "Weighted Expression",
      control_percentage = round(control_prop * 100, 2),
      t1d_percentage = round(t1d_prop * 100, 2),
      p_value = primary_p_value,
      significance = significance
    )
    
  }, error = function(e) {
    weighted_test_results <<- data.frame(
      test = "Weighted Expression",
      control_percentage = NA,
      t1d_percentage = NA,
      p_value = NA,
      significance = "error"
    )
  })
} else {
  weighted_test_results <- data.frame(
    test = "Weighted Expression",
    control_percentage = NA,
    t1d_percentage = NA,
    p_value = NA,
    significance = "no_data"
  )
}

# Add significance information to weighted summary data
weighted_summary_data$significance <- weighted_test_results$significance[1]
weighted_summary_data$p_value <- weighted_test_results$p_value[1]

# Calculate significance star position
max_y <- max(weighted_summary_data$weighted_percentage + weighted_summary_data$se_weighted) * 1.15

# Create plot for weighted expression
p <- ggplot(weighted_summary_data, aes(x = group, y = weighted_percentage, fill = group)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.8), width = 0.7) +
  geom_errorbar(aes(ymin = weighted_percentage - se_weighted, ymax = weighted_percentage + se_weighted), 
                position = position_dodge(width = 0.8), width = 0.2) +
  geom_text(aes(x = group, y = max_y, label = significance),
            position = position_dodge(width = 0.8), hjust = 0.5, vjust = 0, size = 6) +
  scale_fill_manual(values = c("Control" = "lightblue", "T1D" = "lightcoral")) +
  scale_x_discrete(labels = c("Control" = "ND", "T1D" = "T1D")) +
  labs(
    x = NULL,
    y = "miR-155 expression (%)",
    fill = "Condition"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(size = 12, hjust = 0.5),
    plot.title = element_text(hjust = 0.5, size = 14),
    legend.position = "bottom"
  ) +
  scale_y_continuous(expand = expansion(mult = c(0, 0.2)))

print(p)

# Print results
cat("\nWeighted Expression Analysis Results:\n")
cat("====================================\n")
cat("This analysis combines all positive expression levels into a single weighted measure:\n")
cat("- 1+ cells contribute their count × 1\n")
cat("- 2+ cells contribute their count × 2\n") 
cat("- 3+ cells contribute their count × 3\n")
cat("\nStatistical Test: Tests if T1D has significantly HIGHER weighted expression than Control\n")
cat("Significance levels: *** p<0.001, ** p<0.01, * p<0.05, ns = not significant\n\n")

print(weighted_test_results)

cat("\nDetailed Summary:\n")
