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
cat("Select the T2D CSV file...\n")
path3 <- file.choose()


# Read specific columns (column 1 or 2. Either 5,6 or 7,8 or 8,9)
c1 <- 5
c2 <- 6
control <- read.csv(path1)[, c(c1, c2)]
t1d <- read.csv(path2)[, c(c1, c2)]
t2d <- read.csv(path3)[, c(c1, c2)]

# DEBUG: Clean and convert the data to numeric
control[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(control[, 2])))
t1d[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(t1d[, 2])))
t2d[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(t2d[, 2])))

# DEBUG: Remove rows with NA values
control <- control[complete.cases(control), ]
t1d <- t1d[complete.cases(t1d), ]
t2d <- t2d[complete.cases(t2d), ]

# Add group labels
control$group <- "Control"
t1d$group <- "T1D"
t2d$group <- "T2D"

# Combine datasets
combined_data <- rbind(control, t1d, t2d)

# Handle data reshaping
if(ncol(combined_data) == 3) {
  colnames(combined_data)[1:2] <- c("classification", "value")
  combined_data_long <- combined_data
} else {
  control_wide <- control %>% select(-group)
  t1d_wide <- t1d %>% select(-group)
  t2d_wide <- t2d %>% select(-group)
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

  t2d_long <- t2d_wide %>%
    mutate(sample_id = row_number()) %>%
    pivot_longer(cols = all_of(classification_cols),
                 names_to = "classification",
                 values_to = "value") %>%
    mutate(group = "T2D") %>%
    select(-sample_id)

  combined_data_long <- rbind(control_long, t1d_long, t2d_long)
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
  t2d_total_raw <- read.csv(path3)[2, c2]

  control_total <- as.numeric(gsub("[^0-9.-]", "", as.character(control_total_raw)))
  t1d_total <- as.numeric(gsub("[^0-9.-]", "", as.character(t1d_total_raw)))
  t2d_total <- as.numeric(gsub("[^0-9.-]", "", as.character(t2d_total_raw)))

  total_counts <- data.frame(
    group = c("Control", "T1D", "T2D"),
    total = c(control_total, t1d_total, t2d_total)
  )
  
  proportion_data <- combined_data_long %>%
    mutate(value = as.numeric(as.character(value))) %>%
    filter(!is.na(value)) %>%
    merge(total_counts, by = "group") %>%
    mutate(proportion = value / total)
}

# Convert proportions to percentages for boxplot
boxplot_data <- proportion_data %>%
  mutate(percentage = proportion * 100)

# Filter and reorder data for plotting (excluding Negative)
classification_order <- c("Num 1+", "Num 2+", "Num 3+")

filtered_boxplot_data <- boxplot_data %>%
  filter(classification %in% classification_order) %>%
  mutate(classification = factor(classification, levels = classification_order, labels = c("1+","2+","3+")))

# Perform statistical tests for significance annotation
perform_wilcox_test <- function(data, class_name) {
  class_data <- data %>% filter(classification == class_name)
  
  # Test Control vs T1D and T2D
  ctrl_data <- class_data %>% filter(group == "Control") %>% pull(percentage)
  t1d_data <- class_data %>% filter(group == "T1D") %>% pull(percentage)
  t2d_data <- class_data %>% filter(group == "T2D") %>% pull(percentage)
  
  # Initialize empty results data frame
  results <- data.frame(
    classification = character(),
    group = character(),
    p_value = numeric(),
    significance = character(),
    stringsAsFactors = FALSE
  )
  
  if(length(ctrl_data) > 0 && length(t1d_data) > 0) {
    p_val_t1d <- tryCatch({
      wilcox.test(t1d_data, ctrl_data, alternative = "two.sided")$p.value
    }, error = function(e) NA)
    
    sig_t1d <- ifelse(is.na(p_val_t1d), "", 
                      ifelse(p_val_t1d < 0.001, "***", 
                             ifelse(p_val_t1d < 0.01, "**", 
                                    ifelse(p_val_t1d < 0.05, "*", ""))))
    
    results <- rbind(results, data.frame(
      classification = class_name,
      group = "T1D",
      p_value = p_val_t1d,
      significance = sig_t1d,
      stringsAsFactors = FALSE
    ))
  }
  
  if(length(ctrl_data) > 0 && length(t2d_data) > 0) {
    p_val_t2d <- tryCatch({
      wilcox.test(t2d_data, ctrl_data, alternative = "two.sided")$p.value
    }, error = function(e) NA)
    
    sig_t2d <- ifelse(is.na(p_val_t2d), "", 
                      ifelse(p_val_t2d < 0.001, "***", 
                             ifelse(p_val_t2d < 0.01, "**", 
                                    ifelse(p_val_t2d < 0.05, "*", ""))))
    
    results <- rbind(results, data.frame(
      classification = class_name,
      group = "T2D",
      p_value = p_val_t2d,
      significance = sig_t2d,
      stringsAsFactors = FALSE
    ))
  }
  
  return(results)
}

# Calculate statistical tests for each classification
stat_results <- do.call(rbind, lapply(levels(filtered_boxplot_data$classification), 
                                      function(x) perform_wilcox_test(filtered_boxplot_data, x)))

# Calculate y positions for significance labels
max_y <- filtered_boxplot_data %>%
  group_by(classification) %>%
  summarise(max_val = max(percentage, na.rm = TRUE), .groups = "drop")

significance_data <- stat_results %>%
  filter(significance != "" & group != "Control") %>%
  left_join(max_y, by = "classification") %>%
  mutate(y_pos = max_val * 1.1)

# Create boxplot
p <- ggplot(filtered_boxplot_data, aes(x = classification, y = percentage, fill = group)) +
  geom_boxplot(position = position_dodge(width = 0.8), alpha = 0.7) +
  geom_text(data = significance_data,
            aes(x = classification, y = y_pos, label = significance, group = group),
            position = position_dodge(width = 0.8), vjust = 0, size = 4) +
  scale_fill_manual(values = c("Control" = "lightblue", "T1D" = "lightcoral", "T2D" = "khaki3")) +
  labs(
    title = "miR-155 Expression in Pancreatic Cells",
    x = "Cell Classification",
    y = "Percentage (%)",
    fill = "Group"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    plot.title = element_text(hjust = 0.5),
    legend.position = "bottom"
  )

print(p)

# Print statistical test results
cat("\nWilcoxon rank-sum test results (vs Control):\n")
cat("* p < 0.05, ** p < 0.01, *** p < 0.001\n\n")
if(nrow(stat_results) > 0) {
  print(stat_results)
} else {
  cat("No statistical tests performed.\n")
}
