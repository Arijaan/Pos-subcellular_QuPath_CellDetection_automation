library(ggplot2)
library(dplyr)
library(tidyr)
library(tibble)

# Fold Change Analysis Plot (Similar to F-OnlyPosProp but with fold changes)
# This script creates a plot showing fold changes of T1D relative to Control
cat("Select the Control CSV file...\n")
path1 <- file.choose()
cat("Select the T1D CSV file...\n")
path2 <- file.choose()

# Read specific columns (column 1 or 2. Either 5,6 or 7,8 or 8,9)
c1 <- 5
c2 <- 6
control <- read.csv(path1)[, c(c1, c2)]
t1d <- read.csv(path2)[, c(c1, c2)]

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
  # Read totals from row 2
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

# Calculate summary statistics
summary_data <- proportion_data %>%
  group_by(classification, group) %>%
  summarise(
    count = sum(value),
    total = first(total),
    proportion = sum(value) / first(total),
    mean_value = proportion * 100,
    se = sqrt(proportion * (1 - proportion) / first(total)) * 100,
    .groups = "drop"
  )

# Calculate fold changes with Control as baseline
fold_change_data <- summary_data %>%
  select(classification, group, proportion) %>%
  pivot_wider(names_from = group, values_from = proportion) %>%
  mutate(
    fold_change = ifelse(Control > 0, T1D / Control, 
                        ifelse(T1D > 0, Inf, 1)),
    log2_fold_change = ifelse(is.finite(fold_change) & fold_change > 0, 
                             log2(fold_change), NA),
    # For visualization, we'll use log2 fold change but display as fold change
    display_value = ifelse(is.finite(fold_change), fold_change, 
                          ifelse(T1D > 0, 10, 0.1)) # Cap extreme values
  ) %>%
  filter(!is.na(fold_change))

# Perform statistical tests for significance
classifications <- unique(proportion_data$classification)
statistical_results <- data.frame()

for (class in classifications) {
  control_data <- proportion_data[proportion_data$classification == class & proportion_data$group == "Control", ]
  t1d_data <- proportion_data[proportion_data$classification == class & proportion_data$group == "T1D", ]
  
  if (nrow(control_data) > 0 && nrow(t1d_data) > 0) {
    control_count <- control_data$value[1]
    control_total <- control_data$total[1]
    t1d_count <- t1d_data$value[1]
    t1d_total <- t1d_data$total[1]
    
    tryCatch({
      # Two-sided proportion test
      prop_test <- prop.test(c(control_count, t1d_count), c(control_total, t1d_total), 
                            alternative = "two.sided", correct = TRUE)
      
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
      
      statistical_results <- rbind(statistical_results, data.frame(
        classification = class,
        p_value = primary_p_value,
        significance = significance
      ))
    }, error = function(e) {
      statistical_results <<- rbind(statistical_results, data.frame(
        classification = class,
        p_value = NA,
        significance = ""
      ))
    })
  } else {
    statistical_results <- rbind(statistical_results, data.frame(
      classification = class,
      p_value = NA,
      significance = ""
    ))
  }
}

# Merge fold change data with statistical results
plot_data <- merge(fold_change_data, statistical_results, by = "classification", all.x = TRUE)

# Filter and reorder data for plotting (excluding Negative, only positive classifications)
classification_order <- c("Num 1+", "Num 2+", "Num 3+")

filtered_plot_data <- plot_data %>%
  filter(classification %in% classification_order) %>%
  mutate(
    classification = factor(classification, levels = classification_order, 
                           labels = c("1+", "2+", "3+")),
    # Create error bars for fold change (approximate using proportion SE)
    fold_change_se = abs(fold_change) * 0.1, # Approximate SE for visualization
    direction = ifelse(fold_change > 1, "Increased", 
                      ifelse(fold_change < 1, "Decreased", "No Change")),
    # For better visualization, cap extreme fold changes
    plot_fold_change = pmax(pmin(fold_change, 10), 0.1)
  )

# Create the fold change plot (similar style to F-OnlyPosProp)
p <- ggplot(filtered_plot_data, aes(x = classification, y = plot_fold_change)) +
  geom_col(aes(fill = direction), width = 0.7, alpha = 0.8) +
  geom_errorbar(aes(ymin = pmax(plot_fold_change - fold_change_se, 0.01), 
                    ymax = plot_fold_change + fold_change_se), 
                width = 0.2, color = "black") +
  geom_text(aes(y = plot_fold_change + fold_change_se + 0.2, 
                label = significance),
            size = 6, vjust = 0) +
  scale_fill_manual(values = c("Increased" = "lightcoral", 
                              "Decreased" = "lightblue", 
                              "No Change" = "lightgray")) +
  labs(
    title = "Fold Change Analysis: T1D vs Control",
    x = NULL,
    y = "Fold Change (T1D/Control)",
    fill = "Direction in T1D"
  ) +
  

  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1, size = 12),
    axis.text.y = element_text(size = 10),
    axis.title.y = element_text(size = 12),
    plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
    legend.position = "bottom",
    panel.grid.minor = element_blank()
  )

print(p)

# Print fold change results
cat("\nFold Change Results (T1D vs Control):\n")
cat("Values > 1 indicate increase in T1D, values < 1 indicate decrease in T1D\n")
cat("Significance: *** p<0.001, ** p<0.01, * p<0.05\n\n")

results_table <- filtered_plot_data %>%
  select(classification, Control, T1D, fold_change, significance, p_value) %>%
  mutate(
    Control = round(Control * 100, 2),
    T1D = round(T1D * 100, 2),
    fold_change = round(fold_change, 3)
  ) %>%
  rename(
    Classification = classification,
    "Control %" = Control,
    "T1D %" = T1D,
    "Fold Change" = fold_change,
    "Significance" = significance,
    "P-value" = p_value
  )

print(results_table)

# Save results
write.csv(results_table, "fold_change_plot_results.csv", row.names = FALSE)
cat("\nResults saved to: fold_change_plot_results.csv\n")