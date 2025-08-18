library(ggplot2)
library(dplyr)
library(tidyr)
library(tibble)

# Interactive file selection
cat("Please select the Control CSV file...\n")
path1 <- file.choose()
cat("Please select the T1D CSV file...\n")
path2 <- file.choose()

# Read specific columns (H and I - columns 8 and 9)
control <- read.csv(path1)[, c(8, 9)]
t1d <- read.csv(path2)[, c(8, 9)]

# Clean and convert data to numeric
control[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(control[, 2])))
t1d[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(t1d[, 2])))

# Remove rows with NA values
control <- control[complete.cases(control), ]
t1d <- t1d[complete.cases(t1d), ]

# Add group labels
control$group <- "Control"
t1d$group <- "T1D"

# Combine datasets into two columns including all values
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

# Calculate totals
all_classifications <- unique(combined_data_long$classification)

control_total_raw <- read.csv(path1)[2, 9]
t1d_total_raw <- read.csv(path2)[2, 9]

control_total <- as.numeric(gsub("[^0-9.-]", "", as.character(control_total_raw)))
t1d_total <- as.numeric(gsub("[^0-9.-]", "", as.character(t1d_total_raw)))

total_counts <- data.frame(
  group = c("Control", "T1D"),
  total = c(control_total, t1d_total)
)

# Calculate proportions

proportion_data <- combined_data_long %>%
  mutate(value = as.numeric(as.character(value))) %>%
  filter(!is.na(value)) %>%
  merge(total_counts, by = "group") %>%
  mutate(proportion = value / total)

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

# Perform statistical tests for pos cells (RIGHT-SIDED)

classifications <- unique(proportion_data$classification)

# Exclude "Num Cells" from statistical testing

classifications <- classifications[!grepl("Num Cells", classifications, ignore.case = TRUE)]
t_test_results <- data.frame()

for (class in classifications) {
  control_data <- proportion_data[proportion_data$classification == class & proportion_data$group == "Control", ]
  t1d_data <- proportion_data[proportion_data$classification == class & proportion_data$group == "T1D", ]
  
  if (nrow(control_data) > 0 && nrow(t1d_data) > 0) {
    control_count <- control_data$value[1]
    control_total <- control_data$total[1]
    t1d_count <- t1d_data$value[1]
    t1d_total <- t1d_data$total[1]
    
    # Calculate percentages
    control_percent <- (control_count / control_total) * 100
    t1d_percent <- (t1d_count / t1d_total) * 100
    
    tryCatch({
      # Right-sided proportion test (tests if T1D proportion > Control proportion)
      prop_test <- prop.test(c(control_count, t1d_count), c(control_total, t1d_total), 
                            alternative = "less", correct = TRUE)
      
      # For t-test, we'll compare the raw percentages as single values
      # This is less appropriate than prop.test for count data, but included as requested
      control_percent_value <- control_percent
      t1d_percent_value <- t1d_percent
      
      # Since we only have single percentage values per group, we'll use a simpler approach
      # Create vectors with the percentage repeated according to some reasonable sample size
      # or use the actual counts to create binary data for t-test
      
      # Method: Create binary data (0s and 1s) based on counts
      control_binary <- c(rep(1, control_count), rep(0, control_total - control_count))
      t1d_binary <- c(rep(1, t1d_count), rep(0, t1d_total - t1d_count))
      
      # Perform unpaired, two-sided t-test on binary data
      t_test <- t.test(control_binary, t1d_binary, 
                      paired = FALSE, alternative = "two.sided", var.equal = FALSE)
      
      # Use prop.test p-value as primary (more appropriate for proportions)
      primary_p_value <- prop_test$p.value
      
      # Determine significance level
      if (primary_p_value < 0.001) {
        significance <- "***"
      } else if (primary_p_value < 0.01) {
        significance <- "**"
      } else if (primary_p_value < 0.05) {
        significance <- "*"
      } else {
        significance <- "ns"
      }
      
      t_test_results <- rbind(t_test_results, data.frame(
        classification = class,
        p_value = primary_p_value,
        prop_test_p = prop_test$p.value,
        t_test_p = t_test$p.value,
        significance = significance
      ))
    }, error = function(e) {
      t_test_results <<- rbind(t_test_results, data.frame(
        classification = class,
        p_value = NA,
        prop_test_p = NA,
        t_test_p = NA,
        significance = "ns"
      ))
    })
  } else {
    t_test_results <- rbind(t_test_results, data.frame(
      classification = class,
      p_value = NA,
      prop_test_p = NA,
      t_test_p = NA,
      significance = "ns"
    ))
  }
}

# Add significance to summary data

summary_data <- merge(summary_data, t_test_results, by = "classification", all.x = TRUE)

# Filter and reorder data for plotting

classification_order <- c("Num Negative", "Num 1+", "Num 2+", "Num 3+")

filtered_summary_data <- summary_data %>%
  filter(classification %in% classification_order) %>%
  mutate(classification = factor(classification, levels = classification_order, 
                                labels = c("Negative", "1+", "2+", "3+")))

filtered_t_test_results <- t_test_results %>%
  filter(classification %in% classification_order) %>%
  mutate(classification = factor(classification, levels = classification_order,
                                labels = c("Negative", "1+", "2+", "3+")))

# Calculate significance star positions, max values get shown in the plot

max_values <- filtered_summary_data %>%
  group_by(classification) %>%
  summarise(max_y = max(mean_value + se) * 1.1, .groups = "drop")

significance_data <- merge(filtered_t_test_results, max_values, by = "classification")


# Create plot

p <- ggplot(filtered_summary_data, aes(x = classification, y = mean_value, fill = group)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.8), width = 0.7) +
  geom_errorbar(aes(ymin = mean_value - se, ymax = mean_value + se), 
                position = position_dodge(width = 0.8), width = 0.2) +
  geom_text(data = significance_data, 
            aes(x = classification, y = max_y, label = significance),
            inherit.aes = FALSE, hjust = 0.5, vjust = 0, size = 5) +
  scale_fill_manual(values = c("Control" = "lightblue", "T1D" = "lightcoral")) +
  labs(
    title = "mIR-155 expression",
    x = NULL,
    y = "Percentage (%)",
    fill = "Group"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    plot.title = element_text(hjust = 0.5)
  )

print(p)

# Print results
cat("\nRight-sided statistical test results (T1D > Control):\n")
cat("Significant stars indicate T1D has significantly higher values than Control\n")
print(filtered_t_test_results)
