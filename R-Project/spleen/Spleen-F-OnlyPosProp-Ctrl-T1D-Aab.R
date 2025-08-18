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
cat("Select the AAb CSV file...\n")
path3 <- file.choose()


# Read specific columns (column 1 or 2. Either 5,6 or 7,8 or 8,9)
c1 <- 5
c2 <- 6
control <- read.csv(path1)[, c(c1, c2)]
t1d <- read.csv(path2)[, c(c1, c2)]
aab <- read.csv(path3)[, c(c1, c2)]

# DEBUG: Clean and convert the data to numeric
control[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(control[, 2])))
t1d[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(t1d[, 2])))
aab[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(aab[, 2])))

# DEBUG: Remove rows with NA values
control <- control[complete.cases(control), ]
t1d <- t1d[complete.cases(t1d), ]
aab <- aab[complete.cases(aab), ]

# Add group labels
control$group <- "Control"
t1d$group <- "T1D"
aab$group <- "AAb"

# Combine datasets
combined_data <- rbind(control, t1d, aab)

# Handle data reshaping
if(ncol(combined_data) == 3) {
  colnames(combined_data)[1:2] <- c("classification", "value")
  combined_data_long <- combined_data
} else {
  control_wide <- control %>% select(-group)
  t1d_wide <- t1d %>% select(-group)
  aab_wide <- aab %>% select(-group)
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

  aab_long <- aab_wide %>%
    mutate(sample_id = row_number()) %>%
    pivot_longer(cols = all_of(classification_cols),
                 names_to = "classification",
                 values_to = "value") %>%
    mutate(group = "AAb") %>%
    select(-sample_id)

  combined_data_long <- rbind(control_long, t1d_long, aab_long)
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
  aab_total_raw <- read.csv(path3)[2, c2]

  control_total <- as.numeric(gsub("[^0-9.-]", "", as.character(control_total_raw)))
  t1d_total <- as.numeric(gsub("[^0-9.-]", "", as.character(t1d_total_raw)))
  aab_total <- as.numeric(gsub("[^0-9.-]", "", as.character(aab_total_raw)))

  total_counts <- data.frame(
    group = c("Control", "T1D", "AAb"),
    total = c(control_total, t1d_total, aab_total)
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

## Perform directional proportion tests vs Control for T1D and AAb
classifications <- unique(proportion_data$classification)
test_rows <- list()
for (class in classifications) {
  ctrl <- proportion_data[proportion_data$classification == class & proportion_data$group == "Control", ]
  if (nrow(ctrl) == 0) next
  ctrl_count <- ctrl$value[1]; ctrl_total <- ctrl$total[1]
  for (grp in c("T1D","AAb")) {
    grp_data <- proportion_data[proportion_data$classification == class & proportion_data$group == grp, ]
    if (nrow(grp_data) == 0) next
    grp_count <- grp_data$value[1]; grp_total <- grp_data$total[1]
    direction <- ifelse(grepl("Negative", class, ignore.case = TRUE), "less", "greater")
    res <- tryCatch({
      pt <- prop.test(c(grp_count, ctrl_count), c(grp_total, ctrl_total), alternative = direction, correct = TRUE)
      pv <- pt$p.value
      sig <- ifelse(pv < 0.001, "***", ifelse(pv < 0.01, "**", ifelse(pv < 0.05, "*", "")))
      data.frame(classification = class, group = grp, p_value = pv, significance = sig)
    }, error = function(e) {
      data.frame(classification = class, group = grp, p_value = NA, significance = "")
    })
    test_rows[[paste(class, grp)]] <- res
  }
}
t_test_results <- if (length(test_rows)) do.call(rbind, test_rows) else data.frame(classification=character(),group=character(),p_value=numeric(),significance=character())

# Merge per-group significance
summary_data <- summary_data %>% left_join(t_test_results, by = c("classification","group"))

# Filter and reorder data for plotting (excluding Negative)
classification_order <- c("Num 1+", "Num 2+", "Num 3+")

filtered_summary_data <- summary_data %>%
  filter(classification %in% classification_order) %>%
  mutate(classification = factor(classification, levels = classification_order, labels = c("1+","2+","3+")),
         group = factor(group, levels = c("Control", "T1D", "AAb")))

# Calculate per bar y for significance (above its error bar)
significance_data <- filtered_summary_data %>%
  filter(group != "Control") %>%
  group_by(classification, group) %>%
  summarise(y = max(mean_value + se) * 1.08, significance = first(significance), .groups = "drop")

# Create plot
p <- ggplot(filtered_summary_data, aes(x = classification, y = mean_value, fill = group)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.8), width = 0.7) +
  geom_errorbar(aes(ymin = mean_value - se, ymax = mean_value + se), position = position_dodge(width = 0.8), width = 0.2) +
  geom_text(data = significance_data,
            aes(x = classification, y = y, label = significance, group = group),
            position = position_dodge(width = 0.8), vjust = -0.1, size = 5) +
  scale_fill_manual(values = c(Control = "lightblue", T1D = "lightcoral", AAb = "khaki3")) +
  labs(,
    x = NULL,
    y = "miR-155 Expression (%)",
    fill = "Condition"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1, size = 8),
    axis.text.y = element_text(size = 8),
    axis.title.x = element_text(size = 10),
    axis.title.y = element_text(size = 10),
    plot.title = element_text(hjust = 0.5, size = 12),
    legend.title = element_text(size = 10),
    legend.text = element_text(size = 8)
  )

print(p)

# Print results
cat("\nDirectional proportion test results (vs Control):\n")
cat("Negative classes: star = LOWER than Control (disease < Control).\n")
cat("Positive 1+/2+/3+: star = HIGHER than Control (disease > Control).\n\n")
print(t_test_results)
