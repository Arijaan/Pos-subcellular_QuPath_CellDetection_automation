library(ggplot2)
library(dplyr)
library(tidyr)
library(tibble)

# Comprehensive Statistical Testing Script for Control vs T1D
# This script runs multiple statistical tests to compare proportions and distributions

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

# ============================================================================
# COMPREHENSIVE STATISTICAL TESTING
# ============================================================================

cat("\n", strrep("=", 80), "\n")
cat("COMPREHENSIVE STATISTICAL ANALYSIS: CONTROL vs T1D\n")
cat(strrep("=", 80), "\n\n")

classifications <- unique(proportion_data$classification)
all_test_results <- data.frame()

for (class in classifications) {
  cat("\n", strrep("-", 60), "\n")
  cat("ANALYZING CLASSIFICATION:", class, "\n")
  cat(strrep("-", 60), "\n")
  
  control_data <- proportion_data[proportion_data$classification == class & proportion_data$group == "Control", ]
  t1d_data <- proportion_data[proportion_data$classification == class & proportion_data$group == "T1D", ]
  
  if (nrow(control_data) > 0 && nrow(t1d_data) > 0) {
    control_count <- control_data$value[1]
    control_total <- control_data$total[1]
    t1d_count <- t1d_data$value[1]
    t1d_total <- t1d_data$total[1]
    
    # Display basic statistics
    cat("Control: ", control_count, "/", control_total, " (", round(control_count/control_total*100, 2), "%)\n")
    cat("T1D: ", t1d_count, "/", t1d_total, " (", round(t1d_count/t1d_total*100, 2), "%)\n\n")
    
    # Initialize results for this classification
    test_results <- data.frame(
      classification = class,
      control_count = control_count,
      control_total = control_total,
      control_prop = control_count/control_total,
      t1d_count = t1d_count,
      t1d_total = t1d_total,
      t1d_prop = t1d_count/t1d_total
    )
    
    # ========================================================================
    # 1. PROPORTION TESTS (Two-sided, Greater, Less)
    # ========================================================================
    cat("1. PROPORTION TESTS:\n")
    
    # Two-sided proportion test
    tryCatch({
      prop_test_two <- prop.test(c(control_count, t1d_count), c(control_total, t1d_total), 
                                alternative = "two.sided", correct = TRUE)
      cat("   Two-sided: p =", format(prop_test_two$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$prop_two_sided_p <- prop_test_two$p.value
    }, error = function(e) {
      cat("   Two-sided: ERROR -", e$message, "\n")
      test_results$prop_two_sided_p <<- NA
    })
    
    # Greater: Control > T1D
    tryCatch({
      prop_test_greater <- prop.test(c(control_count, t1d_count), c(control_total, t1d_total), 
                                    alternative = "greater", correct = TRUE)
      cat("   Control > T1D: p =", format(prop_test_greater$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$prop_control_greater_p <- prop_test_greater$p.value
    }, error = function(e) {
      cat("   Control > T1D: ERROR -", e$message, "\n")
      test_results$prop_control_greater_p <<- NA
    })
    
    # Less: Control < T1D
    tryCatch({
      prop_test_less <- prop.test(c(control_count, t1d_count), c(control_total, t1d_total), 
                                 alternative = "less", correct = TRUE)
      cat("   Control < T1D: p =", format(prop_test_less$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$prop_control_less_p <- prop_test_less$p.value
    }, error = function(e) {
      cat("   Control < T1D: ERROR -", e$message, "\n")
      test_results$prop_control_less_p <<- NA
    })
    
    # ========================================================================
    # 2. CHI-SQUARE TESTS
    # ========================================================================
    cat("\n2. CHI-SQUARE TESTS:\n")
    
    # Create contingency table
    contingency_table <- matrix(c(control_count, control_total - control_count,
                                 t1d_count, t1d_total - t1d_count), 
                               nrow = 2, byrow = TRUE,
                               dimnames = list(c("Control", "T1D"), c("Positive", "Negative")))
    
    # Pearson's Chi-square test
    tryCatch({
      chi_test <- chisq.test(contingency_table, correct = FALSE)
      cat("   Pearson Chi-square: p =", format(chi_test$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$chi_square_p <- chi_test$p.value
      test_results$chi_square_statistic <- chi_test$statistic
    }, error = function(e) {
      cat("   Pearson Chi-square: ERROR -", e$message, "\n")
      test_results$chi_square_p <<- NA
      test_results$chi_square_statistic <<- NA
    })
    
    # Chi-square with Yates' continuity correction
    tryCatch({
      chi_test_yates <- chisq.test(contingency_table, correct = TRUE)
      cat("   Chi-square (Yates): p =", format(chi_test_yates$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$chi_square_yates_p <- chi_test_yates$p.value
    }, error = function(e) {
      cat("   Chi-square (Yates): ERROR -", e$message, "\n")
      test_results$chi_square_yates_p <<- NA
    })
    
    # ========================================================================
    # 3. FISHER'S EXACT TEST
    # ========================================================================
    cat("\n3. FISHER'S EXACT TEST:\n")
    
    # Two-sided Fisher's exact test
    tryCatch({
      fisher_test_two <- fisher.test(contingency_table, alternative = "two.sided")
      cat("   Two-sided: p =", format(fisher_test_two$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$fisher_two_sided_p <- fisher_test_two$p.value
      test_results$fisher_odds_ratio <- fisher_test_two$estimate
    }, error = function(e) {
      cat("   Two-sided: ERROR -", e$message, "\n")
      test_results$fisher_two_sided_p <<- NA
      test_results$fisher_odds_ratio <<- NA
    })
    
    # Greater: odds ratio > 1
    tryCatch({
      fisher_test_greater <- fisher.test(contingency_table, alternative = "greater")
      cat("   Control > T1D: p =", format(fisher_test_greater$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$fisher_control_greater_p <- fisher_test_greater$p.value
    }, error = function(e) {
      cat("   Control > T1D: ERROR -", e$message, "\n")
      test_results$fisher_control_greater_p <<- NA
    })
    
    # Less: odds ratio < 1
    tryCatch({
      fisher_test_less <- fisher.test(contingency_table, alternative = "less")
      cat("   Control < T1D: p =", format(fisher_test_less$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$fisher_control_less_p <- fisher_test_less$p.value
    }, error = function(e) {
      cat("   Control < T1D: ERROR -", e$message, "\n")
      test_results$fisher_control_less_p <<- NA
    })
    
    # ========================================================================
    # 4. T-TESTS ON BINARY DATA
    # ========================================================================
    cat("\n4. T-TESTS (on binary data):\n")
    
    # Create binary vectors
    control_binary <- c(rep(1, control_count), rep(0, control_total - control_count))
    t1d_binary <- c(rep(1, t1d_count), rep(0, t1d_total - t1d_count))
    
    # Two-sample t-test (Welch)
    tryCatch({
      t_test_welch <- t.test(control_binary, t1d_binary, alternative = "two.sided", var.equal = FALSE)
      cat("   Welch t-test: p =", format(t_test_welch$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$t_test_welch_p <- t_test_welch$p.value
      test_results$t_test_welch_statistic <- t_test_welch$statistic
    }, error = function(e) {
      cat("   Welch t-test: ERROR -", e$message, "\n")
      test_results$t_test_welch_p <<- NA
      test_results$t_test_welch_statistic <<- NA
    })
    
    # Student's t-test (equal variances)
    tryCatch({
      t_test_student <- t.test(control_binary, t1d_binary, alternative = "two.sided", var.equal = TRUE)
      cat("   Student t-test: p =", format(t_test_student$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$t_test_student_p <- t_test_student$p.value
    }, error = function(e) {
      cat("   Student t-test: ERROR -", e$message, "\n")
      test_results$t_test_student_p <<- NA
    })
    
    # ========================================================================
    # 5. NON-PARAMETRIC TESTS
    # ========================================================================
    cat("\n5. NON-PARAMETRIC TESTS:\n")
    
    # Mann-Whitney U test (Wilcoxon rank-sum test)
    tryCatch({
      wilcox_test <- wilcox.test(control_binary, t1d_binary, alternative = "two.sided")
      cat("   Mann-Whitney U: p =", format(wilcox_test$p.value, scientific = TRUE, digits = 4), "\n")
      test_results$mann_whitney_p <- wilcox_test$p.value
      test_results$mann_whitney_statistic <- wilcox_test$statistic
    }, error = function(e) {
      cat("   Mann-Whitney U: ERROR -", e$message, "\n")
      test_results$mann_whitney_p <<- NA
      test_results$mann_whitney_statistic <<- NA
    })
    
    # ========================================================================
    # 6. EFFECT SIZE MEASURES
    # ========================================================================
    cat("\n6. EFFECT SIZE MEASURES:\n")
    
    # Cohen's d
    tryCatch({
      pooled_sd <- sqrt(((length(control_binary)-1)*var(control_binary) + 
                        (length(t1d_binary)-1)*var(t1d_binary)) / 
                       (length(control_binary) + length(t1d_binary) - 2))
      cohens_d <- (mean(control_binary) - mean(t1d_binary)) / pooled_sd
      cat("   Cohen's d =", round(cohens_d, 4), "\n")
      test_results$cohens_d <- cohens_d
    }, error = function(e) {
      cat("   Cohen's d: ERROR -", e$message, "\n")
      test_results$cohens_d <<- NA
    })
    
    # Phi coefficient (effect size for chi-square)
    tryCatch({
      n_total <- control_total + t1d_total
      if (!is.na(test_results$chi_square_statistic)) {
        phi <- sqrt(test_results$chi_square_statistic / n_total)
        cat("   Phi coefficient =", round(phi, 4), "\n")
        test_results$phi_coefficient <- phi
      }
    }, error = function(e) {
      cat("   Phi coefficient: ERROR -", e$message, "\n")
      test_results$phi_coefficient <<- NA
    })
    
    # Risk difference and relative risk
    control_prop <- control_count / control_total
    t1d_prop <- t1d_count / t1d_total
    risk_difference <- t1d_prop - control_prop
    relative_risk <- ifelse(control_prop > 0, t1d_prop / control_prop, NA)
    
    cat("   Risk Difference =", round(risk_difference, 4), "\n")
    cat("   Relative Risk =", round(relative_risk, 4), "\n")
    
    test_results$risk_difference <- risk_difference
    test_results$relative_risk <- relative_risk
    
    # Add to overall results
    all_test_results <- rbind(all_test_results, test_results)
    
  } else {
    cat("Insufficient data for testing\n")
    
    # Add empty row for this classification
    empty_results <- data.frame(
      classification = class,
      control_count = NA, control_total = NA, control_prop = NA,
      t1d_count = NA, t1d_total = NA, t1d_prop = NA,
      prop_two_sided_p = NA, prop_control_greater_p = NA, prop_control_less_p = NA,
      chi_square_p = NA, chi_square_statistic = NA, chi_square_yates_p = NA,
      fisher_two_sided_p = NA, fisher_odds_ratio = NA, 
      fisher_control_greater_p = NA, fisher_control_less_p = NA,
      t_test_welch_p = NA, t_test_welch_statistic = NA, t_test_student_p = NA,
      mann_whitney_p = NA, mann_whitney_statistic = NA,
      cohens_d = NA, phi_coefficient = NA, 
      risk_difference = NA, relative_risk = NA
    )
    all_test_results <- rbind(all_test_results, empty_results)
  }
}

# ============================================================================
# SUMMARY AND EXPORT RESULTS
# ============================================================================

cat("\n", strrep("=", 80), "\n")
cat("SUMMARY OF ALL STATISTICAL TESTS\n")
cat(strrep("=", 80), "\n\n")

# Print summary table
print(all_test_results)

# Save results to CSV
write.csv(all_test_results, "comprehensive_statistical_tests_results.csv", row.names = FALSE)
cat("\nResults saved to: comprehensive_statistical_tests_results.csv\n")

# ============================================================================
# SIGNIFICANCE SUMMARY
# ============================================================================

cat("\n", strrep("=", 60), "\n")
cat("SIGNIFICANCE SUMMARY (p < 0.05)\n")
cat(strrep("=", 60), "\n")

significance_summary <- all_test_results %>%
  select(classification, 
         prop_two_sided_p, prop_control_greater_p, prop_control_less_p,
         chi_square_p, chi_square_yates_p,
         fisher_two_sided_p, fisher_control_greater_p, fisher_control_less_p,
         t_test_welch_p, t_test_student_p, mann_whitney_p) %>%
  mutate(
    significant_tests = apply(select(., -classification), 1, function(x) sum(x < 0.05, na.rm = TRUE)),
    total_tests = apply(select(., -classification), 1, function(x) sum(!is.na(x)))
  )

for(i in 1:nrow(significance_summary)) {
  cat(significance_summary$classification[i], ":\n")
  cat("  Significant tests:", significance_summary$significant_tests[i], 
      "out of", significance_summary$total_tests[i], "\n\n")
}

cat("Analysis complete!\n")
