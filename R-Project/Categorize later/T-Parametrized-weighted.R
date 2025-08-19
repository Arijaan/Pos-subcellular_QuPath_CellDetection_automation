suppressPackageStartupMessages({
  library(dplyr); library(tidyr); library(ggplot2); library(readr)
})

run_miR155_analysis <- function(
  paths,                      # named vector: c(Control="...", T1D="...", T2D="...") (T2D optional)
  cols = c(5,6),              # columns to read
  mode = c("unweighted","weighted"),
  classification_order = c("Num 1+","Num 2+","Num 3+"),
  weights = c("Num 1+"=1,"Num 2+"=2,"Num 3+"=3),
  baseline = "Control",
  colors = c(Control="lightblue", T1D="lightcoral", T2D="khaki3"),
  alpha_dir = 0.05,
  save_prefix = NULL,         # if not NULL: save CSV + PNG
  directional_logic = function(class_name) {
    # Return "greater" or "less" for alternative hypothesis for disease vs baseline
    if (grepl("Negative", class_name, ignore.case=TRUE)) "less" else "greater"
  }
) {
  mode <- match.arg(mode)

  # 1. Load & clean
  df_list <- lapply(names(paths), function(g) {
    dat <- suppressWarnings(read_csv(paths[[g]], show_col_types = FALSE)[, cols])
    dat[[2]] <- as.numeric(gsub("[^0-9.-]", "", as.character(dat[[2]])))
    dat <- dat[complete.cases(dat), ]
    dat$group <- g
    dat
  })
  combined <- bind_rows(df_list)

  # 2. Reshape to long (supports already-long format)
  if (ncol(combined) == 3) {
    colnames(combined)[1:2] <- c("classification","value")
    long <- combined
  } else {
    class_cols <- colnames(combined)[!colnames(combined) %in% "group"]
    long <- combined %>%
      mutate(row_id = row_number()) %>%
      pivot_longer(all_of(class_cols), names_to="classification", values_to="value") %>%
      select(-row_id)
  }

  # 3. Totals: if no "all" row, infer totals from row 2 column 2 of each original file (as in prior scripts)
  has_all <- any(long$classification == "all")
  if (has_all) {
    totals <- long %>% filter(classification=="all") %>%
      transmute(group, total = as.numeric(value))
    prop_data <- long %>% filter(classification!="all") %>% mutate(value=as.numeric(value)) %>%
      left_join(totals, by="group") %>%
      mutate(proportion = value/total)
  } else {
    totals <- tibble(
      group = names(paths),
      total = sapply(seq_along(paths), function(i) {
        raw <- suppressWarnings(read_csv(paths[[i]], show_col_types = FALSE)[2, cols[2], drop=TRUE])
        as.numeric(gsub("[^0-9.-]", "", as.character(raw)))
      })
    )
    prop_data <- long %>%
      mutate(value = as.numeric(value)) %>%
      left_join(totals, by="group") %>%
      filter(!is.na(value) & !is.na(total) & total > 0) %>%
      mutate(proportion = value/total)
  }

  # 4. Branch: weighted vs unweighted summaries
  if (mode == "unweighted") {
    summary_df <- prop_data %>%
      filter(classification %in% classification_order) %>%
      group_by(classification, group) %>%
      summarise(count = sum(value),
                total = first(total),
                proportion = count/total,
                mean_pct = proportion*100,
                se = sqrt(proportion*(1-proportion)/total)*100,
                .groups="drop")
  } else {
    # Weighted: collapse into single pooled classification
    wc <- prop_data %>%
      filter(classification %in% names(weights)) %>%
      group_by(group) %>%
      summarise(
        total = first(total),
        weighted_sum = sum(value * weights[classification], na.rm=TRUE),
        weighted_prop = weighted_sum / total,
        mean_pct = weighted_prop * 100,
        se = sqrt(weighted_prop*(1-weighted_prop)/total)*100,
        .groups='drop'
      ) %>%
      mutate(classification = "Weighted Expression")
    summary_df <- wc
  }

  # 5. Statistical tests (directional vs baseline)
  baseline_rows <- summary_df %>% filter(group == baseline)
  other_groups <- setdiff(summary_df$group, baseline)
  test_results <- list()

  if (mode == "unweighted") {
    for (cls in unique(summary_df$classification)) {
      base_row <- prop_data %>% filter(group==baseline, classification==cls) %>% slice(1)
      if (nrow(base_row)==0) next
      for (g in other_groups) {
        g_row <- prop_data %>% filter(group==g, classification==cls) %>% slice(1)
        if (nrow(g_row)==0) next
        alt <- directional_logic(cls)
        res <- tryCatch({
          pt <- prop.test(
            c(g_row$value, base_row$value),
            c(g_row$total, base_row$total),
            alternative = alt, correct = TRUE
          )
          pv <- pt$p.value
          sig <- ifelse(pv<0.001,'***', ifelse(pv<0.01,'**', ifelse(pv<0.05,'*','')))
          data.frame(mode, classification=cls, group=g, baseline, p_value=pv, significance=sig)
        }, error=function(e) data.frame(mode, classification=cls, group=g, baseline, p_value=NA, significance='err'))
        test_results[[paste(cls,g)]] <- res
      }
    }
  } else {
    # Weighted single classification
    base_row <- prop_data %>%
      filter(group==baseline, classification %in% names(weights)) %>%
      summarise(val = sum(value * weights[classification]),
                total = first(total))
    for (g in other_groups) {
      g_row <- prop_data %>%
        filter(group==g, classification %in% names(weights)) %>%
        summarise(val = sum(value * weights[classification]),
                  total = first(total))
      if (nrow(base_row)==0 || nrow(g_row)==0) next
      alt <- "greater"  # Weighted always testing higher vs baseline
      res <- tryCatch({
        pt <- prop.test(c(g_row$val, base_row$val), c(g_row$total, base_row$total), alternative=alt, correct=TRUE)
        pv <- pt$p.value
        sig <- ifelse(pv<0.001,'***', ifelse(pv<0.01,'**', ifelse(pv<0.05,'*','')))
        data.frame(mode, classification='Weighted Expression', group=g, baseline, p_value=pv, significance=sig)
      }, error=function(e) data.frame(mode, classification='Weighted Expression', group=g, baseline, p_value=NA, significance='err'))
      test_results[[g]] <- res
    }
  }
  tests_df <- if (length(test_results)) bind_rows(test_results) else tibble()

  # Attach significance to summary (per group; blank for baseline)
  summary_df <- summary_df %>% left_join(tests_df %>% select(classification, group, significance), by=c("classification","group"))
  summary_df$significance[is.na(summary_df$significance)] <- ''

  # 6. Plot
  if (mode == "unweighted") {
    plot_df <- summary_df %>% mutate(classification = factor(classification, levels=classification_order,
                                                             labels=gsub("Num ", "", classification_order)))
    p <- ggplot(plot_df, aes(x=classification, y=mean_pct, fill=group)) + 
      geom_bar(stat='identity', position=position_dodge(0.8), width=0.7) +
      geom_errorbar(aes(ymin=mean_pct-se, ymax=mean_pct+se), position=position_dodge(0.8), width=0.2) +
      geom_text(aes(label=significance, y=mean_pct+se*1.15),
                position=position_dodge(0.8), vjust=0, size=5) +
      scale_fill_manual(values=colors) +
      labs(x=NULL, y='miR-155 positive cells (%)', title='miR-155 Expression (Per Intensity Class)') +
      theme_minimal()
  } else {
    p <- ggplot(summary_df, aes(x=factor(group, levels=c(baseline, other_groups)), y=mean_pct, fill=group)) +
      geom_bar(stat='identity', width=0.6) +
      geom_errorbar(aes(ymin=mean_pct-se, ymax=mean_pct+se), width=0.18) +
      geom_text(aes(label=significance, y=mean_pct+se*1.2), vjust=0, size=6) +
      scale_fill_manual(values=colors) +
      scale_x_discrete(labels = c(Control='ND', T1D='T1D', T2D='T2D')) +
      labs(x=NULL, y='Weighted miR-155 expression (%)', title='Weighted miR-155 Expression') +
      theme_minimal()
  }

  # 7. Optional save
  if (!is.null(save_prefix)) {
    write.csv(summary_df, paste0(save_prefix, '_summary.csv'), row.names=FALSE)
    if (nrow(tests_df)) write.csv(tests_df, paste0(save_prefix, '_tests.csv'), row.names=FALSE)
    ggsave(paste0(save_prefix, '_plot.png'), p, width=6, height=4, dpi=300)
  }

  list(summary=summary_df, tests=tests_df, plot=p, mode=mode)
}

# Example usage (uncomment and provide real paths):
# res_unweighted <- run_miR155_analysis(
#   paths = c(Control='ctrl.csv', T1D='t1d.csv', T2D='t2d.csv'),
#   mode = 'unweighted', save_prefix='unweighted_run'
# )
# res_weighted <- run_miR155_analysis(
#   paths = c(Control='ctrl.csv', T1D='t1d.csv', T2D='t2d.csv'),
#   mode = 'weighted', save_prefix='weighted_run'
# )
# print(res_unweighted$plot); print(res_weighted$plot)