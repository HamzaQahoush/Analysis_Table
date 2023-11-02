if (!require("pacman")) install.packages("pacman")
pacman::p_load(tidyverse,
               rlang,
               dplyr,
               readxl,
               openxlsx,
               lubridate,
               data.validator,
               randomcoloR,
               mgsub,
               data.table,
               rstatix,
               stringi
)




reshape_data_total <- function(df, val){
  df <- df %>% 
    filter(!is.na(!!sym(val))) %>%
    group_by(!!sym(val)) %>%
    summarise(count = sum(weight)) %>%
    mutate(total_of_totals = sum(count),
           'percent' = (count / total_of_totals)) %>%
    select(-total_of_totals) %>%
    ungroup()
  
  return(df)
}


reshape_data_grouped_value <- function(df, group_by_val, value) {
    df_grouped <- df %>%
      filter(!is.na(!!sym(value))) %>% 
      group_by(!!sym(group_by_val), !!sym(value)) %>% 
      summarise(count = sum(weight, na.rm = TRUE)) %>% 
      mutate(percent = count / sum(count)) %>% 
      ungroup()
    
    
    df_grouped[] <- lapply(df_grouped, function(col) {
      sapply(col, function(x) if (is.null(x)) list(0, 0) else x, simplify = FALSE)
    })
    
    return(df_grouped)
  }
  
reshape_data_grouped_values <- function(df, group_by_vals, value) {
  group_by_cols <- rlang::syms(group_by_vals)  # Convert the column names to symbols
  
 # Convert the column names to symbols
  value_sym <- enquo(value)            # Capture the value column as a quosure
  
  df_grouped <- df %>%
    filter(!is.na(!!value_sym)) %>% 
    group_by(!!!group_by_cols, !!value_sym) %>% 
    summarise(count = sum(weight, na.rm = TRUE)) %>% 
    mutate(percent = count / sum(count)) %>% 
    ungroup()
  
  df_grouped[] <- lapply(df_grouped, function(col) {
    sapply(col, function(x) if (is.null(x)) list(0, 0) else x, simplify = FALSE)
  })
  
  return(df_grouped)
}




combine_and_pivot <- function(df, column_name) {
  # Combine count and percent into a list column
  df$count_percent <- mapply(list, df$count, df$percent, SIMPLIFY = FALSE)
  
  # Drop the count and percent columns
  df <- df %>% select(-count, -percent)
  
  # Pivot the data wider and order columns
  df_pivoted <- df %>%
    pivot_wider(names_from = column_name, values_from = count_percent) %>% 
    select(order(names(.)))
  
  
  df_pivoted[] <- lapply(df_pivoted, function(col) {
    sapply(col, function(x) if (is.null(x)) list(0, 0) else x, simplify = FALSE)
  })
  
  return(df_pivoted)
}


write_total_to_workbook <- function(df) {
  wb <- createWorkbook()
  addWorksheet(wb, "Data")
  
  # Write "total" in first row, first column
  writeData(wb, sheet = 1, x = "total", startCol = 1, startRow = 3)
  
  center_style <- createStyle(halign = "center")
  start_col_total <- 3
  
  for (i in 1:ncol(df)) {
    col_data <- df[[names(df)[i]]]
    count_data <- sapply(col_data, `[[`, 1)
    percent_data <- sapply(col_data, `[[`, 2)
    
    writeData(wb, sheet = 1, x = names(df)[i], startCol = start_col_total, startRow = 1)
    addStyle(wb, sheet = 1, style = center_style, cols = start_col_total:(start_col_total+1), rows = 1)
    
    writeData(wb, sheet = 1, x = "count", startCol = start_col_total, startRow = 2, colNames = FALSE, rowNames = FALSE)
    writeData(wb, sheet = 1, x = "%", startCol = start_col_total + 1, startRow = 2, colNames = FALSE, rowNames = FALSE)
    
    # Write data
    writeData(wb, sheet = 1, x = count_data, startCol = start_col_total, startRow = 3, colNames = FALSE, rowNames = FALSE)
    writeData(wb, sheet = 1, x = percent_data, startCol = start_col_total + 1, startRow = 3, colNames = FALSE, rowNames = FALSE)
    
    mergeCells(wb, sheet = 1, cols = start_col_total:(start_col_total+1), rows = 1:1)
    start_col_total <- start_col_total + 2
  }
  
  ## adding total count and % and get their values.
  total_count <- sum(sapply(df, function(col) sum(sapply(col, `[[`, 1))))
  total_percent <- sum(sapply(df, function(col) sum(sapply(col, `[[`, 2))))
  
  writeData(wb, sheet = 1, x = "total", startCol = start_col_total, startRow = 1)
  writeData(wb, sheet = 1, x = "count", startCol = start_col_total, startRow = 2)
  writeData(wb, sheet = 1, x = "%", startCol = start_col_total + 1, startRow = 2)
  writeData(wb, sheet = 1, x = total_count, startCol = start_col_total, startRow = 3)
  writeData(wb, sheet = 1, x = total_percent, startCol = start_col_total + 1, startRow = 3)
  
  mergeCells(wb, sheet = 1, cols = start_col_total:(start_col_total+1), rows = 1:1)
  
  # Apply percentage formatting to the cells with percentages
  style_percent <- createStyle(numFmt = "0.00%")
  addStyle(wb, sheet = 1, style = style_percent, cols = seq(4, by = 2, length.out = ncol(df) + 1), rows = 3, gridExpand = TRUE)
  
  return(wb)
}

write_grouped_data_to_workbook <- function(wb, df_grouped, grouped_val, startRow = 4) {
  characteristic_col_str <- as.character(grouped_val)
  browser()
  # Write grouped_val
  # here 
  
  writeData(wb, sheet = 1, x = characteristic_col_str, startCol = 1, startRow = startRow)
  for (i in 1:nrow(df_grouped)) {
    writeData(wb, sheet = 1, x = df_grouped[[grouped_val]][i], startCol = 2, startRow = startRow)
    startRow <- startRow + 1
  }
  grouped_val_char <- as.character(grouped_val)
  
  df_grouped <- df_grouped[, setdiff(names(df_grouped), grouped_val_char), drop = FALSE]
  startCol <- 3
  startRow <- 4
  
  for (i in 1:ncol(df_grouped)) {
    
    col_data <- df_grouped[[i]]
    count_data <- sapply(col_data, function(x) x[[1]])
    percent_data <- sapply(col_data, function(x) x[[2]])
    
    # Write data
    writeData(wb, sheet = 1, x = count_data, startCol = startCol, startRow = startRow, colNames = FALSE, rowNames = FALSE)
    writeData(wb, sheet = 1, x = percent_data, startCol = startCol + 1, startRow = startRow, colNames = FALSE, rowNames = FALSE)
    
    startCol <- startCol + 2
  }
  
  # Adding total count and % and get their values.
  total_count_rowwise <- apply(sapply(df_grouped, function(col) sapply(col, `[[`, 1)), 1, sum)
  total_percent_rowwise <- apply(sapply(df_grouped, function(col) sapply(col, `[[`, 2)), 1, sum)
  
  writeData(wb, sheet = 1, x = total_count_rowwise, startCol = startCol, startRow = startRow, colNames = FALSE, rowNames = FALSE)
  writeData(wb, sheet = 1, x = total_percent_rowwise, startCol = startCol + 1, startRow = startRow, colNames = FALSE, rowNames = FALSE)
  
  style_percent <- createStyle(numFmt = "0.00%")
  addStyle(wb, sheet = 1, style = style_percent, cols = seq(4, by = 2, length.out = ncol(df_grouped) + 1), rows = 3:10, gridExpand = TRUE)
  return(wb)
}


write_grouped_values <- function (df, group_by_cols, val) {
  reshaped_total_df <- reshape_data_total(df, val)
  pivoted_total_df <- combine_and_pivot(reshaped_total_df, val)
  wb_total <- write_total_to_workbook(pivoted_total_df)
  
  group_by_cols <- rlang::syms(group_by_cols)  
  
   startRow <- 4
  
   for (group in group_by_cols) {
     browser()
     output <- reshape_data_grouped_value(df, group, val)
     output_pivoted <- combine_and_pivot(output, val)
     print(paste0("nrow(output_pivoted)><<<",nrow(output_pivoted)))
    write_in_wb <- write_grouped_data_to_workbook(wb_total, output_pivoted, group, startRow = startRow) # here is the problem
     startRow <- startRow + nrow(output_pivoted)
   }
   
  return(wb_total)
}

