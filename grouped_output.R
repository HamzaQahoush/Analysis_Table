




result_fn_total <- write_total_to_workbook(df_total_c1_pivoted)
# end of total function############


###### start of grouping #######
# startRow will be in the first iteration 
startRow <- 4 

#A5_characteristic will be provided in the function  also df_grouped

writeData(wb, sheet = 1, x = "A5_characteristic", startCol = 1, startRow = startRow)
for (i in 1:nrow(df_grouped )) {
  writeData(wb, sheet = 1, x = df_grouped$A5_characteristic[i], startCol = 2, startRow = startRow)
  startRow <- startRow +1 
  last_row <- startRow
}


df_grouped <- df_grouped[-1]
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

# adding total count and % and get their values.
total_count_rowwise <- apply(sapply(df_grouped, function(col) sapply(col, `[[`, 1)), 1, sum)
total_percent_rowwise <- apply(sapply(df_grouped, function(col) sapply(col, `[[`, 2)), 1, sum)

writeData(wb, sheet = 1, x = total_count_rowwise, startCol = startCol, startRow = startRow, colNames = FALSE, rowNames = FALSE)
writeData(wb, sheet = 1, x = total_percent_rowwise, startCol = startCol + 1, startRow = startRow, colNames = FALSE, rowNames = FALSE)

style_percent <- createStyle(numFmt = "0.00%")
addStyle(wb, sheet = 1, style = style_percent, cols = seq(4, by = 2, length.out = ncol(df_grouped) + 1), rows = 3:10, gridExpand = TRUE)




saveWorkbook(result_fn, "output_total_fn.xlsx", overwrite = TRUE)
