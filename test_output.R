

df <- sample_data_grouped

wb <- createWorkbook()
addWorksheet(wb, "Data")

# Write "total" in first row, first column
writeData(wb, sheet = 1, x = "total", startCol = 1, startRow = 3)

center_style <- createStyle(halign = "center")

start_col <- 3

# Loop to write headers and data for each C1_main_water_source
for (i in 1:nrow(df)) {
  writeData(wb, sheet = 1, x = df$C1_main_water_source[i], startCol = start_col, startRow = 1)
  addStyle(wb, sheet = 1, style = center_style, cols = start_col:(start_col+1), rows = 1)
  
  writeData(wb, sheet = 1, x = "count", startCol = start_col, startRow = 2, colNames = FALSE, rowNames = FALSE)
  writeData(wb, sheet = 1, x = "%", startCol = start_col + 1, startRow = 2, colNames = FALSE, rowNames = FALSE)
  writeData(wb, sheet = 1, x = df$count[i], startCol = start_col, startRow = 3, colNames = FALSE, rowNames = FALSE)
  
  # Write percentage as decimal value
  writeData(wb, sheet = 1, x = df$percent[i] , startCol = start_col + 1, startRow = 3, colNames = FALSE, rowNames = FALSE)
  
  mergeCells(wb, sheet = 1, cols = start_col:(start_col+1), rows = 1:1)
  start_col <- start_col + 2
}

# Add "total" column after loop
writeData(wb, sheet = 1, x = "total", startCol = start_col, startRow = 1)
writeData(wb, sheet = 1, x = "count", startCol = start_col, startRow = 2)
writeData(wb, sheet = 1, x = "%", startCol = start_col + 1, startRow = 2)
writeData(wb, sheet = 1, x = sum(df$count), startCol = start_col, startRow = 3)

# Write total percentage as decimal value
writeData(wb, sheet = 1, x = sum(df$percent), startCol = start_col + 1, startRow = 3)

mergeCells(wb, sheet = 1, cols = start_col:(start_col+1), rows = 1:1)

# Apply percentage formatting to the cells with percentages
style_percent <- createStyle(numFmt = "0.00%")
addStyle(wb, sheet = 1, style = style_percent, cols = seq(4, by = 2, length.out = nrow(df) + 1), 
         rows = 3, gridExpand = TRUE)

saveWorkbook(wb, "output.xlsx", overwrite = TRUE)
