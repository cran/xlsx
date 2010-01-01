# Write a data.frame to a new xlsx file. 
# A convenience function.
#

write.xlsx <- function(x, file, sheetName="Sheet 1", formatTemplate=NULL,
  col.names=TRUE, row.names=TRUE)
{
  iOffset <- jOffset <- 0
  if (col.names)
    iOffset <- 1
  if (row.names)
    jOffset <- 1
    
  noRows <- nrow(x) + iOffset
  noCols <- ncol(x) + jOffset
  
  wb <- createWorkbook()
  sheet <- createSheet(wb, sheetName)

  rows  <- createRow(sheet, 1:noRows)      # create rows 
  cells <- createCell(rows, col=1:noCols)  # create cells

  for (ic in 1:ncol(x))
    mapply(setCellValue, cells[(1+iOffset):noRows,ic+jOffset], x[,ic])

  if (col.names)
    mapply(setCellValue, cells[1,(1+jOffset):noCols], colnames(x))

  if (row.names)
    mapply(setCellValue, cells[(1+iOffset):noRows, 1], rownames(x))

  # Date and POSIXct classes need to be formatted
  indDT <- which(sapply(x, class) == "Date")
  dateFormat <- createCellStyle(wb, dataFormat="m/d/yyyy")
  for (ic in indDT){
    lapply(cells[(1+iOffset):noRows,ic+jOffset], setCellStyle, dateFormat)
  }
  indDT <- which(sapply(x, class) == "POSIXct")
  datetimeFormat <- createCellStyle(wb, dataFormat="m/d/yyyy h:mm:ss;@")
  for (ic in indDT){
    lapply(cells[(1+iOffset):noRows,ic+jOffset], setCellStyle, datetimeFormat)
  }
  
  # add a general format template
  if (!is.null(formatTemplate)){
    cat("formatTemplate is not yet supported.  Patches welcome.")
    wb <- wb  # call the custom function
  }
    
  saveWorkbook(wb, file)

  invisible()
}
