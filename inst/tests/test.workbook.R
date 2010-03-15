#
#
#

test.workbook <- function()
{
  require(xlsx)
  cat("##################################################\n")
  cat("Testing basic workbook functions\n")
  cat("##################################################\n")

  cat("Create an empty workbook ... ") 
  wb <- createWorkbook()
  cat("OK\n")

  cat("Create a sheet called 'Sheet1' ... ")
  sheet1 <- createSheet(wb, sheetName="Sheet1")
  cat("OK\n")

  cat("Create another sheet called 'Sheet2' ... ")
  sheet2 <- createSheet(wb, sheetName="Sheet2")
  cat("OK\n")
  
  cat("Get sheets ... ")
  sheets <- getSheets(wb)
  stopifnot(length(sheets) == 2)
  cat("OK\n")

  cat("Remove sheet named 'Sheet2' ... ")
  removeSheet(wb, sheetName="Sheet2")
  sheets <- getSheets(wb)  
  stopifnot(length(sheets) == 1)
  cat("OK\n")
   
}








