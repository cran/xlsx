#
#
#

test.workbook <- function()
{
  require(xlsx)
  #print(.jclassPath())

  cat("Create an empty workbook ... ") 
  wb <- createWorkbook()

  cat("Create a sheet called 'Sheet1' ... \n")
  sheet1 <- createSheet(wb, sheetName="Sheet1")

  cat("Create another sheet called 'Sheet2' ... \n")
  sheet2 <- createSheet(wb, sheetName="Sheet2")
  
  cat("Get sheets ...\n")
  print(getSheets(wb))

  cat("Remove sheet named 'Sheet2' ...\n")
  removeSheet(wb, sheetName="Sheet2")
  
  cat("Get sheets ...\n")
  print(getSheets(wb))
 
}








