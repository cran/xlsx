#
#
#

test.ranges <- function(type="xlsx")
{
  cat("##################################################\n")
  cat("Testing ranges\n")
  cat("##################################################\n")
  filename <- paste("test_import.", type, sep="")

  cat("Get named ranges from test_import.xlsx ... ")
  file <- system.file("tests", filename, package = "xlsx")
  wb <- loadWorkbook(file)
  sheets <- getSheets(wb)
  sheet <- sheets[["deletedFields"]]
  
  ranges <- getRanges(wb)
  range <- ranges[[1]]
  res <- readRange(range, sheet, colClasses="numeric")

  cat("Make a new range ...")
  firstCell <- sheet$getRow(14L)$getCell(4L)
  lastCell  <- sheet$getRow(20L)$getCell(7L)
  rangeName <- "Test2"
  createRange(rangeName, firstCell, lastCell)
  
  if (length(getRanges(wb)) != 2){
    cat("STOP!!! Range not created!")
  } else {
    cat("OK\n")
  }
  
  
  
  
  
  
  
  
}








