#
#
#

test.import <- function(outdir="C:/Temp/")
{
  require(xlsx)

  #######################################################################
  # READ an xlsx file into R

  cat("Test low level import ...\n")
  file <- paste(.Library, "/xlsx/tests/test_import.xlsx", sep="")

  wb <- loadWorkbook(file)
  cat("Workbook loaded\n")
  sheets <- getSheets(wb)

  sheet  <- sheets[[2]]
  rows   <- getRows(sheet)
  cells  <- getCells(rows)
  values <- lapply(cells, getCellValue)

  cat("Extract cell [5,2] and see if == 'Apr' ...")
  cell <- cells[["5.2"]]
  #cell <- cells[["2.5"]]
  if (getCellValue(cell) == "Apr")
    cat("OK\n")
    

  cat("Test high level import ... \n")
  res <- read.xlsx(file, 2)
  print(res)   # some cells are errors because of wrong formulas
  
  
  cat("Test high level import keeping formulas... \n")
  res <- read.xlsx(file, 2, keepFormulas=TRUE)
  print(res)   # show the formulas explicitly

  
  cat("Test high level with colClasses... \n")
  colClasses <- rep(NA, length=6); colClasses[5] <- "numeric"
  res <- read.xlsx(file, 2, colClasses=colClasses)
  print(res)   # force convesion to numeric


  #######################################################################
  # WRITE an xlsx file into R

  cat("Test high level export ... \n")
  x <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
    date=seq(as.Date("2009-01-01"), by="1 month", length.out=10),
    bool=ifelse(1:10 %% 2, TRUE, FALSE))

  file <- paste(outdir, "test_export.xlsx", sep="")
  write.xlsx(x, file)
  cat("Wrote file ", file, "\n")

  #write.xlsx(x, file, row.names=FALSE)  # without row.names
  
}










