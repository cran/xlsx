# test different data formats
#
#

test.dataFormats <- function()
{  
  require(xlsx)
  Sys.setenv(tz="")

  # create a test data.frame
  data <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
    date=seq(as.Date("2009-01-01"), by="1 month", length.out=10),
    bool=ifelse(1:10 %% 2, TRUE, FALSE), log=log(1:10),
    rnorm=10000*rnorm(10),
    datetime=seq(as.POSIXct("2009-01-01 00:00:00"), by="1 hour", length.out=10))

  wb <- createWorkbook()
  sheet <- createSheet(wb, "Sheet1")
  rows  <- createRow(sheet, 1:10)          # 10 rows
  cells <- createCell(rows, col=1:8)       # 8 columns

  # or do them all by looping over columns
  for (ic in 1:ncol(data))
    mapply(setCellValue, cells[,ic], data[,ic]) 

  # format "log" column with two decimals
  cellStyle1 <- createCellStyle(wb, dataFormat="#,##0.00")
  #cellStyle1$getDataFormat()
  lapply(cells[,6], setCellStyle, cellStyle1)

  # format date column
  cellStyle2 <- createCellStyle(wb, dataFormat="m/d/yyyy")
  #cellStyle2$getDataFormat()
  lapply(cells[,4], setCellStyle, cellStyle2)

  # format datetime column
  cellStyle3 <- createCellStyle(wb, dataFormat="m/d/yyyy h:mm:ss;@")
  #cellStyle2$getDataFormat()
  lapply(cells[,8], setCellStyle, cellStyle3)

  # format "rnorm" column with two decimals, comma separator, red
  cellStyle4 <- createCellStyle(wb, dataFormat="#,##0.00_);[Red](#,##0.00)")
  #cellStyle1$getDataFormat()
  lapply(cells[,7], setCellStyle, cellStyle4)

  file <- "C:/Temp/test_dataFormats.xlsx"
  saveWorkbook(wb, file)

}











