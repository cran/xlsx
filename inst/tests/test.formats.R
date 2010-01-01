# test different formats
#
#

test.formats <- function(outdir="C:/Temp/")
{  
  require(xlsx)

  ##########################################################################
  cat("TEST COLORS, FILLS, BORDERS ...\n")
  ##########################################################################

  wb <- createWorkbook()
  sheet <- createSheet(wb, "Sheet1")

  rows  <- createRow(sheet, 1:10)          # 10 rows

  cell.1 <- createCell(rows[1], colIndex=1)[[1,1]]     
  setCellValue(cell.1, "Thick bottom border, lavender background")

  cat("Set cell [1,1] with bottom thick border, fill with lavender.\n")
  cellStyle1 <- createCellStyle(wb, borderPosition="BOTTOM",
    borderPen="BORDER_THICK", fillBackgroundColor="lavender",
    fillForegroundColor="lavender", fillPattern="SOLID_FOREGROUND")
  setCellStyle(cell.1, cellStyle1)   # does not set the background?!

  cell.2 <- createCell(rows[3], col=1)[[1,1]]      
  setCellValue(cell.2, "Dashed right border, yellow fill, red hashes")
  cellStyle2 <- createCellStyle(wb, borderPosition="RIGHT",
    borderPen="BORDER_DASHED", fillBackgroundColor="yellow",
    fillForegroundColor="tomato", fillPattern="BIG_SPOTS")

  setCellStyle(cell.2, cellStyle2)   # does not set the background?!


  cell.3 <- createCell(rows[5], col=1)[[1,1]]      
  setCellValue(cell.3, "Courier New, Italicised")

  font3 <- createFont(wb, fontHeightInPoints=20, isBold=TRUE, isItalic=TRUE,
    fontName="Courier New", color="tomato")
  cellStyle3 <- createCellStyle(wb, fillBackgroundColor="skyblue",
    fillForegroundColor="skyblue", fillPattern="SOLID_FOREGROUND",
    font=font3)

  setCellStyle(cell.3, cellStyle3)   # does not set the background?!

  autoSizeColumn(sheet, 1)  # test autosizing

  file <- paste(outdir, "test_export_colors.xlsx", sep="")
  saveWorkbook(wb, file)
  cat("Wrote file ", file, "\n\n")


  #########################################################################
  #  Test comments -- DOES NOT WORK YET.  KNOWN BUG with POI!
  #########################################################################

  # add a comment
  #cmnt <- createComment("Does it work?", 1, 3, sheet, author="Plato")
  #getCellComment(1,3,sheet)

##   x <- "Does it work??!"; row <- 0; col <- 0
##   cHelper <- .jcall(wb, "Lorg/apache/poi/xssf/usermodel/XSSFCreationHelper;",
##     "getCreationHelper")
##   rts <- .jcall(cHelper, "Lorg/apache/poi/ss/usermodel/RichTextString;",
##     "createRichTextString", x)
##   cmnt2 <- createComment(rts, row, col, sheet, author="Plato")
  
}


## wb <- createWorkbook()
## sheet <- createSheet(wb, "Sheet1")
## rows  <- createRow(sheet, 1:10)          # 10 rows

## cells <- createCell(rows, cell=1:5)      # 5 columns

## # create a test data.frame
## data <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
##   date=seq(as.Date("2009-01-01"), by="1 month", length.out=10),
##   bool=ifelse(1:10 %% 2, TRUE, FALSE))

## # set value of one cell by hand 
## setCellValue(cells[[1,1]], data[1,1])

## # or do them all by looping over columns
## for (ic in 1:ncol(data))
##   mapply(setCellValue, cells[,ic], data[,ic])


## file <- "C:/Temp/test_export.xlsx"
## saveWorkbook(wb, file)


## ########################## test dates and datetime objects #############

## wb <- createWorkbook()
## sheet <- createSheet(wb, "Sheet1")
## rows  <- createRow(sheet, 1)          # 1 rows

## cells <- createCell(rows, cell=1)      # 1 columns

## setCellValue(cells[[1,1]], as.Date("2009-11-16"))

## file <- "C:/Temp/test_export.xlsx"
## saveWorkbook(wb, file)








