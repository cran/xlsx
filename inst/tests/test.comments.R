# test different data formats
#
#

test.comments <- function(dirout="C:/Temp/")
{
  cat("###################################################\n")
  cat("Start testing comments\n")

  wb <- createWorkbook()
  sheet1 <- createSheet(wb, "Sheet1")
  rows   <- createRow(sheet1, rowIndex=1:10)     # 10 rows
  cells  <- createCell(rows, colIndex=1:8)       # 8 columns

  cell1 <- cells[[1,1]]
  setCellValue(cell1, 1)   # add value 1 to cell A1

  createCellComment(cell1, "Cogito", author="Descartes")
  
  file <- paste(dirout, "test_comments.xlsx", sep="")
  saveWorkbook(wb, file)
  cat("Saved file", file, "\n")

  comment <- getCellComment(cell1)
  stopifnot(comment$getAuthor()=="Descartes")
  stopifnot(comment$getString()$toString()=="Cogito")
  

}


##   factory <- .jcall(wb,
##     "Lorg/apache/poi/xssf/usermodel/XSSFCreationHelper;", "getCreationHelper")

##   anchor <- .jcall(factory,
##     "Lorg/apache/poi/ss/usermodel/ClientAnchor;", "createClientAnchor")

  
##   drawing <- .jcall(sheet1,
##     "Lorg/apache/poi/xssf/usermodel/XSSFDrawing;", "createDrawingPatriarch")
  
##   comment1 <- .jcall(drawing,
##     "Lorg/apache/poi/ss/usermodel/Comment;", "createCellComment", anchor)
##   .jcall(comment1, "V", "setString", "Ergo sum")
##   .jcall(comment1, "V", "setAuthor", "Descartes")
##   .jcall(cells[[1,1]], "V", "setCellComment", comment1)









