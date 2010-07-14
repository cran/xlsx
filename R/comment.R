######################################################################
# create ONE comment
# x is a string not a rich text string.  

createCellComment <- function(cell, string,  
  author=NULL, visible=TRUE)
{
  sheet <- .jcall(cell,
    "Lorg/apache/poi/xssf/usermodel/XSSFSheet;", "getSheet")
  wb <- .jcall(sheet,
    "Lorg/apache/poi/xssf/usermodel/XSSFWorkbook;", "getWorkbook")

  factory <- .jcall(wb,
    "Lorg/apache/poi/xssf/usermodel/XSSFCreationHelper;", "getCreationHelper")

  anchor <- .jcall(factory,
    "Lorg/apache/poi/ss/usermodel/ClientAnchor;", "createClientAnchor")
  
  drawing <- .jcall(sheet,
    "Lorg/apache/poi/xssf/usermodel/XSSFDrawing;", "createDrawingPatriarch")
  
  comment <- .jcall(drawing,
    "Lorg/apache/poi/ss/usermodel/Comment;", "createCellComment", anchor)
  .jcall(comment, "V", "setString", string)
  if (!is.null(author))
    .jcall(comment, "V", "setAuthor", author)
  if (visible)
    .jcall(cell, "V", "setCellComment", comment)

  invisible(comment)
}


removeCellComment <- function(cell){
  .jcall(cell, "V", "removeCellComment")
}

## setCellComment <- function(cell, comment){
##   .jcall(cell, "V", "setCellComment", comment)
## }

getCellComment <- function(cell)
{
  comment <- .jcall(cell, "Lorg/apache/poi/xssf/usermodel/XSSFComment;",
    "getCellComment")

  comment
}



