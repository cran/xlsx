######################################################################
# Create some cells in the row object.  "cell" is the index of columns. 
# You can pass in a list of rows.
# Return a matrix of lists.  Each element is a cell object.
createCell <- function(row, colIndex=1:5)
{
  cells <- matrix(list(), nrow=length(row), ncol=length(colIndex),
    dimnames=list(names(row), colIndex))
  
  for (ir in seq_along(row))
    for (ic in seq_along(colIndex))
      cells[[ir,ic]] <- .jcall(row[[ir]], "Lorg/apache/poi/ss/usermodel/Cell;",
        "createCell", as.integer(colIndex[ic]-1))
    
  cells
}

######################################################################
# Get the cells for a list of rows.  Users who want basic things only
# don't need to use this function. 
# 
getCells <- function(row, simplify=TRUE)
{
  res <- row
  for (ir in seq_along(row)){
    minColIx <- .jcall(row[[ir]], "T", "getFirstCellNum")   # 0-based
    maxColIx <- .jcall(row[[ir]], "T", "getLastCellNum")-1  # 0-based
    indCol   <- seq.int(minColIx, maxColIx)                  # actual col index
    rowCells <- vector("list", length=(maxColIx-minColIx+1))
    namesCells <- vector("character", length=(maxColIx-minColIx+1))
    for (ic in seq_along(rowCells)){
      rowCells[[ic]] <- .jcall(row[[ir]],
        "Lorg/apache/poi/xssf/usermodel/XSSFCell;", "getCell", indCol[ic])
      namesCells[ic] <- .jcall(rowCells[[ic]], "I", "getColumnIndex")+1
    }
    names(rowCells) <- namesCells  # need namesCells if spreadsheet is ragged
    res[[ir]] <- rowCells
  }

  if (simplify)
    res <- unlist(res)
  
  res
}


######################################################################
# Only one cell and one value.  
# You vectorize outside this function if you want.
# Maybe I do a vectorized function inside java if this one is slow.
# 
#    Date    = .jnew("java/text/SimpleDateFormat",
#      "yyyy-MM-dd")$parse(as.character(value)),     # does not format it!
#
setCellValue <- function(cell, value, richTextString=FALSE)
{
  value <- switch(class(value)[1],
    integer = as.numeric(value),
    numeric = value,              
    Date    = as.numeric(value) + 25569,             # add Excel origin
    POSIXct = as.numeric(value)/86400 + 25569,               
    as.character(value))  # for factors and other types
#browser()

  if (richTextString)
    value <- .jnew("org/apache/poi/xssf/usermodel/XSSFRichTextString",
      as.character(value))  # do I need to convert to as.character ?!!
  
  invisible(.jcall(cell, "V", "setCellValue", value))
}


######################################################################
# get cell value. ONE cell only
#
getCellValue <- function(cell, keepFormulas=FALSE)
{
  cellType <- .jcall(cell, "I", "getCellType") + 1
  value <- switch(cellType,
    .jcall(cell, "D", "getNumericCellValue"),        # numeric              
    .jcall(.jcall(cell,
      "Lorg/apache/poi/xssf/usermodel/XSSFRichTextString;",
      "getRichStringCellValue"), "S", "toString"),   # string
    ifelse(keepFormulas, .jcall(cell, "S", "getCellFormula"),
      tryCatch(.jcall(cell, "D", "getNumericCellValue"), error=function(e) e,
               finally=NA)),   # formula
    NA,                                              # blank cell
    .jcall(cell, "Z", "getBooleanCellValue"),        # boolean
    .jcall(cell, "B", "getErrorCellValue"),          # error
    "Error"                                          # catch all
  ) 

  value
}


######################################################################
# create ONE comment
# x is a string not a rich text string.  
## createComment <- function(x, cell, sheet, author=NULL)
## {
##   cmnt <- .jcall(sheet, "Lorg/apache/poi/xssf/usermodel/XSSFComment;",
##                  "createComment")
##   .jcall(cmnt, "V", "setString", x) 
##   if (!is.null(author))
##     .jcall(cmnt, "V", "setAuthor", author)
## browser()

##   aux <- .jcall(sheet, "Lorg/apache/poi/ss/usermodel/Comment;",
##                  "createComment")
  
##   .jcall(cell, "V", "setCellComment", cmnt)

##   cmnt
## }

## createComment <- function(x, row, col, sheet, author=NULL)
## {
##   cmnt <- .jcall(sheet, "Lorg/apache/poi/xssf/usermodel/XSSFComment;",
##                  "createComment")
##   .jcall(cmnt, "V", "setString", x)
##   .jcall(cmnt, "V", "setRow", as.integer(row))
##   .jcall(cmnt, "V", "setColumn", .jshort(col))
  
##   if (!is.null(author))
##     .jcall(cmnt, "V", "setAuthor", author)

##   cmnt
## }

## getCellComment <- function(row, col, sheet)
## {
##   aux <- .jcall(sheet, "Lorg/apache/poi/xssf/usermodel/XSSFComment;",
##     "getCellComment", as.integer(row), as.integer(col))

##   cmnt <- structure(c(aux$getAuthor(), aux$getString()$toString()),
##     names=c("author", "comment"))

##   cmnt
## }
