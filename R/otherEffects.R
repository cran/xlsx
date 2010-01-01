
######################################################################
# 
#
addMergedRegion <- function(sheet, startRow, endRow, startColumn,
  endColumn)
{

  cellRangeAddress <- .jnew("org/apache/poi/ss/util/CellRangeAddress",
    as.integer(startRow-1), as.integer(endRow-1), as.integer(startColumn-1),
    as.integer(endColumn-1))
  ind <- .jcall(sheet, "I", "addMergedRegion", cellRangeAddress)
    
  invisible(ind)  
}
removeMergedRegion <- function(sheet, ind)
  .jcall(sheet, "V", as.integer(ind))



######################################################################
# fit contents to column. col can be a vector of column indices.
#
autoSizeColumn <- function(sheet, colIndex)
{
  for (ic in (colIndex-1))
    .jcall(sheet, "V", "autoSizeColumn", as.integer(ic), TRUE)
  
    
  invisible()  
}

######################################################################
# fit contents to column. col can be a vector of column indices.
#
createFreezePane <- function(sheet, rowSplit, colSplit, startRow=NULL,
  startColumn=NULL)
{
  if (is.null(startRow) & is.null(startColumn)){
    .jcall(sheet, "V", "createFreezePane", as.integer(colSplit-1),
      as.integer(rowSplit-1))
  } else {   
    .jcall(sheet, "V", "createFreezePane", as.integer(colSplit-1),
      as.integer(rowSplit-1), as.integer(startColumn-1),
      as.integer(startRow-1))
  }
  
  invisible()  
}

######################################################################
# fit contents to column. col can be a vector of column indices.
#
createSplitPane <- function(sheet, xSplitPos=2000, ySplitPos=2000,
  startRow=1, startColumn=1, position="PANE_LOWER_LEFT")
{
  ind <- .jfield(sheet, "B", position)
  
  .jcall(sheet, "V", "createSplitPane", as.integer(xSplitPos),
    as.integer(ySplitPos), as.integer(startRow-1),
    as.integer(startColumn-1), ind)
    
  invisible()  
}



######################################################################
# set the zoom
#
setPrintArea <- function(wb, sheetIndex, startColumn, endColumn, startRow,
  endRow)
{
   .jcall(wb, "V", "setPrintArea", as.integer(sheetIndex-1),
     as.integer(startColumn-1),  as.integer(endColumn-1),
     as.integer(startRow-1),  as.integer(endRow-1))
}



######################################################################
# set the zoom
#
setZoom <- function(sheet, numerator=100, denominator=100)
{
  .jcall(sheet, "V", "setZoom", as.integer(numerator),
         as.integer(denominator))
    
  invisible()  
}


