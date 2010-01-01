######################################################################
# Create a cell style.  It needs a workbook object!
#
createCellStyle <- function(wb, hAlign=NULL, vAlign=NULL, borderPosition=NULL,
  borderPen="BORDER_NONE", borderColor=NULL, fillBackgroundColor=NULL,
  fillForegroundColor=NULL, fillPattern=NULL, font=NULL, dataFormat=NULL)
{
  cellStyle <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/CellStyle;",
    "createCellStyle") 

  if (!is.null(hAlign))
    .jcall(cellStyle, "V", "setAlignment", xlsx:::.CELL_STYLES[hAlign])

  if (!is.null(vAlign))
    .jcall(cellStyle, "V", "setVerticalAlignment", xlsx:::.CELL_STYLES[vAlign])
#browser()
  if (!is.null(borderPosition)){
    switch(borderPosition,
      BOTTOM = .jcall(cellStyle, "V", "setBorderBottom",
        .jshort(xlsx:::.CELL_STYLES[borderPen])), 
      LEFT   = .jcall(cellStyle, "V", "setBorderLeft",
        .jshort(xlsx:::.CELL_STYLES[borderPen])),
      TOP    = .jcall(cellStyle, "V", "setBorderTop",
        .jshort(xlsx:::.CELL_STYLES[borderPen])),
      RIGHT  = .jcall(cellStyle, "V", "setBorderRight",
        .jshort(xlsx:::.CELL_STYLES[borderPen])))
  }

  if (!is.null(borderColor))
     .jcall(cellStyle, "V", "setBorderColor", .xssfcolor(borderColor))
  
  if (!is.null(fillBackgroundColor))
    .jcall(cellStyle, "V", "setFillBackgroundColor",
       .xssfcolor(fillBackgroundColor))
  
  if (!is.null(fillForegroundColor))
    .jcall(cellStyle, "V", "setFillForegroundColor",
       .xssfcolor(fillForegroundColor))
  
  if (!is.null(fillPattern))
    .jcall(cellStyle, "V", "setFillPattern",
       .jshort(xlsx:::.CELL_STYLES[fillPattern]))

  if (!is.null(font))
    .jcall(cellStyle, "V", "setFont", font)
  
  if (!is.null(dataFormat)){
    fmt <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/DataFormat;",
             "createDataFormat")
    .jcall(cellStyle, "V", "setDataFormat",
           .jshort(fmt$getFormat(dataFormat)))
  }
  
  cellStyle
}


######################################################################
# Set the cell style for one cell. 
# Only one cell and one value.
#
setCellStyle <- function(cell, cellStyle)
{ 
  .jcall(cell, "V", "setCellStyle", cellStyle)
  invisible(NULL)
}

######################################################################
# Get the cell style for one cell. 
# Only one cell and one value.
#
getCellStyle <- function(cell)
{ 
  .jcall(cell,  "Lorg/apache/poi/ss/usermodel/CellStyle;", "getCellStyle")
}

