######################################################################
# Create a Font.  It needs a workbook object!
#  - color is an R color string.
#
createFont <- function(wb, color=NULL, fontHeightInPoints=NULL, fontName=NULL,
  isItalic=FALSE, isStrikeout=FALSE, isBold=FALSE, underline=NULL,
  boldweight=NULL) # , setFamily=NULL
{
  font <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/Font;",
    "createFont") 

  if (!is.null(color))
    .jcall(font, "V", "setColor", .xssfcolor(color))   

  if (!is.null(fontHeightInPoints))
    .jcall(font, "V", "setFontHeightInPoints", .jshort(fontHeightInPoints))

  if (!is.null(fontName))
    .jcall(font, "V", "setFontName", fontName)

  if (isItalic)
    .jcall(font, "V", "setItalic", TRUE)

  if (isStrikeout)
    .jcall(font, "V", "setStrikeout", TRUE)

  if (isBold)
    .jcall(font, "V", "setBold", TRUE)

  if (!is.null(boldweight))
    .jcall(font, "V", "setBoldweigth", .jshort(boldweight))

  
  font
}
