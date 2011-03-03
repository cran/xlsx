# test different formats
#
#

test.formats <- function(outdir="C:/Temp/", type="xlsx")
{  
  cat("##################################################\n")
  cat("Testing COLORS, FILLS, BORDERS ...\n")
  cat("##################################################\n")

  cat("Create an empty workbook ... ") 
  wb <- createWorkbook(type=type)
  sheet <- createSheet(wb, "Sheet1")
  cat("OK\n")

  
  cat("Create 10 rows ... ") 
  rows  <- createRow(sheet, 1:10)         
  cat("OK\n")

  
  cat("Create cell[1,1] ... ") 
  cell.1 <- createCell(rows[1], colIndex=1)[[1,1]]     
  setCellValue(cell.1, "Thick bottom border, lavender background")
  cat("OK\n")

  cat("Set cell [1,1] with bottom thick border, fill with lavender ... ")
  cellStyle1 <- createCellStyle(wb, borderPosition="BOTTOM",
    borderPen="BORDER_THICK", fillBackgroundColor="lavender",
    fillForegroundColor="lavender", fillPattern="SOLID_FOREGROUND")
  setCellStyle(cell.1, cellStyle1)   # does not set the background?!
  cat("OK\n")

  
  cat("Create cell[3,1] ... ") 
  cell.2 <- createCell(rows[3], col=1)[[1,1]]      
  setCellValue(cell.2, "Dashed right border, yellow fill, red hashes")
  cellStyle2 <- createCellStyle(wb, borderPosition="RIGHT",
    borderPen="BORDER_DASHED", fillBackgroundColor="yellow",
    fillForegroundColor="red", fillPattern="BIG_SPOTS")
  setCellStyle(cell.2, cellStyle2)   # does not set the background?!
  cat("OK\n")


  cat("Create cell[5,1] ... ") 
  cell.3 <- createCell(rows[5], col=1)[[1,1]]      
  setCellValue(cell.3, "Courier New, Italicised")
  font3 <- createFont(wb, fontHeightInPoints=20, isBold=TRUE, isItalic=TRUE,
    fontName="Courier New", color="orange")
  cellStyle3 <- createCellStyle(wb, fillBackgroundColor="turquoise",
    fillForegroundColor="turquoise", fillPattern="SOLID_FOREGROUND",
    font=font3)
  setCellStyle(cell.3, cellStyle3)   # does not set the background?!
  cat("OK\n")


  cat("check ALIGN_CENTER not working??")
  cell.4 <- createCell(rows[7], col=1)[[1,1]]      
  cellStyle4 <- createCellStyle(wb, hAlign="ALIGN_CENTER")
  setCellStyle(cell.4, cellStyle4)
  cat("OK\n")
  
  cat("Autosize first column ... ") 
  autoSizeColumn(sheet, 1)  # test autosizing
  cat("OK\n")

  file <- paste(outdir, "test_export_colors.", type, sep="")
  saveWorkbook(wb, file)
  cat("Wrote file ", file, "\n\n")
  cat("Please check if the file is OK\n")
  
}












