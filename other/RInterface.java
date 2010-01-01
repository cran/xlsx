package dev;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
//import org.apache.poi.ss.util.CellReference;

public class RInterface {
	
  
   public String getCellValue(Cell cell, boolean keepFormulas){
     
    String out="";
    switch(cell.getCellType()) {
      case Cell.CELL_TYPE_STRING:
      out = "S\t" + cell.getRichStringCellValue().getString();
      break;
    case Cell.CELL_TYPE_NUMERIC:
      if(DateUtil.isCellDateFormatted(cell)) {
        out = "D\t" + String.valueOf(cell.getDateCellValue().getTime());
      } else {                
        out = "N\t" + String.valueOf(cell.getNumericCellValue());
      }
      break;
    case Cell.CELL_TYPE_BOOLEAN:
      out = "B\t" + String.valueOf(cell.getBooleanCellValue());
      break;
    case Cell.CELL_TYPE_FORMULA:
      if (keepFormulas){
        out = "S\t" + cell.getCellFormula();
      } else {
        try{ 
          out = "N\t"+String.valueOf(cell.getNumericCellValue());
        } catch (Exception ex){
          out = "S\t";
        }
      }
      break;
    default:
      out="S\t";
    }
     
    return out;
   }
  
    /*
     * read Sheet.  It ignores missing cells.  So you need to pass the
     * cell coordinates.
     * To fix: Dates are shown as Java Dates.  
     */
	public String readSheet(String filename, int sheetIndex, int nrows, 
	  int skip, boolean keepFormulas) 
	  throws IOException, InvalidFormatException {
		
	  StringBuffer out = new StringBuffer();
		
	  InputStream inp = new FileInputStream(filename);

      Workbook wb = WorkbookFactory.create(inp);
      Sheet sheet = wb.getSheetAt(sheetIndex);
      int lastRowIndex = sheet.getLastRowNum(); 
      if (nrows < 1)
        nrows = lastRowIndex - skip + 1;        // get them all
      
      for (Row row : sheet) {
        int rowNum = row.getRowNum();
        if (rowNum < skip)                  // skip first skip rows
          continue;
        if (rowNum >= skip + nrows)
          break;
                  
        out.append(rowNum + "\t");    // write the row number first 
    	for (Cell cell : row) { 
    	  out.append(cell.getColumnIndex()+"\t");  // write column index
    	  out.append(getCellValue(cell, keepFormulas));
    	  out.append("\t");  // elements separated by tabs
    	}
    	out.append("\n");    // rows separated by new lines
      }

      return out.toString();
	}
	
}
