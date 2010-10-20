package tests;

import java.io.FileOutputStream;
import java.io.IOException;
//import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
//import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import dev.RInterface;


public class TestRInterface {
	public static void main(String[] args) throws IOException, InvalidFormatException {
		
	  // create the R interface
	  RInterface R = new RInterface();
      R.NCOLS = 15;
      R.NROWS = 13000;
      //R.CELL_ARRAY = new Cell[R.NROWS][R.NCOLS];
	  
      double[] X = new double[R.NROWS];
      for (int i = 0; i < R.NROWS; i++){
        X[i] = 0.123 + (double) i;
      }
      
      // create a new file
      FileOutputStream out = new FileOutputStream("C:/Temp/junk.xlsx");
      // create a new workbook
      Workbook wb = new XSSFWorkbook();
      // create a new sheet
      Sheet sheet = wb.createSheet();
      
      R.createCells(sheet, 0, 0);
      // write one column at a time
      for (int j = 0; j < R.NCOLS; j++) {
        R.writeColDoubles(sheet, 0, j, X);
      }
      
      // write the workbook, close the file handle
      wb.write(out);
      out.close();
      
	  
		
      System.out.println("Done!");

	}
	
}
