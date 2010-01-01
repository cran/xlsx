package tests;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import dev.RInterface;


public class TestRInterface {
	public static void main(String[] args) throws IOException, InvalidFormatException {
		
	  RInterface R = new RInterface();
      
	  int sheetIndex = 1;
	  int skip  = 0;
	  int nrows = -1;
	  boolean keepFormulas = false;
	  String aux = R.readSheet("C:/Temp/test_import.xlsx", sheetIndex, nrows, 
	    skip, keepFormulas); 
	  
	  System.out.println(aux);

		
    System.out.println("Done!");

	}
	
}
