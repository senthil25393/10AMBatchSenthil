package org.frame;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Repalcecell {

	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Datadriven\\target\\Excel Data\\details.xlsx");
		
		FileInputStream fin = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(fin);
		
		Sheet s = w.getSheet("details");

		Row r = s.getRow(3);
		Cell c = r.getCell(0);
	
	
	}

}
