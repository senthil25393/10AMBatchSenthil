package org.frame;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Samp {
	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Datadriven\\target\\Excel Data\\details.xlsx");
		
		FileInputStream fin = new FileInputStream(f);
		
		XSSFWorkbook w = new XSSFWorkbook(fin);
		
		Sheet s = w.getSheet("details");
		
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			
			Row row = s.getRow(i);
			
			for (int j = 0; j < s.getPhysicalNumberOfRows(); j++) {
				Cell cell = row.getCell(j);
				
				int cellType = cell.getCellType();
				System.out.println(cellType);
			}
		}
	}

}
