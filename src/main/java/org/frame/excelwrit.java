package org.frame;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelwrit {

	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Datadriven\\target\\Excel Data\\Family.xlsx");
		boolean fil = f.createNewFile();
		System.out.println(fil);
		
	Workbook w = new XSSFWorkbook();
	Sheet sheet = w.createSheet("family");
	Row row = sheet.createRow(0);
	Cell cell = row.createCell(0);
	cell.setCellValue("familia");
	
	Row r1 = sheet.createRow(1);
	Cell c1 = r1.createCell(1);
	c1.setCellValue("mike");
	
	Row r2 = sheet.createRow(0);
	Cell c2 = r2.createCell(1);
	c2.setCellValue("eight month");
	
	Row r3 = sheet.createRow(0);
	Cell c3 = r3.createCell(0);
	c3.setCellValue("no");
	
	
	FileOutputStream fout = new FileOutputStream(f);
	w.write(fout);
	}

}
