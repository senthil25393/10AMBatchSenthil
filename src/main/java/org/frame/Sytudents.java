package org.frame;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Sytudents {

	public static void main(String[] args) throws IOException {
File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Datadriven\\target\\Excel Data\\personal.xlsx");
FileInputStream fin = new FileInputStream(f);
Workbook w = new XSSFWorkbook(fin);
Sheet s = w.getSheet("students");
//		int row = s.getPhysicalNumberOfRows();
	//	System.out.println("physical no of rows is: "+row);
		//Row r = s.getRow(1);
//int cells = r.getPhysicalNumberOfCells();
//System.out.println("Physical no of cells is: "+cells);

for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
	
	Row r = s.getRow(i);
	for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
		Cell cell = r.getCell(j);
	
	System.out.println(cell);
	
	}
}
Row r1 = s.createRow(10);
Cell c = r1.createCell(0);
c.setCellValue("mic");

	FileOutputStream fout = new FileOutputStream(f);
	w.write(fout);
	
	
	
	
	}

}
