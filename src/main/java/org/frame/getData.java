package org.frame;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class getData {
	
	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Datadriven\\target\\Excel Data\\details.xlsx");
		
		FileInputStream fin = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(fin);
		
		Sheet sheet = w.getSheet("details");

	//int phyrow = sheet.getPhysicalNumberOfRows();
	//System.out.println("physical no of rows is :"+phyrow);
	
	Row row = sheet.getRow(5);	
	
//	int phycells = row.getPhysicalNumberOfCells();
	//System.out.println("physical no of cell is: "+phycells);
	
	System.out.println("****Itteration****");
	
	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
		Row row2 = sheet.getRow(i);
		
		for (int j = 0; j < row2.getPhysicalNumberOfCells(); j++) {
			Cell cell = row2.getCell(j);
			//System.out.println(cell);
			int cellType = cell.getCellType();
		//	System.out.println(cellType);
			//1------>String
			//0------->Date or numbers
			if (cellType==1) {
				
				String value = cell.getStringCellValue();
				System.out.println(value);
			}
			else if (cellType==0) {
				if (DateUtil.isCellDateFormatted(cell)) {
					Date d = cell.getDateCellValue();
					SimpleDateFormat sim = new SimpleDateFormat("MM-dd-yyyy");
					String value = sim.format(d);
					System.out.println(value);
				}
				else {
					double d = cell.getNumericCellValue();
					long l = (long) d;
					String value = String.valueOf(l);
					System.out.println(value);
					
				}
			}
		}
	}
	
	}

}
