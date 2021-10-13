package org.frame;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
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

public class Iteration {

	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Datadriven\\target\\Excel Data\\details.xlsx");
		
		FileInputStream fin = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(fin);
		
		Sheet s = w.getSheet("details");
		
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				
				Cell cell = r.getCell(j);
				
				System.out.println(cell);
				
				int cellType = cell.getCellType();
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
					else
					{
						double d = cell.getNumericCellValue();
						long l = (long) d;
						String value = String.valueOf(1);
					System.out.println(value);
					
					}
				}
			}
			
		}

	}

}
