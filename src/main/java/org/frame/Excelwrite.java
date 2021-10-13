package org.frame;

import java.io.File;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelwrite {

	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Datadriven\\target\\Excel Data\\svspm.xlsx");
		boolean a = f.createNewFile();
		System.out.println(a);
	
		Workbook w = new XSSFWorkbook();
		
		Sheet sheet = w.createSheet("familia");
	}

}
