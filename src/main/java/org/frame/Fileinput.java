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

public class Fileinput {

	public static void main(String[] args) throws IOException {
	
		File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Datadriven\\target\\Excel Data\\FB.xlsx");
		FileInputStream fin = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fin);
		Sheet s = w.createSheet();
		Row r = s.createRow(6);
		Cell c = r.createCell(1);
		c.setCellValue("mike");
		FileOutputStream fout = new FileOutputStream(f);
		w.write(fout);

	}

}
