package org.frame;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BAseclass {

	public static WebDriver driver;
	public static Actions a;
	public static Robot r;
	public static Select s;
	public static void launchChrome() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
	}
	public void loadUrl(String Url) {
		driver.get("Url");
	}
	public static void printTitle() {
		driver.getTitle();

	}
	public static void printCurrenturl() {
		driver.getCurrentUrl();
	}
	public static void winMax() {
		driver.manage().window().maximize();
	}

	public static void fill(WebElement ele, String Value) {

		ele.sendKeys(Value);
	}
	public static void btnClick(WebElement e) {

		e.click();
	}

	public static void dClick() {

		a= new Actions(driver);
		a.doubleClick().perform();
	}
	public static void rClick() {
		a= new Actions(driver);
		a.contextClick().perform();
	}
	public static void MoveElement(WebElement n) {
		a= new Actions(driver);
		a.moveToElement(n).perform();
	}
	
	public static void selectAll() {
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_A);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_A);
	}
	public static void downKey() {
		r.keyPress(KeyEvent.VK_DOWN);
		r.keyRelease(KeyEvent.VK_DOWN);
	}
	public static void upKey() {
		r.keyPress(KeyEvent.VK_UP);
		r.keyRelease(KeyEvent.VK_UP);
	}

	public static void enter() {
		r.keyPress(KeyEvent.VK_ENTER);
		r.keyRelease(KeyEvent.VK_ENTER);
	}
	public static void copy() {
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_C);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_C);
	}
	public static void paste() {
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_V);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_V);
	}
	public static void cut() {
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_X);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_X);
	}
	public static void byValue(WebElement ref, String value) {
		s= new Select(ref);
		s.selectByValue(value);

	}
	public static void byIndex(WebElement ref, int n) {
		s= new Select(ref);
		s.selectByIndex(n);
	}
	public static void byVisibletext(WebElement ref, String value) {
		s= new Select(ref);
		s.selectByVisibleText(value);
	}
	public static void isMultiple(WebElement ref) {
		s= new Select(ref);
		boolean multiple = s.isMultiple();
		System.out.println(multiple);
	}
	public static void options(WebElement ref) {
		s=new Select(ref);
		List<WebElement> li = s.getOptions();

		for (int i = 0; i < li.size(); i++) {
			WebElement ele = li.get(i);
			String text = ele.getText();
			System.out.println(text);
		}
	}
	public static void allOptions(WebElement ref) {
		s=new Select(ref);
		List<WebElement> allselec = s.getAllSelectedOptions();
		for (WebElement x : allselec) {
			String text = x.getText();
			System.out.println(text);
		}
	}
	public static void firstOptions(WebElement ref) {
		s= new Select(ref);
		WebElement ele = s.getFirstSelectedOption();
		String text = ele.getText();
		System.out.println(text);

	}
	public static void deselectByindex(WebElement ref,int a) {

		s= new Select(ref);
		s.deselectByIndex(a);
	}
	public void deselectAll(WebElement ref) {
		s= new Select(ref);
		s.deselectAll();
	}
	public static void deselectValue(WebElement ref, String value) {
		s= new Select(ref);
		s.deselectByValue(value);
	}
	public static void deselectVisibText(WebElement ref, String value) {
		s= new Select(ref);
		s.deselectByVisibleText(value);
	}
	public static void setValue(WebElement ref, String value) {
		JavascriptExecutor jk = (JavascriptExecutor) driver;
		jk.executeScript("arguments[0].setAttributre('value,'value')", ref);

	}
	public static void scrollDown(WebElement ref) {
		JavascriptExecutor jk = (JavascriptExecutor) driver;
		jk.executeScript("arguments[0].scrollIntoView('false')", ref);
	}
	public static void scrollUp(WebElement ref) {
		JavascriptExecutor jk = (JavascriptExecutor) driver;
		jk.executeScript("arguments[0].scrollIntoView('true')", ref);
	}
	public static  String getData(int rowno, int colno) throws IOException {

		File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Datadriven\\target\\Excel Data\\details.xlsx");
		FileInputStream fin = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fin);
		Sheet s = w.getSheet("details");
		Row row = s.getRow(rowno);
		Cell cell = row.getCell(colno);
		int cellType = cell.getCellType();
		String value = "";
		if (cellType==1) {
			value = cell.getStringCellValue();	
		}
		else if (cellType==0) {
			
			if (DateUtil.isCellDateFormatted(cell)) {
				Date d = cell.getDateCellValue();
				SimpleDateFormat sim = new SimpleDateFormat("MM-dd-yyyy");
				value=sim.format(d);
			}
			else {
				double d = cell.getNumericCellValue();
				long l = (long) d;
				 value = String.valueOf(l);
			}
		}
		return value;
	}
	public static void newFile(String n,int r,int c,String v) throws IOException {
		File f = new File("C:\\Users\\SENTHIL\\eclipse-workspace\\Sample\\target\\Excel\\details.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet(n);
		Row row = s.createRow(r);
		Cell cell = row.createCell(c);
		cell.setCellValue(v);
		FileOutputStream fout = new FileOutputStream(f);
		w.write(fout);
		System.out.println("Sucess : ("+r+ "," +c+ ")");
	}
	public static void closeChrome() {
driver.close();
	}
	public static void printstartTime() {
		Date d = new Date();
		System.out.println(d);

	}
	public static void printendTime() {
		Date d = new Date();
		System.out.println(d);
	}
}
