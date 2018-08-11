package test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class NewTest 
{
	XSSFWorkbook workbook;
	int count = 0;
	@BeforeTest
	public void beforeTest() {
		System.out.println("Before Test");
	}
	
	@AfterTest
	public void afterTest() {
		System.out.println("After Test");
	}
	
	@BeforeClass
	public void beforeClass() {
		System.out.println("Before Class");
	}
	
	@AfterClass
	public void afterClass() {
		System.out.println("After Class");
	}
	
	@BeforeMethod
	public void beforeMethod() {
		System.out.println("Before Method");
	}
	
	@AfterMethod
	public void afterMethod() {
		System.out.println("After Method");
	}
	
	@BeforeSuite
	public void beforeSuite() {
		System.out.println("Before Suite");
	}
	
	@AfterSuite
	public void afterSuite() throws FileNotFoundException, IOException {
		System.out.println("After Suite");
		FileOutputStream fo = new FileOutputStream("D:\\test1.xlsx"); 
		workbook.write(fo);
		workbook.close();
		fo.close();
	}
	
	@Test(dataProvider="testProvider")
	public void test1(Double num1,Double num2) {
		XSSFSheet sheet = workbook.getSheetAt(0);
		XSSFRow row = sheet.getRow(count);
		XSSFCell cell = row.createCell(3);
		System.out.println("Test 1");
		count++;
		Double res = num1 + num2;
		cell.setCellValue(res);
	}
	
	
	@DataProvider(name="testProvider")
	public Object[][] provider() throws InvalidFormatException, IOException {
		Object[][] data = new Object[2][2];
		workbook = new XSSFWorkbook(ExcelUtil.file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		for(int i = 0; i<2; i++) {
			XSSFRow row = sheet.getRow(i);
			for(int j = 0; j<2; j++) {
				XSSFCell cell = row.getCell(j);
				data[i][j] = cell.getNumericCellValue();
			}
		}
		return data;
	}
}
