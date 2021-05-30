package util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromExcel {

	static String projectPath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	static int rowCount;
	static int colCount;	 

	public static void main(String[] args)  {
		getRowCount();
		getColCount();
		GetCellDataString();
		GetCellDataNumber();

	}	
	
	public static void getRowCount() {

		try {
			String projectPath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(projectPath + "/Excel Files/data.xlsx");
			sheet = workbook.getSheet("Sheet1");
			rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("No of Rows : " + rowCount);

		} catch (IOException exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
	}

	public static void getColCount() {

		try {
			String projectPath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(projectPath + "/Excel Files/data.xlsx");
			sheet = workbook.getSheet("Sheet1");
			colCount = sheet.getRow(0).getPhysicalNumberOfCells();
			System.out.println("No of Columns : " + colCount);

		} catch (IOException exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
	}

	public static void GetCellDataString() {

		try {
			String projectPath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(projectPath + "/Excel Files/data.xlsx");
			sheet = workbook.getSheet("Sheet1");
			String cellData = sheet.getRow(1).getCell(0).getStringCellValue();
			System.out.println(cellData);

		} catch (IOException exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}

	}

	public static void GetCellDataNumber() {

		try {
			String projectPath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(projectPath + "/Excel Files/data.xlsx");
			sheet = workbook.getSheet("Sheet1");
			double cellData = sheet.getRow(1).getCell(2).getNumericCellValue();
			System.out.println(cellData);

		} catch (IOException exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}

	}

}
