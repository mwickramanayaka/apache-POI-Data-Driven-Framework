package util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelSheet {

	static String projectPath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	static int rowCount;
	static int colCount;
	static File excel;
	static FileInputStream fis;
	static FileOutputStream fout;

	public static void main(String[] args)  {
		projectPath = System.getProperty("user.dir");
		excel = new File(projectPath + "/Excel Files/data.xlsx");
		WriteExcel();

	}	

	public static void WriteExcel() {

		try {
			//new object to get sheet from byte code
			fis = new FileInputStream(excel);
			//new object   
			workbook = new XSSFWorkbook(fis);

			//get sheet by name
			sheet = workbook.getSheet("Sheet1");
			//get sheet by index
			//sheet = workbook.getSheetAt(0);

			int rowCount = sheet.getLastRowNum();
			System.out.println(rowCount);

			sheet.getRow(0).createCell(3).setCellValue("designation");
			sheet.getRow(1).createCell(3).setCellValue("Senior QA");

			sheet.getRow(2).createCell(3).setCellValue("Manager");
			sheet.getRow(3).createCell(3).setCellValue("Junior HR");

			fout = new FileOutputStream(excel);
			workbook.write(fout);
			//close the work book avoid memory leak
			workbook.close();

		} catch (IOException exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}

	}

}
