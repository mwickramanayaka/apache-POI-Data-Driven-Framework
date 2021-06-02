package util;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtilTest {

	static String projectPath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	public ExcelUtilTest (String excelPath, String sheetName ) {

		try {
			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheet(sheetName);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public int getRowCount() {
		int rowCount = 0;
		rowCount = sheet.getPhysicalNumberOfRows();
		//System.out.println("No of Rows : " + rowCount);
		return rowCount;
	}

	public int getColCount() {
		int colCount = 0;
		colCount = sheet.getRow(0).getPhysicalNumberOfCells();
		//System.out.println("No of Columns : " + colCount);
		return colCount;
	}

	public String GetCellDataString(int rowNum, int colNum) {
		String cellData = null;
		cellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
		//System.out.println(cellData);
		return cellData;

	}

	public double GetCellDataNumber(int rowNum, int colNum) {
		double cellData = 0;
		cellData = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
		//System.out.println(cellData);
		return cellData;

	}


}
