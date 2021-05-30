package util;

public class ExcelDataProvider {

	public static void main(String[] args) {
		String projectPath = System.getProperty("user.dir");
		String excelPath = projectPath + "/Excel Files/data1.xlsx";

		testData(excelPath, "Sheet1");
	}	

	public static Object[][] testData(String exelPath, String sheetName) {

		ExcelUtilTest excel = new ExcelUtilTest (exelPath, sheetName);

		int rowCount = excel.getRowCount();
		int colCount = excel.getColCount();

		Object data[][] = new Object[rowCount-1][colCount];

		for (int i = 1; i < rowCount; i++) {
			for(int j = 0; j < colCount; j++) {
				String cellData = excel.GetCellDataString(i, j);
				System.out.print(cellData + " | ");
				data[i-1][j] = cellData;

			}
			System.out.println();
		}
		return data;
	}

}
