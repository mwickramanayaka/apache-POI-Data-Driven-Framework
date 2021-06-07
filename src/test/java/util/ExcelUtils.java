package util;

public class ExcelUtils {

	public static void main(String[] args) {

		String projectPath = System.getProperty("user.dir");
		ExcelUtilTest excel = new ExcelUtilTest(projectPath + "/Excel Files/data.xlsx", "Sheet1");
		
		excel.getRowCount();
		excel.getColCount();
		excel.GetCellDataString(1, 1);
		excel.GetCellDataNumber(3, 2);
		
		
	}

}
