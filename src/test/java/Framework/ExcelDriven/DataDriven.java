package Framework.ExcelDriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

public class DataDriven {
	public List<String> getData(String testCaseName) throws IOException {

		// Identify testcases column by scaning the entire 1st row
		ArrayList<String> actualDataRow = new ArrayList<String>();

		FileInputStream fileInputStream = new FileInputStream(
				System.getProperty("user.dir") + "/src/main/java/resouces/DataDrivenTest.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

		int sheets = workbook.getNumberOfSheets();

		for (int i = 0; i < sheets; i++) {

			if (workbook.getSheetName(i).equalsIgnoreCase("TestData")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				Row firstRow = rows.next();
				Iterator<Cell> cell = firstRow.cellIterator();
				int k = 0;
				int column = 0;
				while (cell.hasNext()) {
					Cell value = cell.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						column = k;
					}

					k++;
				}
				System.out.println(column);
				// Once column is identified then scan entire testcase column to identify
				// purchase testcase row
				while (rows.hasNext()) {
					Row row = rows.next();
					if (row.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						// after you grab purchase testcase row pull all the data of that row and feed
						// into test

						Iterator<Cell> cv = row.cellIterator();
						while (cv.hasNext()) {

							Cell c = cv.next();

							if (c.getCellType() == CellType.STRING) {
								actualDataRow.add(c.getStringCellValue());
							} else {
								actualDataRow.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}

						}
					}

				}

			}
		}

		return actualDataRow;
	}

}
