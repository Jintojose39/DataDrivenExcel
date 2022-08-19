import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class dataDriven {

	// identify the testcases coloumn by scannning the entie first row
	// once coloumn is identified then scan entire testcase to purchase testcase row

	// after you grab purchase testcase row=pull all the data into the test

	public ArrayList<String> getData(String testcasename) throws IOException {

		ArrayList<String> a = new ArrayList<String>();
		FileInputStream fis = new FileInputStream("C://Users//JINTO//OneDrive//Documents//datademo.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				// identify the testcases coloumn by scannning the entire first row
				Iterator<Row> rows = sheet.iterator();// sheet is a collection of rows

				Row firstrow = rows.next();

				Iterator<Cell> ce = firstrow.cellIterator();// row is a collection of cell
				int k = 0;
				int coloumn = 0;
				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("testcases")) {
						coloumn = k;
					}
					k++;
				}
				System.out.println(coloumn);

				// once coloumn is identified then scan entire testcase to purchase testcase row

				while (rows.hasNext()) {
					Row r = rows.next();

					if (r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcasename)) {

						// after you grab purchase testcase row=pull all the data into the test

						Iterator<Cell> cv = r.cellIterator();

						while (cv.hasNext()) {
							// System.out.println(cv.next().getStringCellValue());

							Cell c = cv.next();

							if (c.getCellType() == CellType.STRING) {
								a.add(c.getStringCellValue());
							} else {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}

							// System.out.println();
						}

					}

				}
			}
		}
		return a;
	}

	// public static void main(String[] args) throws IOException {
}
