package codegym;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	public static void readFile(String fileName) throws IOException {
		InputStream is = new FileInputStream(new File(fileName));
		XSSFWorkbook wb = new XSSFWorkbook(is);
		XSSFSheet sheet = wb.getSheetAt(0);
		Iterator itRow = sheet.rowIterator();

		while (itRow.hasNext()) {
			XSSFRow row = (XSSFRow) itRow.next();
			Iterator itCell = row.cellIterator();

			while (itCell.hasNext()) {
				XSSFCell cell = (XSSFCell) itCell.next();
				System.out.print(cell + " ");
			}
			System.out.println("");
		}
	}

	public static void writeFile(String fileName, List<String[]> dataList) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Test Sheet");
		XSSFRow titleRow = sheet.createRow(0);

		XSSFCell cell1 = titleRow.createCell(0);
		cell1.setCellValue("id");
		XSSFCell cell2 = titleRow.createCell(1);
		cell2.setCellValue("name");
		XSSFCell cell3 = titleRow.createCell(2);
		cell3.setCellValue("age");

		for (int i = 0; i < dataList.size(); i++) {
			XSSFRow row = sheet.createRow(i + 1);
			String[] rowData = dataList.get(i);

			for (int j = 0; j < rowData.length; j++) {
				XSSFCell cell = row.createCell(j);
				cell.setCellValue(rowData[j]);

			}

		}
		FileOutputStream fis = new FileOutputStream(fileName);
		wb.write(fis);
		fis.flush();
		fis.close();

	}

}
