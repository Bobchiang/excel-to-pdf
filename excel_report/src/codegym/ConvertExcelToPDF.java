package codegym;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class ConvertExcelToPDF {
	public static void main(String[] args) throws IOException, DocumentException {
		Workbook workbook = new XSSFWorkbook("/Users/bobchiang/tmp/test.xlsx");
		Sheet sheet = workbook.getSheetAt(0); // get the first sheet

		// iterate through the rows
		for (Row row : sheet) {
			// iterate through the cells
			for (Cell cell : row) {
				// do something with the cell value
			}
		}

		Document pdfDoc = new Document();
		PdfWriter pdfWriter = PdfWriter.getInstance(pdfDoc, new FileOutputStream("/Users/bobchiang/tmp/read.pdf"));
		pdfDoc.open();

		// create a table
		PdfPTable table = new PdfPTable(sheet.getRow(0).getLastCellNum());
		table.setWidthPercentage(100);

		// add the header row
//		for (Cell headerCell : sheet.getRow(0)) {
//			table.addCell(new PdfPCell(
//					new Phrase(headerCell.getStringCellValue(), new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD))));
//		}
		


		// add the data rows
		for (Row dataRow : sheet) {
			for (Cell dataCell : dataRow) {
				dataCell.setCellType(CellType.STRING);
				CellStyle cellStyle = dataCell.getCellStyle();
				XSSFColor backgroundColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
				System.out.println(backgroundColor);

				if (backgroundColor != null) {
					byte[] rgb = backgroundColor.getRGB();
				    int red = (int) (rgb[0] & 0xff);
				    int green = (int) (rgb[1] & 0xff);
				    int blue = (int) (rgb[2] & 0xff);
				    BaseColor bgColor = new BaseColor(red, green, blue);
				    PdfPCell pdfCell = new PdfPCell(new Phrase(dataCell.getStringCellValue(), new Font(Font.FontFamily.HELVETICA, 10)));
				    pdfCell.setBackgroundColor(bgColor);
				    table.addCell(pdfCell);
				} else {
					table.addCell(new PdfPCell(new Phrase(dataCell.getStringCellValue(), new Font(Font.FontFamily.HELVETICA, 10))));
				}
			}
		}

		pdfDoc.add(table);
		pdfDoc.close();

	}
}
