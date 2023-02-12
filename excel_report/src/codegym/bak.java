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
	import org.apache.poi.xssf.usermodel.XSSFFont;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	import com.itextpdf.text.BaseColor;
	import com.itextpdf.text.Document;
	import com.itextpdf.text.DocumentException;
	import com.itextpdf.text.Font;
	import com.itextpdf.text.Phrase;
	import com.itextpdf.text.pdf.BaseFont;
	import com.itextpdf.text.pdf.PdfPCell;
	import com.itextpdf.text.pdf.PdfPTable;
	import com.itextpdf.text.pdf.PdfWriter;

	public class bak {
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
//			for (Cell headerCell : sheet.getRow(0)) {
//				table.addCell(new PdfPCell(
//						new Phrase(headerCell.getStringCellValue(), new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD))));
//			}

			// add the data rows
			for (Row dataRow : sheet) {
				for (Cell dataCell : dataRow) {
					dataCell.setCellType(CellType.STRING);
					CellStyle cellStyle = dataCell.getCellStyle();

					// background
					XSSFColor backgroundColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
					System.out.println(backgroundColor);

					// font color
					XSSFFont font = (XSSFFont) workbook.getFontAt(cellStyle.getFontIndex());
					XSSFColor fontColor = (XSSFColor) font.getXSSFColor();
					System.out.println(fontColor);

//					if (backgroundColor != null && fontColor == null) {
//						byte[] rgb1 = backgroundColor.getRGB();
//						int red1 = (int) (rgb1[0] & 0xff);
//						int green1 = (int) (rgb1[1] & 0xff);
//						int blue1 = (int) (rgb1[2] & 0xff);
//						BaseColor bgColor1 = new BaseColor(red1, green1, blue1);
//						PdfPCell pdfCell1 = new PdfPCell(
//								new Phrase(dataCell.getStringCellValue(), new Font(Font.FontFamily.HELVETICA, 10)));
//						pdfCell1.setBackgroundColor(bgColor1);
//						table.addCell(pdfCell1);
//					} else if (fontColor != null && backgroundColor == null) {
//						byte[] rgb2 = fontColor.getRGB();
//						int red2 = (int) (rgb2[0] & 0xff);
//						int green2 = (int) (rgb2[1] & 0xff);
//						int blue2 = (int) (rgb2[2] & 0xff);
//						BaseColor bgColor2 = new BaseColor(red2, green2, blue2);
//						String fontFamily = font.getFontName();
//						int fontSize = font.getFontHeightInPoints();
//						Font pdfFont = new Font(BaseFont.createFont(fontFamily, BaseFont.WINANSI, BaseFont.EMBEDDED), fontSize, Font.NORMAL, bgColor2);
//						PdfPCell pdfCell2 = new PdfPCell(new Phrase(dataCell.getStringCellValue(), pdfFont));
//						table.addCell(pdfCell2);
//					} else if (fontColor != null && backgroundColor != null) {
//						byte[] rgb_bg = backgroundColor.getRGB();
//						byte[] rgb_ft = fontColor.getRGB();
//						int red3 = (int) (rgb_bg[0] & 0xff);
//						int green3 = (int) (rgb_bg[1] & 0xff);
//						int blue3 = (int) (rgb_bg[2] & 0xff);
//						int red4 = (int) (rgb_ft[0] & 0xff);
//						int green4 = (int) (rgb_ft[1] & 0xff);
//						int blue4 = (int) (rgb_ft[2] & 0xff);
//						BaseColor bgColor3 = new BaseColor(red3, green3, blue3);
//						BaseColor bgColor4 = new BaseColor(red4, green4, blue4);
//						Font pdfFont2 = new Font(Font.FontFamily.valueOf(font.getFontName()), font.getFontHeightInPoints(),
//								Font.NORMAL, bgColor4);
//						PdfPCell pdfCell3 = new PdfPCell(new Phrase(dataCell.getStringCellValue(), pdfFont2));
//						pdfCell3.setBackgroundColor(bgColor3);
//						table.addCell(pdfCell3);
	//
//					} else {
//						table.addCell(new PdfPCell(
//								new Phrase(dataCell.getStringCellValue(), new Font(Font.FontFamily.HELVETICA, 10))));
					}
				}
			}

//			pdfDoc.add(table);
//			pdfDoc.close();

		}



}
