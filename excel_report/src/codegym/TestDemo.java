package codegym;

import java.io.IOException;
import java.util.List;

public class TestDemo {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		// 讀取 Excel 資料

		String readFileName = "/Users/bobchiang/tmp/read.xlsx";
		try {
			ExcelUtil.readFile(readFileName);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// 寫入 Excel 資料
//		String writeFileName = "/Users/bobchiang/tmp/write.xlsx";
//		DataHelp datahelp = new DataHelpImp();
//		List<String[]> list = datahelp.getData();
//		try {
//			ExcelUtil.writeFile(writeFileName, list);
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		System.out.println("Write Excel End");
	}
}
