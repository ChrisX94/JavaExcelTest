package TestPOI.ExcelTest;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class TestDemo {

	public static void main(String[] args) {
		
		String filePath = "/Users/chrisxiao/Desktop/Tibame/Java_Hahow/Material/read.xlsx";
//		
//		File file = new File(filePath);
//		System.out.println(file.getName());
//		
//		
//		try {
//			ExcelUtil.readFile(filePath);
//		}catch(Exception e) {
//			e.printStackTrace();
//		}

		String writeFileName = "/Users/chrisxiao/Desktop/Tibame/Java_Hahow/Material/write.xlsx";
		DataHelp dataHelp = new DataHelpImp(); // 這裡是模擬取得資料的物件
		List<String[]> list = dataHelp.getData(); 
		try {
			ExcelUtil.writeFile(writeFileName, list);
		}catch(Exception e) {
			e.printStackTrace();
		}
		
		System.out.println("End of Data");
		
	}

}
