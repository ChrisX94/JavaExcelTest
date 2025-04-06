package TestPOI.ExcelTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	public static void readFile(String fileName) throws IOException {
		InputStream is = new FileInputStream(new File(fileName));
		XSSFWorkbook wb = new XSSFWorkbook(is); // workbook
		XSSFSheet sheet = wb.getSheetAt(0); // read the first Sheet
		Iterator itRow = sheet.rowIterator(); // iterate rows

		while (itRow.hasNext()) {
			XSSFRow row = (XSSFRow) itRow.next(); // casting the elements of itRow(Object) to type (XSSFRow)
			Iterator itCell = row.cellIterator(); // iterate the cell

			while (itCell.hasNext()) {
				XSSFCell cell = (XSSFCell) itCell.next();
				System.out.print(cell + " ");
			}

			System.out.println();
		}
	}

	public static void writeFile(String fileName, List<String[]> dataList) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook(); //建立一個活頁簿物件
		XSSFSheet sheet = wb.createSheet(); // 建立一個Sheet物件(sheetName)
		XSSFRow titleRow = sheet.createRow(0); // 建立第一行 (0)是第一個(欄位名稱欄)
		
		XSSFCell cell1 = titleRow.createCell(0); // 第(1,1)的物件
		cell1.setCellValue("id"); //傳入儲存格資料
		
		XSSFCell cell2 = titleRow.createCell(1); // 第(1,2)的物件
		cell2.setCellValue("name"); //傳入儲存格資料

		XSSFCell cell3 = titleRow.createCell(2); // 第(1,3)的物件
		cell3.setCellValue("age"); //傳入儲存格資料
		
		// Rows
		for(int i =0; i < dataList.size(); i++) { // 將dataList中的資料寫入cell
			XSSFRow row = sheet.createRow(i+1); // +1是因為要從第二行開始
			String[] rowData  = dataList.get(i); // 取得第i個row的資料
			// Columns
			for(int j =0; j < rowData.length; j++) {
				XSSFCell cell = row.createCell(j); // create a cell the row
				cell.setCellValue(rowData[j]); // set value for cell
			}
		}
		// 寫入檔案
		FileOutputStream fis = new FileOutputStream(fileName);
		wb.write(fis);
		fis.flush();
		fis.close();
		
		

	}

}
