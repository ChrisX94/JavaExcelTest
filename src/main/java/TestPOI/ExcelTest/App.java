package TestPOI.ExcelTest;

import org.apache.poi.xssf.usermodel.XSSFWorkbook; // working on Excel workbook
import org.apache.poi.xssf.usermodel.XSSFSheet; // working on sheet
import java.io.FileInputStream; // input file
import java.io.FileOutputStream; // output file
import java.io.IOException; // Exception for I/O 


public class App 
{
    public static void main( String[] args )
    {
        String inputPath = "/Users/chrisxiao/Desktop/Tibame/Java_Hahow/Material/test/input.xlsx";
        String templatePath= "/Users/chrisxiao/Desktop/Tibame/Java_Hahow/Material/test/template.xlsx";
        String reportPath = "/Users/chrisxiao/Desktop/Tibame/Java_Hahow/Material/test/report.xlsx";
        
        try
        {
        	FileInputStream inputFile = new FileInputStream(inputPath); // load the input file
        	XSSFWorkbook inputWorkbook = new XSSFWorkbook(inputFile); // parse workbook
        	XSSFSheet inputSheet = inputWorkbook.getSheetAt(0); // get the first sheet index(0)
        	
        	FileInputStream templateFile = new FileInputStream(templatePath); // load the template file
        	XSSFWorkbook templateWorkbook = new XSSFWorkbook(templateFile); // parse workbook
        	XSSFSheet templateSheet = templateWorkbook.getSheetAt(0); // get the first sheet index(0)
        	
        	fillTemplate(inputSheet, templateSheet);
        	
        	FileOutputStream outputFile = new FileOutputStream(reportPath); // file output
        	templateWorkbook.write(outputFile); // use templateWorkbook write the data into output file
        	
        	inputFile.close(); 
        	templateFile.close();
        	outputFile.close();
        	
        	System.out.println("Report created");
        	
        }catch(IOException e){
        	e.printStackTrace();
        	}
        
    }
    
    
    private static void fillTemplate(XSSFSheet inputSheet, XSSFSheet templateSheet) {
    	templateSheet.getRow(17).getCell(2).setCellValue( // input value into templateSheet(cell:18, 3)
    			inputSheet.getRow(1).getCell(1).getNumericCellValue()); // get the value in the input sheet(cell : 2,2)
    	
    	templateSheet.getRow(17).getCell(3).setCellValue( // input value into templateSheet(cell:18, 3)
    			inputSheet.getRow(1).getCell(2).getNumericCellValue()); // get the value in the input sheet(cell : 2,2)
    
    	templateSheet.getRow(18).getCell(2).setCellValue( // input value into templateSheet(cell:18, 3)
    			inputSheet.getRow(2).getCell(1).getNumericCellValue()); // get the value in the input sheet(cell : 2,2)
    	templateSheet.getRow(18).getCell(3).setCellValue( // input value into templateSheet(cell:18, 3)
    			inputSheet.getRow(2).getCell(2).getNumericCellValue()); // get the value in the input sheet(cell : 2,2)
        
    	templateSheet.getRow(19).getCell(2).setCellValue( // input value into templateSheet(cell:18, 3)
    			inputSheet.getRow(3).getCell(1).getNumericCellValue()); // get the value in the input sheet(cell : 2,2)
    	templateSheet.getRow(19).getCell(3).setCellValue( // input value into templateSheet(cell:18, 3)
    			inputSheet.getRow(3).getCell(2).getNumericCellValue()); // get the value in the input sheet(cell : 2,2)
        
    }
}
