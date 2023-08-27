package Helpers.DataProvider;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class ExcelReader {

    public ExcelReader(){

    }

    public void readDataFromExcel(String filePath, int sheetIndex, int startRow){
        try
        {
            File file = new File(filePath);
            FileInputStream fileInputStream = new FileInputStream(file);
            XSSFWorkbook wb = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = wb.getSheetAt(sheetIndex);
            int numOfRows = sheet.getLastRowNum();

        }catch (Exception e){
            System.out.println(String.format("Error: %s", e.getMessage()));
        }
    }
}
