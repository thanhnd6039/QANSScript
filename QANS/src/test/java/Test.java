import Helpers.DataProvider.ExcelReader;
import org.apache.commons.collections4.list.SetUniqueList;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;


public class Test {
    public static void main(String[] args) {
        String filePath = "C:\\CucumberFramework\\Downloads\\Output.xlsx";
        try {
            File file = new File(filePath);
            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet("Sheet1");
            FileOutputStream fileOut = new FileOutputStream(file);
            wb.write(fileOut);
            fileOut.close();
            System.out.println("Excel File has been created successfully.");
        }catch (Exception e){
            System.out.println(String.format("Error: %s", e.getMessage()));
        }

    }
}
