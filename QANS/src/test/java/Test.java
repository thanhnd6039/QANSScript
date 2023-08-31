import Helpers.DataProvider.ExcelReader;

import java.util.ArrayList;
import java.util.List;


public class Test {
    public static void main(String[] args) {
        String filePath = "C:\\CucumberFramework\\Downloads\\VTOEMGroup.xlsx";
        ExcelReader excelReader = new ExcelReader();
        List<Object[]> dataOfVTOEMGroup = new ArrayList<>();
        dataOfVTOEMGroup = excelReader.readDataFromExcel(filePath, 0, 1, 0);
        System.out.println("Num of rows: "+dataOfVTOEMGroup.size());

    }
}
