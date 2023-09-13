import Helpers.DataProvider.ExcelReader;
import Helpers.Manager.FileReaderManager;
import Pages.Reports.RSSaleGapAccountAssignmentPage;
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
        RSSaleGapAccountAssignmentPage rsSaleGapAccountAssignmentPage = new RSSaleGapAccountAssignmentPage();
        rsSaleGapAccountAssignmentPage.getAllOEMGroupByMainSaleRep();
        System.out.println("Done");
    }
}
