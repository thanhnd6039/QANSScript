package Pages.Reports;

import Helpers.Manager.FileReaderManager;
import Pages.CommonPage;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.List;

public class RSSaleGapReportPage{

    private String SGReportFilePath = "C:\\CucumberFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx";
    public RSSaleGapReportPage(){

    }
    public void getOEMGroupAndMainSalesRepFromSGReport(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfSGReport = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(SGReportFilePath, 0, 5, 2);
        for (int rowIndex = 0; rowIndex < dataOfSGReport.size(); rowIndex++){
            String oemGroupCol = dataOfSGReport.get(rowIndex)[1].toString().trim();
            String mainSalesRepCol = dataOfSGReport.get(rowIndex)[2].toString().trim();
            if (!oemGroupCol.isEmpty()){
                Object[] cols = new Object[2];
                cols[0] = oemGroupCol;
                cols[1] = mainSalesRepCol;
                table.add(cols);
            }
        }
        String output = "C:\\CucumberFramework\\Downloads\\Output.xlsx";
        Object[] headerCols = new Object[2];
        headerCols[0] = "OEM Group";
        headerCols[1] = "Main Sales Rep";
        FileReaderManager.getInstance().getExcelReader().getOutputFromData(output, headerCols, table);
    }
    public void getDataFromSGReport(int fromYear, int toYear, int fromQuarter, int toQuarter){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfSGReport = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(SGReportFilePath, 0, 5, 2);
        String oemGroup = "";
        String mainSR = "";
        String pn = "";
        float revQty = 0;
        int numOfQuarter = (toYear - fromYear + 1)*(toQuarter-fromQuarter+1);
        for (int rowIndex = 0; rowIndex < dataOfSGReport.size(); rowIndex++){
            String oemGroupCol = dataOfSGReport.get(rowIndex)[1].toString().trim();
            String mainSRCol = dataOfSGReport.get(rowIndex)[2].toString().trim();
            String pnCol = dataOfSGReport.get(rowIndex)[3].toString().trim();
            String revQtyCol = dataOfSGReport.get(rowIndex)[5].toString().trim();
            if (!oemGroupCol.isEmpty()){
                oemGroup = oemGroupCol;
                mainSR = mainSRCol;
            }
            if (!pnCol.isEmpty()){
                pn = pnCol;
                if (revQtyCol.isEmpty()){
                    revQty = 0;
                }
                else {
                    revQty = Float.parseFloat(revQtyCol);
                }
                Object[] cols = new Object[4];
                cols[0] = oemGroup;
                cols[1] = mainSR;
                cols[2] = pn;
                cols[3] = revQty;
                table.add(cols);
            }
        }
        String output = "C:\\CucumberFramework\\Downloads\\Output.xlsx";
        Object[] headerCols = new Object[4];
        headerCols[0] = "OEM GROUP";
        headerCols[1] = "MAIN SALES REP";
        headerCols[2] = "PART NUMBER";
        headerCols[3] = "REV QTY";
        FileReaderManager.getInstance().getExcelReader().getOutputFromData(output, headerCols, table);
    }

}
