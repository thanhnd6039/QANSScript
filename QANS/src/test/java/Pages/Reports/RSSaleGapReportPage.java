package Pages.Reports;

import Helpers.Manager.FileReaderManager;

import java.util.ArrayList;
import java.util.List;

public class RSSaleGapReportPage {
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
    public void getDataFromSGReport(){
        List<Object[]> dataOfSGReport = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(SGReportFilePath, 0, 5, 2);
        for (int rowIndex = 0; rowIndex < dataOfSGReport.size(); rowIndex++){
            String oemGroupCol = dataOfSGReport.get(rowIndex)[1].toString().trim();
            String mainSalesRepCol = dataOfSGReport.get(rowIndex)[2].toString().trim();
            String pnCol = dataOfSGReport.get(rowIndex)[3].toString().trim();
        }
    }
}
