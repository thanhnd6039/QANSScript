package Pages.Reports;

import Helpers.Manager.FileReaderManager;

import java.util.ArrayList;
import java.util.List;

public class RSMarginReportPage {
    private String marginReportFilePath = "C:\\CucumberFramework\\Downloads\\Margin Reporting By OEM Part.xlsx";
    public RSMarginReportPage(){

    }
    public void getDataFromMarginReport(){
        List<Object[]> dataOfMarginReport = new ArrayList<>();
        dataOfMarginReport = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(marginReportFilePath, 0, 7, 6);
        for (int rowIndex = 0; rowIndex < dataOfMarginReport.size(); rowIndex++){
            String oemGroupCol = dataOfMarginReport.get(rowIndex)[0].toString().trim();
            String pnCol = dataOfMarginReport.get(rowIndex)[1].toString().trim();
            int qtyVal =  Integer.parseInt(dataOfMarginReport.get(rowIndex)[4].toString());

        }
    }
}
