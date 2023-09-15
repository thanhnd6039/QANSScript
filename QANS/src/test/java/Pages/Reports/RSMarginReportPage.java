package Pages.Reports;

import Helpers.Manager.FileReaderManager;

import java.util.ArrayList;
import java.util.List;

public class RSMarginReportPage {
    private String marginReportFilePath = "C:\\CucumberFramework\\Downloads\\Margin Reporting By OEM Part.xlsx";
    public RSMarginReportPage(){

    }
    public void getDataFromMarginReport(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfMarginReport = new ArrayList<>();
        dataOfMarginReport = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(marginReportFilePath, 0, 6, 5);
        String oemGroup = "";
        String pn = "";
        float qty = 0;
        for (int rowIndex = 0; rowIndex < dataOfMarginReport.size(); rowIndex++){
            String oemGroupCol = dataOfMarginReport.get(rowIndex)[0].toString().trim();
            String pnCol = dataOfMarginReport.get(rowIndex)[1].toString().trim();
            String qtyCol = dataOfMarginReport.get(rowIndex)[4].toString();
            if (!oemGroupCol.isEmpty()){
                oemGroup = oemGroupCol;
            }
            if (!pnCol.equalsIgnoreCase("Total")){
                pn = pnCol;
                if (qtyCol.isEmpty()){
                    qtyCol = "0";
                }
                qty = Float.parseFloat(qtyCol);
                Object[] cols = new Object[3];
                cols[0] = oemGroup;
                cols[1] = pn;
                cols[2] = qty;
                table.add(cols);
            }
        }
        String output = "C:\\CucumberFramework\\Downloads\\Output.xlsx";
        Object[] headerCols = new Object[3];
        headerCols[0] = "OEM Group";
        headerCols[1] = "Part Number";
        headerCols[2] = "QTY";
        FileReaderManager.getInstance().getExcelReader().getOutputFromData(output, headerCols, table);
    }
}
