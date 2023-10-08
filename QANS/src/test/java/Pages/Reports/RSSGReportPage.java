package Pages.Reports;

import Helpers.KeywordWebUI;
import Helpers.Manager.FileReaderManager;
import Pages.CommonPage;
import org.apache.commons.collections4.list.SetUniqueList;
import org.junit.Assert;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import java.util.ArrayList;
import java.util.List;

public class RSSGReportPage extends CommonPage {
    private WebDriver driver;
    private String SGReportFilePath = "C:\\CucumberFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx";
    private String ssRevCostDumpFilePath = "C:\\CucumberFramework\\Downloads\\REV Cost Dump.xlsx";
    private String targetOutput = "C:\\CucumberFramework\\Downloads\\Target.xlsx";
    private String sourceOutput = "C:\\CucumberFramework\\Downloads\\Source.xlsx";
    @FindBy(xpath = "//span[contains(text(),'Sales Gap Report')]")
    private WebElement txtTitle;
    public RSSGReportPage(WebDriver driver){
        super(driver);
        this.driver = driver;
        PageFactory.initElements(driver, this);
    }
    public void shouldSeeTitleOfReport(String expectedTitle){
        Assert.assertTrue(waitForElementVisibility(txtTitle));
        String actualTitle = getTextFromElement(txtTitle);
        Assert.assertTrue(actualTitle.contains(expectedTitle));
    }
    public void getDataFromSGReport(String quarter, String year){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfSGReport = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(SGReportFilePath, 0, 5, 2);

        String oemGroup = "";
        String pn = "";
        float rQty = 0;
        float rAmount = 0;
        float bQty = 0;
        float bAmount = 0;
        float bfQty = 0;
        float bfAmount = 0;
        float cfQty = 0;
        float cfAmount = 0;
        boolean hasR = false;
        boolean hasB = false;
        boolean hasBF = false;
        boolean hasCF = false;

        String searchStrR = String.format("%s.Q%s R", year, quarter);
        String searchStrB = String.format("%s.Q%s B", year, quarter);
        String searchStrBF = String.format("%s.Q%s BF", year, quarter);
        String searchStrCF = String.format("%s.Q%s CF", year, quarter);

        int posOfR = FileReaderManager.getInstance().getExcelReader().getPosOfCol(SGReportFilePath, 0, 2, searchStrR);
        int posOfB = FileReaderManager.getInstance().getExcelReader().getPosOfCol(SGReportFilePath, 0, 2, searchStrB);
        int posOfBF = FileReaderManager.getInstance().getExcelReader().getPosOfCol(SGReportFilePath, 0, 2, searchStrBF);
        int posOfCF = FileReaderManager.getInstance().getExcelReader().getPosOfCol(SGReportFilePath, 0, 2, searchStrCF);

        if (posOfR != -1){
            hasR = true;
        }
        if (posOfB != -1){
            hasB = true;
        }
        if (posOfBF != -1){
            hasBF = true;
        }
        if (posOfCF != -1){
            hasCF = true;
        }

        for (int rowIndex = 0; rowIndex < dataOfSGReport.size(); rowIndex++){
            String oemGroupCol = dataOfSGReport.get(rowIndex)[1].toString().trim();
            String pnCol = dataOfSGReport.get(rowIndex)[3].toString().trim();
            String rQtyCol = "";
            String rAmountCol = "";
            String bQtyCol = "";
            String bAmountCol = "";
            String bfQtyCol = "";
            String bfAmountCol = "";
            String cfQtyCol = "";
            String cfAmountCol = "";

            if (hasR == true){
                rQtyCol = dataOfSGReport.get(rowIndex)[posOfR].toString().trim();
                rAmountCol = dataOfSGReport.get(rowIndex)[posOfR+2].toString().trim();
            }

            if (hasB == true){
                bQtyCol = dataOfSGReport.get(rowIndex)[posOfB].toString().trim();
                bAmountCol = dataOfSGReport.get(rowIndex)[posOfB+2].toString().trim();
            }

            if (hasBF == true){
                bfQtyCol = dataOfSGReport.get(rowIndex)[posOfBF].toString().trim();
                bfAmountCol = dataOfSGReport.get(rowIndex)[posOfBF+2].toString().trim();
            }

            if (hasCF == true){
                cfQtyCol = dataOfSGReport.get(rowIndex)[posOfCF].toString().trim();
                cfAmountCol = dataOfSGReport.get(rowIndex)[posOfCF+2].toString().trim();
            }

            if (!oemGroupCol.isEmpty()){
                oemGroup = oemGroupCol;
            }

            if (!pnCol.isEmpty()){
                pn = pnCol;

                if (rQtyCol.isEmpty()){
                    rQty = 0;
                }
                else {
                    rQty = Float.parseFloat(rQtyCol);
                }

                if (rAmountCol.isEmpty()){
                    rAmount = 0;
                }
                else {
                    rAmount = Float.parseFloat(rAmountCol);
                }

                if (bQtyCol.isEmpty()){
                    bQty = 0;
                }
                else {
                    bQty = Float.parseFloat(bQtyCol);
                }

                if (bAmountCol.isEmpty()){
                    bAmount = 0;
                }
                else {
                    bAmount = Float.parseFloat(bAmountCol);
                }

                if (bfQtyCol.isEmpty()){
                    bfQty = 0;
                }
                else {
                    bfQty = Float.parseFloat(bfQtyCol);
                }

                if (bfAmountCol.isEmpty()){
                    bfAmount = 0;
                }
                else {
                    bfAmount = Float.parseFloat(bfAmountCol);
                }

                if (cfQtyCol.isEmpty()){
                    cfQty = 0;
                }
                else {
                    cfQty = Float.parseFloat(cfQtyCol);
                }

                if (cfAmountCol.isEmpty()){
                    cfAmount = 0;
                }
                else {
                    cfAmount = Float.parseFloat(cfAmountCol);
                }

                Object[] cols = new Object[10];
                cols[0] = oemGroup;
                cols[1] = pn;
                cols[2] = rQty;
                cols[3] = rAmount;
                cols[4] = bQty;
                cols[5] = bAmount;
                cols[6] = bfQty;
                cols[7] = bfAmount;
                cols[8] = cfQty;
                cols[9] = cfAmount;
                table.add(cols);
            }
        }

        Object[] headerCols = new Object[10];
        headerCols[0] = "OEM GROUP";
        headerCols[1] = "PART NUMBER";
        headerCols[2] = "REV QTY";
        headerCols[3] = "REV AMOUNT";
        headerCols[4] = "BL QTY";
        headerCols[5] = "BL AMOUNT";
        headerCols[6] = "BF QTY";
        headerCols[7] = "BF AMOUNT";
        headerCols[8] = "CF QTY";
        headerCols[9] = "CF AMOUNT";
        FileReaderManager.getInstance().getExcelReader().getOutputFromData(targetOutput, headerCols, table);
    }
    public void getDataFromSSRevCostDump(String quarter, String year){
        List<Object[]> tempTable = new ArrayList<>();
        List<Object[]> dataOfRevCostDump = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(ssRevCostDumpFilePath, 0, 1, 0);
        String strQuar = String.format("Q%s-%s", quarter, year);
        String oemGroupCol = "";
        String pnCol = "";
        String parentClassCol = "";
        String yearCol = "";
        String quarCol = "";
        String rQtyCol = "";
        String rAmountCol = "";
        String bQtyCol = "";
        String bAmountCol = "";
        String bfQtyCol = "";
        String bfAmountCol = "";
        String cfQtyCol = "";
        String cfAmountCol = "";

        for (int rowIndex = 0; rowIndex < dataOfRevCostDump.size(); rowIndex++){
            parentClassCol = dataOfRevCostDump.get(rowIndex)[8].toString().trim();
            yearCol = dataOfRevCostDump.get(rowIndex)[16].toString().trim();
            quarCol = dataOfRevCostDump.get(rowIndex)[17].toString().trim();
            if (parentClassCol.equalsIgnoreCase("COMPONENTS") ||
                    parentClassCol.equalsIgnoreCase("MEM") ||
                    parentClassCol.equalsIgnoreCase("STORAGE") ||
                    parentClassCol.equalsIgnoreCase("NI ITEMS")){
                if (yearCol.equalsIgnoreCase(year) && quarCol.equalsIgnoreCase(strQuar)){
                    oemGroupCol = dataOfRevCostDump.get(rowIndex)[1].toString().trim();
                    if (oemGroupCol.contains(":")){
                        String[] strArr = oemGroupCol.split(":", 2);
                        oemGroupCol = strArr[1].toString().trim();
                    }
                    pnCol = dataOfRevCostDump.get(rowIndex)[10].toString().trim();
                    rQtyCol = dataOfRevCostDump.get(rowIndex)[28].toString().trim();
                    rAmountCol = dataOfRevCostDump.get(rowIndex)[29].toString().trim();
                    bQtyCol = dataOfRevCostDump.get(rowIndex)[33].toString().trim();
                    bAmountCol = dataOfRevCostDump.get(rowIndex)[35].toString().trim();
                    bfQtyCol = dataOfRevCostDump.get(rowIndex)[36].toString().trim();
                    bfAmountCol = dataOfRevCostDump.get(rowIndex)[38].toString().trim();
                    cfQtyCol = dataOfRevCostDump.get(rowIndex)[39].toString().trim();
                    cfAmountCol = dataOfRevCostDump.get(rowIndex)[41].toString().trim();
                    Object[] cols = new Object[10];
                    cols[0] = oemGroupCol;
                    cols[1] = pnCol;
                    cols[2] = rQtyCol;
                    cols[3] = rAmountCol;
                    cols[4] = bQtyCol;
                    cols[5] = bAmountCol;
                    cols[6] = bfQtyCol;
                    cols[7] = bfAmountCol;
                    cols[8] = cfQtyCol;
                    cols[9] = cfAmountCol;
                    tempTable.add(cols);
                }
            }
        }
        List<Object[]> table = new ArrayList<>();
        List<String> listOfOEMGroup = new ArrayList<>();
        List<String> listOfPN = new ArrayList<>();
        listOfOEMGroup = getListOfOEMGroup(tempTable);
        listOfPN = getListOfPN(tempTable);

        for (int rowIndexFromListOfOEMGroup = 0; rowIndexFromListOfOEMGroup < listOfOEMGroup.size(); rowIndexFromListOfOEMGroup++){
            String oemGroupColFromListOfOEMGroup = listOfOEMGroup.get(rowIndexFromListOfOEMGroup).toString().trim();
            for (int rowIndexFromListOfPN = 0; rowIndexFromListOfPN < listOfPN.size(); rowIndexFromListOfPN++){
                String pnColFromListOfPN = listOfPN.get(rowIndexFromListOfPN).toString().trim();
                float rQty = 0;
                float rAmount = 0;
                float bQty = 0;
                float bAmount = 0;
                float bfQty = 0;
                float bfAmount = 0;
                float cfQty = 0;
                float cfAmount = 0;
                boolean hasData = false;

                for (int rowIndex = 0; rowIndex < tempTable.size(); rowIndex++){
                    oemGroupCol = tempTable.get(rowIndex)[0].toString().trim();
                    pnCol = tempTable.get(rowIndex)[1].toString().trim();
                    if (oemGroupColFromListOfOEMGroup.equalsIgnoreCase(oemGroupCol) && pnColFromListOfPN.equalsIgnoreCase(pnCol)){
                        rQtyCol = tempTable.get(rowIndex)[2].toString().trim();
                        rAmountCol = tempTable.get(rowIndex)[3].toString().trim();
                        bQtyCol = tempTable.get(rowIndex)[4].toString().trim();
                        bAmountCol = tempTable.get(rowIndex)[5].toString().trim();
                        bfQtyCol = tempTable.get(rowIndex)[6].toString().trim();
                        bfAmountCol = tempTable.get(rowIndex)[7].toString().trim();
                        cfQtyCol = tempTable.get(rowIndex)[8].toString().trim();
                        cfAmountCol = tempTable.get(rowIndex)[9].toString().trim();
                        float tempRQty = Float.parseFloat(rQtyCol);
                        float tempRAmount = Float.parseFloat(rAmountCol);
                        float tempBQty = Float.parseFloat(bQtyCol);
                        float tempBAmount = Float.parseFloat(bAmountCol);
                        float tempBFQty = Float.parseFloat(bfQtyCol);
                        float tempBFAmount = Float.parseFloat(bfAmountCol);
                        float tempCFQty = Float.parseFloat(cfQtyCol);
                        float tempCFAmount = Float.parseFloat(cfAmountCol);
                        rQty += tempRQty;
                        rAmount += tempRAmount;
                        bQty += tempBQty;
                        bAmount += tempBAmount;
                        bfQty += tempBFQty;
                        bfAmount += tempBFAmount;
                        cfQty += tempCFQty;
                        cfAmount += tempCFAmount;
                        hasData = true;
                    }
                }
                if (hasData == true){
                    Object[] cols = new Object[10];
                    cols[0] = oemGroupColFromListOfOEMGroup;
                    cols[1] = pnColFromListOfPN;
                    cols[2] = rQty;
                    cols[3] = rAmount;
                    cols[4] = bQty;
                    cols[5] = bAmount;
                    cols[6] = bfQty;
                    cols[7] = bfAmount;
                    cols[8] = cfQty;
                    cols[9] = cfAmount;
                    table.add(cols);
                }
            }
        }

        Object[] headerCols = new Object[10];
        headerCols[0] = "OEM GROUP";
        headerCols[1] = "PART NUMBER";
        headerCols[2] = "REV QTY";
        headerCols[3] = "REV AMOUNT";
        headerCols[4] = "BL QTY";
        headerCols[5] = "BL AMOUNT";
        headerCols[6] = "BF QTY";
        headerCols[7] = "BF AMOUNT";
        headerCols[8] = "CF QTY";
        headerCols[9] = "CF AMOUNT";
        FileReaderManager.getInstance().getExcelReader().getOutputFromData(sourceOutput, headerCols, table);
    }
    public List<String> getListOfOEMGroup(List<Object[]> data){
        List<String> listOfOEMGroup = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < data.size(); rowIndex++){
            String oemGroupCol = data.get(rowIndex)[0].toString().trim();
            listOfOEMGroup.add(oemGroupCol);
        }
        listOfOEMGroup = SetUniqueList.setUniqueList(listOfOEMGroup);
        return listOfOEMGroup;
    }
    public List<String> getListOfPN(List<Object[]> data){
        List<String> listOfPN = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < data.size(); rowIndex++){
            String pnCol = data.get(rowIndex)[1].toString().trim();
            listOfPN.add(pnCol);
        }
        listOfPN = SetUniqueList.setUniqueList(listOfPN);
        return listOfPN;
    }



}
