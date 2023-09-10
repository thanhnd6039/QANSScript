package Pages.Reports;

import Helpers.Manager.FileReaderManager;
import Pages.CommonPage;
import org.apache.commons.collections4.list.SetUniqueList;
import org.junit.Assert;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import java.util.*;

public class RSGapToSFPNDetailsPage extends CommonPage {
    private WebDriver driver;
    private String vtOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\VTOEMGroup.xlsx";
    @FindBy(xpath = "//*/div[contains(text(),'GAP to SF - Part Number Detail')]")
    private WebElement txtTitle;
    public RSGapToSFPNDetailsPage(WebDriver driver){
        super(driver);
        PageFactory.initElements(driver, this);
        this.driver = driver;
    }
    public void shouldSeeTitle(String title){
        Assert.assertTrue(waitForElementVisibility(txtTitle));
        String actualTitle = getTextFromElement(txtTitle);
        Assert.assertEquals(title, actualTitle);
    }
    public void getAllOEMGroupByMainSaleRep(){
        getOEMByMainSaleRepFromVTOEMGroupOnNS();
//        getOEMGroupByMainSaleRepFromVTOEMGroupFile();
//        getOEMGroupByMainSaleRepFromApprovedSaleFCFile();
    }
    public List<Object[]> getOEMByMainSaleRepFromVTOEMGroupOnNS(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfVTOEMGroup = new ArrayList<>();
        dataOfVTOEMGroup = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(vtOEMGroupFilePath, 0, 1, 0);
        for (int index = 0; index < dataOfVTOEMGroup.size(); index++){
            String oemGroup = dataOfVTOEMGroup.get(index)[1].toString().trim();
            String saleRep = dataOfVTOEMGroup.get(index)[4].toString().trim();
            Object[] cols = new Object[2];
            cols[0] = oemGroup;
            cols[1] = saleRep;
            table.add(cols);
        }
//        String output = "C:\\CucumberFramework\\Downloads\\Output.xlsx";
//        Object[] headerCols = new Object[2];
//        headerCols[0] = "OEM Group";
//        headerCols[1] = "Sale Rep";
//        FileReaderManager.getInstance().getExcelReader().getOutputFromData(output, headerCols, table);
        return table;
    }


//    public List<Object[]> getOEMGroupByMainSaleRepFromVTOEMGroupFile(){
//        List<Object[]> table = new ArrayList<>();
//
////        String allOEMWithOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\AllOEMwithOEMGroup.xlsx";
//        List<Object[]> dataOfVTOEMGroupFile = new ArrayList<>();
////        List<Object[]> dataOfAllOEMWithOEMGroupFile = new ArrayList<>();
////        List<Object[]> dataOfOEMGroupByMainSaleRep = new ArrayList<>();
//
//        dataOfVTOEMGroupFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(vtOEMGroupFilePath, 0, 1, 0);
//        dataOfAllOEMWithOEMGroupFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(allOEMWithOEMGroupFilePath, 0, 1, 0);
//
//        for (int dataVTOEMGroupIndex = 0; dataVTOEMGroupIndex < dataOfVTOEMGroupFile.size(); dataVTOEMGroupIndex++){
//            String oemGroupFromVTOEMGroupFile = dataOfVTOEMGroupFile.get(dataVTOEMGroupIndex)[1].toString().trim();
//            String saleRepFromVTOEMGroupFile = dataOfVTOEMGroupFile.get(dataVTOEMGroupIndex)[4].toString().trim();
//            Object[] colsArrOfOEMGroupByMainSaleRep = new Object[2];
//            colsArrOfOEMGroupByMainSaleRep[0] = oemGroupFromVTOEMGroupFile;
//            if (saleRepFromVTOEMGroupFile.isEmpty()){
//                for (int dataOfAllOEMWithOEMGroupIndex = 0; dataOfAllOEMWithOEMGroupIndex < dataOfAllOEMWithOEMGroupFile.size(); dataOfAllOEMWithOEMGroupIndex++){
//                    String cusName =  dataOfAllOEMWithOEMGroupFile.get(dataOfAllOEMWithOEMGroupIndex)[2].toString().trim();
//                    String saleRepFromAllOEMWithOEMGroupFile = dataOfAllOEMWithOEMGroupFile.get(dataOfAllOEMWithOEMGroupIndex)[4].toString().trim();
//                    String oemGroupFromAllOEMWithOEMGroupFile = dataOfAllOEMWithOEMGroupFile.get(dataOfAllOEMWithOEMGroupIndex)[7].toString().trim();
//                    if (oemGroupFromVTOEMGroupFile.equalsIgnoreCase(oemGroupFromAllOEMWithOEMGroupFile)){
//                        colsArrOfOEMGroupByMainSaleRep[1] = saleRepFromAllOEMWithOEMGroupFile;
//                        break;
//                    }
//                    if (oemGroupFromVTOEMGroupFile.equalsIgnoreCase(cusName)){
//                        colsArrOfOEMGroupByMainSaleRep[1] = saleRepFromAllOEMWithOEMGroupFile;
//                        break;
//                    }
//                    colsArrOfOEMGroupByMainSaleRep[1] = "";
//                }
//            }
//            else {
//                colsArrOfOEMGroupByMainSaleRep[1] = saleRepFromVTOEMGroupFile;
//            }
//            dataOfOEMGroupByMainSaleRep.add(colsArrOfOEMGroupByMainSaleRep);
//        }
//
//        return table;
//    }
    public List<Object[]> getOEMGroupByMainSaleRepFromApprovedSaleFCFile(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfApprovedSaleFCFile = new ArrayList<>();
        List<Object[]> dataOfVTOEMGroupFile = new ArrayList<>();
        List<Object[]> dataOfAllOEMWithOEMGroupFile = new ArrayList<>();
        String approvedSaleFCFilePath = "C:\\CucumberFramework\\Downloads\\ApprovedSalesForecast.xlsx";
        String vtOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\VTOEMGroup.xlsx";
        String allOEMWithOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\AllOEMwithOEMGroup.xlsx";

        dataOfApprovedSaleFCFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(approvedSaleFCFilePath,0, 1,0);
        dataOfVTOEMGroupFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(vtOEMGroupFilePath, 0, 1, 0);
        dataOfAllOEMWithOEMGroupFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(allOEMWithOEMGroupFilePath, 0, 1, 0);

        List<String> listOfOEMGroupFromApprovedSaleFCFile =  new ArrayList<>();
        for (int approvedSaleFCIndex = 0; approvedSaleFCIndex < dataOfApprovedSaleFCFile.size(); approvedSaleFCIndex++){
            listOfOEMGroupFromApprovedSaleFCFile.add(dataOfApprovedSaleFCFile.get(approvedSaleFCIndex)[1].toString().trim());
        }
        List<String> listOfOEMGroupFromApprovedSaleFCFileWithoutDuplicated = SetUniqueList.setUniqueList(listOfOEMGroupFromApprovedSaleFCFile);
        List<String> listOfOEMGroupFromApprovedSaleFCFileNeedToTakeSaleRep = new ArrayList<>();

        for (int approvedSaleFCIndex = 0; approvedSaleFCIndex < listOfOEMGroupFromApprovedSaleFCFileWithoutDuplicated.size(); approvedSaleFCIndex++){
            String oemGroupFromApprovedSaleFCFile = listOfOEMGroupFromApprovedSaleFCFileWithoutDuplicated.get(approvedSaleFCIndex).toString().trim();
            boolean isOEMGroupFromApprovedSaleFCInVTOEMGroupFile = false;
            for (int vtOEMGroupIndex = 0; vtOEMGroupIndex < dataOfVTOEMGroupFile.size(); vtOEMGroupIndex++){
                String oemGroupFromVTOEMGroupFile = dataOfVTOEMGroupFile.get(vtOEMGroupIndex)[1].toString().trim();
                if (oemGroupFromApprovedSaleFCFile.contains(":")){
                    String[] strArr = oemGroupFromApprovedSaleFCFile.split(":", 2);
                    oemGroupFromApprovedSaleFCFile = strArr[1].toString().trim();
                }
                if (oemGroupFromApprovedSaleFCFile.equalsIgnoreCase(oemGroupFromVTOEMGroupFile)){
                    isOEMGroupFromApprovedSaleFCInVTOEMGroupFile = true;
                    break;
                }
            }
            if (isOEMGroupFromApprovedSaleFCInVTOEMGroupFile == false){
                listOfOEMGroupFromApprovedSaleFCFileNeedToTakeSaleRep.add(oemGroupFromApprovedSaleFCFile);
            }
        }

        for (int approvedSaleFCIndex = 0; approvedSaleFCIndex < listOfOEMGroupFromApprovedSaleFCFileNeedToTakeSaleRep.size(); approvedSaleFCIndex++){
            String oemGroupFromApprovedSaleFCFile = listOfOEMGroupFromApprovedSaleFCFileNeedToTakeSaleRep.get(approvedSaleFCIndex).toString().trim();
            for(int allOEMWithOEMGroupIndex = 0; allOEMWithOEMGroupIndex < dataOfAllOEMWithOEMGroupFile.size(); allOEMWithOEMGroupIndex++){
                String oemGroupFromAllOEMWithOEMGroupFile = dataOfAllOEMWithOEMGroupFile.get(allOEMWithOEMGroupIndex)[7].toString().trim();
                String saleRepFromAllOEMWithOEMGroupFile = dataOfAllOEMWithOEMGroupFile.get(allOEMWithOEMGroupIndex)[4].toString().trim();
                String cusName =  dataOfAllOEMWithOEMGroupFile.get(allOEMWithOEMGroupIndex)[2].toString().trim();
                Object[] cols = new Object[2];
                if (oemGroupFromAllOEMWithOEMGroupFile.contains(":")){
                    String[] strArr = oemGroupFromAllOEMWithOEMGroupFile.split(":", 2);
                    oemGroupFromAllOEMWithOEMGroupFile = strArr[1].toString().trim();
                }
                if (oemGroupFromApprovedSaleFCFile.equalsIgnoreCase(oemGroupFromAllOEMWithOEMGroupFile)){
                    cols[0] = oemGroupFromApprovedSaleFCFile;
                    cols[1] = saleRepFromAllOEMWithOEMGroupFile;
                    table.add(cols);
                    break;
                }
                if (cusName.contains(":")){
                    String[] strArr = cusName.split(":", 2);
                    cusName = strArr[1].toString().trim();
                }
                if (oemGroupFromApprovedSaleFCFile.equalsIgnoreCase(cusName)){
                    cols[0] = oemGroupFromApprovedSaleFCFile;
                    cols[1] = saleRepFromAllOEMWithOEMGroupFile;
                    table.add(cols);
                    break;
                }
            }
        }
//        String output = "C:\\CucumberFramework\\Downloads\\Result.xlsx";
//        FileReaderManager.getInstance().getExcelReader().writeDataToExcel(output, 0, 0, 0, "OEM");
//        FileReaderManager.getInstance().getExcelReader().writeDataToExcel(output, 0, 0, 1, "Sale");
//        for (int index = 0; index < table.size(); index++){
//            String oem = table.get(index)[0].toString().trim();
//            String sale = table.get(index)[1].toString().trim();
//            FileReaderManager.getInstance().getExcelReader().writeDataToExcel(output, 0, index+1, 0, oem);
//            FileReaderManager.getInstance().getExcelReader().writeDataToExcel(output, 0, index+1, 1, sale);
//        }
        return table;
    }
    public List<Object[]> getOEMGroupByMainSaleRepFromApprovedBudgetFile(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfVTOEMGroupFile = new ArrayList<>();
        List<Object[]> dataOfApprovedBudgetFile = new ArrayList<>();
        List<Object[]> dataOfAllOEMWithOEMGroupFile = new ArrayList<>();
        String vtOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\VTOEMGroup.xlsx";
        String approvedBudgetFilePath = "C:\\CucumberFramework\\Downloads\\ApprovedBudget.xlsx";
        String allOEMWithOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\AllOEMwithOEMGroup.xlsx";

        dataOfVTOEMGroupFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(vtOEMGroupFilePath, 0, 1, 0);
        dataOfApprovedBudgetFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(approvedBudgetFilePath,0, 1,0);
        dataOfAllOEMWithOEMGroupFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(allOEMWithOEMGroupFilePath, 0, 1, 0);
        List<String> listOfOEMGroupFromApprovedBudgetFile =  new ArrayList<>();
        for (int approvedBudgetIndex = 0; approvedBudgetIndex < dataOfApprovedBudgetFile.size(); approvedBudgetIndex++){
            listOfOEMGroupFromApprovedBudgetFile.add(dataOfApprovedBudgetFile.get(approvedBudgetIndex)[1].toString().trim());
        }
        List<String> listOfOEMGroupFromApprovedBudgetFileWithoutDuplicated = SetUniqueList.setUniqueList(listOfOEMGroupFromApprovedBudgetFile);
        List<String> listOfOEMGroupFromApprovedBudgetFileNeedToTakeSaleRep = new ArrayList<>();

        for (int approvedBudgetIndex = 0; approvedBudgetIndex < listOfOEMGroupFromApprovedBudgetFileWithoutDuplicated.size(); approvedBudgetIndex++){
            boolean isOEMGroupFromApprovedBudgetInVTOEMGroupFile = false;
            String oemGroupFromApprovedBudgetFile = listOfOEMGroupFromApprovedBudgetFileWithoutDuplicated.get(approvedBudgetIndex).toString().trim();
            for (int vtOEMGroupIndex = 0; vtOEMGroupIndex < dataOfVTOEMGroupFile.size(); vtOEMGroupIndex++){
                String oemGroupFromVTOEMGroupFile = dataOfVTOEMGroupFile.get(vtOEMGroupIndex)[1].toString().trim();
                if (oemGroupFromApprovedBudgetFile.contains(":")){
                    String[] strArr = oemGroupFromApprovedBudgetFile.split(":", 2);
                    oemGroupFromApprovedBudgetFile = strArr[1].toString().trim();
                }
                if (oemGroupFromApprovedBudgetFile.equalsIgnoreCase(oemGroupFromVTOEMGroupFile)){
                    isOEMGroupFromApprovedBudgetInVTOEMGroupFile = true;
                    break;
                }
            }
            if (isOEMGroupFromApprovedBudgetInVTOEMGroupFile == false){
                listOfOEMGroupFromApprovedBudgetFileNeedToTakeSaleRep.add(oemGroupFromApprovedBudgetFile);
            }
        }
        for (int approvedBudgetIndex = 0; approvedBudgetIndex < listOfOEMGroupFromApprovedBudgetFileNeedToTakeSaleRep.size(); approvedBudgetIndex++){
            String oemGroupFromApprovedBudgetFile = listOfOEMGroupFromApprovedBudgetFileNeedToTakeSaleRep.get(approvedBudgetIndex).toString().trim();
            for(int allOEMWithOEMGroupIndex = 0; allOEMWithOEMGroupIndex < dataOfAllOEMWithOEMGroupFile.size(); allOEMWithOEMGroupIndex++){
                String oemGroupFromAllOEMWithOEMGroupFile = dataOfAllOEMWithOEMGroupFile.get(allOEMWithOEMGroupIndex)[7].toString().trim();
                String saleRepFromAllOEMWithOEMGroupFile = dataOfAllOEMWithOEMGroupFile.get(allOEMWithOEMGroupIndex)[4].toString().trim();
                String cusName =  dataOfAllOEMWithOEMGroupFile.get(allOEMWithOEMGroupIndex)[2].toString().trim();
                Object[] cols = new Object[2];
                if (oemGroupFromAllOEMWithOEMGroupFile.contains(":")){
                    String[] strArr = oemGroupFromAllOEMWithOEMGroupFile.split(":", 2);
                    oemGroupFromAllOEMWithOEMGroupFile = strArr[1].toString().trim();
                }
                if (oemGroupFromApprovedBudgetFile.equalsIgnoreCase(oemGroupFromAllOEMWithOEMGroupFile)){
                    cols[0] = oemGroupFromApprovedBudgetFile;
                    cols[1] = saleRepFromAllOEMWithOEMGroupFile;
                    table.add(cols);
                    break;
                }
                if (cusName.contains(":")){
                    String[] strArr = cusName.split(":", 2);
                    cusName = strArr[1].toString().trim();
                }
                if (oemGroupFromApprovedBudgetFile.equalsIgnoreCase(cusName)){
                    cols[0] = oemGroupFromApprovedBudgetFile;
                    cols[1] = saleRepFromAllOEMWithOEMGroupFile;
                    table.add(cols);
                    break;
                }
            }
        }
        return table;
    }
    public List<Object[]> getOEMGroupByMainSaleRepFromRevenueCostDumpFile(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfVTOEMGroupFile = new ArrayList<>();
        List<Object[]> dataOfRevenueCostDumpFile = new ArrayList<>();
        List<Object[]> dataOfAllOEMWithOEMGroupFile = new ArrayList<>();
        String vtOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\VTOEMGroup.xlsx";
        String revenueCostDumpFilePath = "C:\\CucumberFramework\\Downloads\\RevenueCostDump.xlsx";
        String allOEMWithOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\AllOEMwithOEMGroup.xlsx";

        dataOfVTOEMGroupFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(vtOEMGroupFilePath, 0, 1, 0);
        dataOfRevenueCostDumpFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(revenueCostDumpFilePath, 0, 1, 0);
        dataOfAllOEMWithOEMGroupFile = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(allOEMWithOEMGroupFilePath, 0, 1, 0);
        List<String> listOfOEMGroupFromRevenueCostDumpFile =  new ArrayList<>();
        for (int revenueCostDumpIndex = 0; revenueCostDumpIndex < dataOfRevenueCostDumpFile.size(); revenueCostDumpIndex++){
            listOfOEMGroupFromRevenueCostDumpFile.add(dataOfRevenueCostDumpFile.get(revenueCostDumpIndex)[1].toString().trim());
        }
        List<String> listOfOEMGroupFromRevenueCostDumpFileWithoutDuplicated = SetUniqueList.setUniqueList(listOfOEMGroupFromRevenueCostDumpFile);
        List<String> listOfOEMGroupFromRevenueCostDumpFileNeedToTakeSaleRep = new ArrayList<>();
        for (int revenueCostDumpIndex = 0; revenueCostDumpIndex < listOfOEMGroupFromRevenueCostDumpFileWithoutDuplicated.size(); revenueCostDumpIndex++){
            boolean isOEMGroupFromRevenueCostDumpInVTOEMGroupFile = false;
            String oemGroupFromRevenueCostDumpFile = listOfOEMGroupFromRevenueCostDumpFileWithoutDuplicated.get(revenueCostDumpIndex).toString().trim();
            for (int vtOEMGroupIndex = 0; vtOEMGroupIndex < dataOfVTOEMGroupFile.size(); vtOEMGroupIndex++){
                String oemGroupFromVTOEMGroupFile = dataOfVTOEMGroupFile.get(vtOEMGroupIndex)[1].toString().trim();
                if (oemGroupFromRevenueCostDumpFile.contains(":")){
                    String[] strArr = oemGroupFromRevenueCostDumpFile.split(":", 2);
                    oemGroupFromRevenueCostDumpFile = strArr[1].toString().trim();
                }
                if (oemGroupFromRevenueCostDumpFile.equalsIgnoreCase(oemGroupFromVTOEMGroupFile)){
                    isOEMGroupFromRevenueCostDumpInVTOEMGroupFile = true;
                    break;
                }
            }
            if (isOEMGroupFromRevenueCostDumpInVTOEMGroupFile == false){
                listOfOEMGroupFromRevenueCostDumpFileNeedToTakeSaleRep.add(oemGroupFromRevenueCostDumpFile);
            }
        }
        for (int revenueCostDumpIndex = 0; revenueCostDumpIndex < listOfOEMGroupFromRevenueCostDumpFileNeedToTakeSaleRep.size(); revenueCostDumpIndex++){
            String oemGroupFromRevenueCostDumpFile = listOfOEMGroupFromRevenueCostDumpFileNeedToTakeSaleRep.get(revenueCostDumpIndex).toString().trim();
            for(int allOEMWithOEMGroupIndex = 0; allOEMWithOEMGroupIndex < dataOfAllOEMWithOEMGroupFile.size(); allOEMWithOEMGroupIndex++){
                String oemGroupFromAllOEMWithOEMGroupFile = dataOfAllOEMWithOEMGroupFile.get(allOEMWithOEMGroupIndex)[7].toString().trim();
                String saleRepFromAllOEMWithOEMGroupFile = dataOfAllOEMWithOEMGroupFile.get(allOEMWithOEMGroupIndex)[4].toString().trim();
                String cusName =  dataOfAllOEMWithOEMGroupFile.get(allOEMWithOEMGroupIndex)[2].toString().trim();
                Object[] cols = new Object[2];
                if (oemGroupFromAllOEMWithOEMGroupFile.contains(":")){
                    String[] strArr = oemGroupFromAllOEMWithOEMGroupFile.split(":", 2);
                    oemGroupFromAllOEMWithOEMGroupFile = strArr[1].toString().trim();
                }
                if (oemGroupFromRevenueCostDumpFile.equalsIgnoreCase(oemGroupFromAllOEMWithOEMGroupFile)){
                    cols[0] = oemGroupFromRevenueCostDumpFile;
                    cols[1] = saleRepFromAllOEMWithOEMGroupFile;
                    table.add(cols);
                    break;
                }
                if (cusName.contains(":")){
                    String[] strArr = cusName.split(":", 2);
                    cusName = strArr[1].toString().trim();
                }
                if (oemGroupFromRevenueCostDumpFile.equalsIgnoreCase(cusName)){
                    cols[0] = oemGroupFromRevenueCostDumpFile;
                    cols[1] = saleRepFromAllOEMWithOEMGroupFile;
                    table.add(cols);
                    break;
                }
            }
        }

        return table;
    }


    public void getSourceDataForGapToSFPNDetailReport(){

    }




}
