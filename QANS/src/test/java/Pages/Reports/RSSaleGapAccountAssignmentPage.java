package Pages.Reports;

import Helpers.Manager.FileReaderManager;
import org.apache.commons.collections4.list.SetUniqueList;

import java.util.ArrayList;
import java.util.List;

public class RSSaleGapAccountAssignmentPage {
    private String allOEMWithOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\AllOEMwithOEMGroup.xlsx";
    private String vtOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\VTOEMGroup.xlsx";
    private String SFFilePath = "C:\\CucumberFramework\\Downloads\\ApprovedSalesForecast.xlsx";
    private String approvedBudgetFilePath = "C:\\CucumberFramework\\Downloads\\ApprovedBudget.xlsx";
    private String revCostDumpFilePath = "C:\\CucumberFramework\\Downloads\\RevenueCostDump.xlsx";
    public RSSaleGapAccountAssignmentPage(){

    }
    public void getAllOEMGroupByMainSaleRep(){
        List<Object[]> allOEMGroupByMainSaleRepTable = new ArrayList<>();
        List<Object[]> VTOEMGroupTable = new ArrayList<>();
        VTOEMGroupTable = getOEMGroupByMainSaleRepFromVTOEMGroupOnNS();
        for (int rowIndex = 0; rowIndex < VTOEMGroupTable.size(); rowIndex++){
            String oemGroup = VTOEMGroupTable.get(rowIndex)[0].toString().trim();
            String saleRep = VTOEMGroupTable.get(rowIndex)[1].toString().trim();
            Object[] cols = new Object[2];
            cols[0] = oemGroup;
            cols[1] = saleRep;
            allOEMGroupByMainSaleRepTable.add(cols);
        }
        List<Object[]> SFTable = new ArrayList<>();
        SFTable = getOEMGroupByMainSaleRepFromSFOnNS();
        boolean isOEMGoupInAllOEMGroupByMainSaleRepTable = false;
        for (int rowIndex = 0; rowIndex < SFTable.size(); rowIndex++){
            isOEMGoupInAllOEMGroupByMainSaleRepTable = false;
            String oemGroup = SFTable.get(rowIndex)[0].toString().trim();
            String saleRep = SFTable.get(rowIndex)[1].toString().trim();
            Object[] cols = new Object[2];
            cols[0] = oemGroup;
            cols[1] = saleRep;
            for (int indexTemp = 0; indexTemp < allOEMGroupByMainSaleRepTable.size(); indexTemp++){
                String oemGroupFromAllOEMGroupByMainSaleRepTable = allOEMGroupByMainSaleRepTable.get(indexTemp)[0].toString().trim();
                if (oemGroup.equalsIgnoreCase(oemGroupFromAllOEMGroupByMainSaleRepTable)){
                    isOEMGoupInAllOEMGroupByMainSaleRepTable = true;
                    break;
                }
            }
            if (isOEMGoupInAllOEMGroupByMainSaleRepTable == false){
                allOEMGroupByMainSaleRepTable.add(cols);
            }
        }
        List<Object[]> budgetTable = new ArrayList<>();
        budgetTable = getOEMGroupByMainSaleRepFromBudgetOnNS();
        for (int rowIndex = 0; rowIndex < budgetTable.size(); rowIndex++){
            isOEMGoupInAllOEMGroupByMainSaleRepTable = false;
            String oemGroup = budgetTable.get(rowIndex)[0].toString().trim();
            String saleRep = budgetTable.get(rowIndex)[1].toString().trim();
            Object[] cols = new Object[2];
            cols[0] = oemGroup;
            cols[1] = saleRep;
            for (int indexTemp = 0; indexTemp < allOEMGroupByMainSaleRepTable.size(); indexTemp++){
                String oemGroupFromAllOEMGroupByMainSaleRepTable = allOEMGroupByMainSaleRepTable.get(indexTemp)[0].toString().trim();
                if (oemGroup.equalsIgnoreCase(oemGroupFromAllOEMGroupByMainSaleRepTable)){
                    isOEMGoupInAllOEMGroupByMainSaleRepTable = true;
                    break;
                }
            }
            if (isOEMGoupInAllOEMGroupByMainSaleRepTable == false){
                allOEMGroupByMainSaleRepTable.add(cols);
            }
        }
        List<Object[]> revCostDumpTable = new ArrayList<>();
        revCostDumpTable = getOEMGroupByMainSaleRepFromRevCostDumpOnNS();
        for (int rowIndex = 0; rowIndex < revCostDumpTable.size(); rowIndex++){
            isOEMGoupInAllOEMGroupByMainSaleRepTable = false;
            String oemGroup = revCostDumpTable.get(rowIndex)[0].toString().trim();
            String saleRep = revCostDumpTable.get(rowIndex)[1].toString().trim();
            Object[] cols = new Object[2];
            cols[0] = oemGroup;
            cols[1] = saleRep;
            for (int indexTemp = 0; indexTemp < allOEMGroupByMainSaleRepTable.size(); indexTemp++){
                String oemGroupFromAllOEMGroupByMainSaleRepTable = allOEMGroupByMainSaleRepTable.get(indexTemp)[0].toString().trim();
                if (oemGroup.equalsIgnoreCase(oemGroupFromAllOEMGroupByMainSaleRepTable)){
                    isOEMGoupInAllOEMGroupByMainSaleRepTable = true;
                    break;
                }
            }
            if (isOEMGoupInAllOEMGroupByMainSaleRepTable == false){
                allOEMGroupByMainSaleRepTable.add(cols);
            }
        }
        List<Object[]> allOEMGroupByMainSaleRepIncludedDataTable = new ArrayList<>();
        List<Object[]> dataOfSF = new ArrayList<>();
        List<Object[]> dataOfRevCostDump = new ArrayList<>();
        dataOfRevCostDump = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(revCostDumpFilePath, 0, 1, 0);
        dataOfSF = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(SFFilePath,0, 1,0);
        for (int rowIndex = 0; rowIndex < allOEMGroupByMainSaleRepTable.size(); rowIndex++){
            String oemGroup = allOEMGroupByMainSaleRepTable.get(rowIndex)[0].toString().trim();
            String mainSaleRep = allOEMGroupByMainSaleRepTable.get(rowIndex)[1].toString().trim();
            float revVal = 0;
            float backLogVal = 0;
            float backLogFCVal = 0;
            float customerFCVal = 0;
            float revTotal = 0;
            for (int rowIndexFromRevCostDump = 0; rowIndexFromRevCostDump < dataOfRevCostDump.size(); rowIndexFromRevCostDump++){
                String oemGroupCol = dataOfRevCostDump.get(rowIndexFromRevCostDump)[1].toString().trim();
                float revCol = (float)dataOfRevCostDump.get(rowIndexFromRevCostDump)[28];
                float backLogCol = (float)dataOfRevCostDump.get(rowIndexFromRevCostDump)[34];
                float backLogFCCol = (float)dataOfRevCostDump.get(rowIndexFromRevCostDump)[37];
                float customerFCCol = (float)dataOfRevCostDump.get(rowIndexFromRevCostDump)[40];
                if (oemGroupCol.contains(":")){
                    String[] strArr = oemGroupCol.split(":", 2);
                    oemGroupCol = strArr[1].toString().trim();
                }
                if (oemGroup.equalsIgnoreCase(oemGroupCol)){
                    revTotal += revCol + backLogCol + backLogFCCol + customerFCCol;
                }
            }
            float SFVal = 0;
            for (int rowIndexFromSF = 0; rowIndexFromSF < dataOfSF.size(); rowIndexFromSF++){
                String oemGroupCol = dataOfSF.get(rowIndexFromSF)[1].toString().trim();
                float SFCol = (float) dataOfSF.get(rowIndexFromSF)[7];
                if (oemGroupCol.contains(":")){
                    String[] strArr = oemGroupCol.split(":", 2);
                    oemGroupCol = strArr[1].toString().trim();
                }
                if (oemGroup.equalsIgnoreCase(oemGroupCol)){
                    SFVal += SFCol;
                }
            }

            Object[] cols = new Object[8];
            cols[0] = oemGroup;
            cols[1] = mainSaleRep;
            cols[2] = revTotal;
        }

        String output = "C:\\CucumberFramework\\Downloads\\Output.xlsx";
        Object[] headerCols = new Object[2];
        headerCols[0] = "OEM Group";
        headerCols[1] = "Sale Rep";
        FileReaderManager.getInstance().getExcelReader().getOutputFromData(output, headerCols, allOEMGroupByMainSaleRepTable);
    }
    public List<Object[]> getOEMGroupByMainSaleRepFromVTOEMGroupOnNS(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfVTOEMGroup = new ArrayList<>();
        dataOfVTOEMGroup = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(vtOEMGroupFilePath, 0, 1, 0);
        for (int rowIndex = 0; rowIndex < dataOfVTOEMGroup.size(); rowIndex++){
            String oemGroup = dataOfVTOEMGroup.get(rowIndex)[1].toString().trim();
            String saleRep = dataOfVTOEMGroup.get(rowIndex)[4].toString().trim();
            Object[] cols = new Object[2];
            cols[0] = oemGroup;
            cols[1] = saleRep;
            table.add(cols);
        }
        return table;
    }
    public List<Object[]> getOEMGroupByMainSaleRepFromSFOnNS(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfSF = new ArrayList<>();
        dataOfSF = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(SFFilePath,0, 1,0);
        List<String> listOfOEMGroupFromSF = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < dataOfSF.size(); rowIndex++){
            listOfOEMGroupFromSF.add(dataOfSF.get(rowIndex)[1].toString().trim());
        }
        listOfOEMGroupFromSF = SetUniqueList.setUniqueList(listOfOEMGroupFromSF);
        listOfOEMGroupFromSF = getOEMGroupNeedToTakeSaleRep(listOfOEMGroupFromSF);
        for (int rowIndex = 0; rowIndex < listOfOEMGroupFromSF.size(); rowIndex++){
            String oemGroup = listOfOEMGroupFromSF.get(rowIndex).toString().trim();
            String saleRep = getSaleRepFromOEMGroup(oemGroup);
            Object[] cols = new Object[2];
            cols[0] = oemGroup;
            cols[1] = saleRep;
            table.add(cols);
        }
        return table;
    }
    public List<Object[]> getOEMGroupByMainSaleRepFromBudgetOnNS(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfBudget = new ArrayList<>();
        dataOfBudget = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(approvedBudgetFilePath,0, 1,0);
        List<String> listOfOEMGroupFromBudget = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < dataOfBudget.size(); rowIndex++){
            listOfOEMGroupFromBudget.add(dataOfBudget.get(rowIndex)[1].toString().trim());
        }
        listOfOEMGroupFromBudget = SetUniqueList.setUniqueList(listOfOEMGroupFromBudget);
        listOfOEMGroupFromBudget = getOEMGroupNeedToTakeSaleRep(listOfOEMGroupFromBudget);
        for (int rowIndex = 0; rowIndex < listOfOEMGroupFromBudget.size(); rowIndex++){
            String oemGroup = listOfOEMGroupFromBudget.get(rowIndex).toString().trim();
            String saleRep = getSaleRepFromOEMGroup(oemGroup);
            Object[] cols = new Object[2];
            cols[0] = oemGroup;
            cols[1] = saleRep;
            table.add(cols);
        }
        return table;
    }
    public List<Object[]> getOEMGroupByMainSaleRepFromRevCostDumpOnNS(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfRevCostDump = new ArrayList<>();
        dataOfRevCostDump = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(revCostDumpFilePath, 0, 1, 0);
        List<String> listOfOEMGroupFromRevCostDump = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < dataOfRevCostDump.size(); rowIndex++){
            listOfOEMGroupFromRevCostDump.add(dataOfRevCostDump.get(rowIndex)[1].toString().trim());
        }
        listOfOEMGroupFromRevCostDump = SetUniqueList.setUniqueList(listOfOEMGroupFromRevCostDump);
        listOfOEMGroupFromRevCostDump = getOEMGroupNeedToTakeSaleRep(listOfOEMGroupFromRevCostDump);
        for (int rowIndex = 0; rowIndex < listOfOEMGroupFromRevCostDump.size(); rowIndex++){
            String oemGroup = listOfOEMGroupFromRevCostDump.get(rowIndex).toString().trim();
            String saleRep = getSaleRepFromOEMGroup(oemGroup);
            Object[] cols = new Object[2];
            cols[0] = oemGroup;
            cols[1] = saleRep;
            table.add(cols);
        }
        return table;
    }
    public String getSaleRepFromOEMGroup(String oemGroup){
        String saleRep = "";
        List<Object[]> dataOfAllOEMWithOEMGroup = new ArrayList<>();
        dataOfAllOEMWithOEMGroup = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(allOEMWithOEMGroupFilePath, 0, 1, 0);
        for (int rowIndex = 0; rowIndex < dataOfAllOEMWithOEMGroup.size(); rowIndex++){
            String cusName =  dataOfAllOEMWithOEMGroup.get(rowIndex)[2].toString().trim();
            String saleRepFromAllOEMWithOEMGroup = dataOfAllOEMWithOEMGroup.get(rowIndex)[4].toString().trim();
            if (cusName.contains(":")){
                String[] strArr = cusName.split(":", 2);
                cusName = strArr[1].toString().trim();
            }
            if (oemGroup.equalsIgnoreCase(cusName)){
                saleRep = saleRepFromAllOEMWithOEMGroup;
                break;
            }
        }
        return saleRep;
    }
    public List<String> getOEMGroupNeedToTakeSaleRep(List<String> listOfOEMGroup){
        List<String> listOfOEMGroupNeedToTakeSaleRep = new ArrayList<>();
        List<Object[]> dataOfVTOEMGroup = new ArrayList<>();
        dataOfVTOEMGroup = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(vtOEMGroupFilePath, 0, 1, 0);
        for (int SFRowIndex = 0; SFRowIndex < listOfOEMGroup.size(); SFRowIndex++){
            String oemGroup = listOfOEMGroup.get(SFRowIndex).toString().trim();
            if (oemGroup.contains(":")){
                String[] strArr = oemGroup.split(":", 2);
                oemGroup = strArr[1].toString().trim();
            }
            boolean isOEMGroupSFInVTOEMGroup = false;
            for (int VTOEMGroupRowIndex = 0; VTOEMGroupRowIndex < dataOfVTOEMGroup.size(); VTOEMGroupRowIndex++){
                String oemGroupFromVTOEMGroup = dataOfVTOEMGroup.get(VTOEMGroupRowIndex)[1].toString().trim();
                if (oemGroup.equalsIgnoreCase(oemGroupFromVTOEMGroup)){
                    isOEMGroupSFInVTOEMGroup = true;
                    break;
                }
            }
            if (isOEMGroupSFInVTOEMGroup == false){
                listOfOEMGroupNeedToTakeSaleRep.add(oemGroup);
            }
        }
        return listOfOEMGroupNeedToTakeSaleRep;
    }


}
