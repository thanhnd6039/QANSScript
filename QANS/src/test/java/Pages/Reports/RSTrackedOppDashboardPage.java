package Pages.Reports;

import Helpers.Manager.FileReaderManager;

import java.util.ArrayList;
import java.util.List;

public class RSTrackedOppDashboardPage {
    private String trackedOppDashboardFilePath = "C:\\CucumberFramework\\Downloads\\Tracked Opp Dashboard.xlsx";
    private String ssMasterOppFilePath = "C:\\CucumberFramework\\Downloads\\MasterOpportunityReportTesting.xlsx";
    public RSTrackedOppDashboardPage(){

    }

    public void getDataFromTrackedOppDashboard(){
        List<Object[]> table = new ArrayList<>();
        List<Object[]> dataOfTrackedOppDashboard = new ArrayList<>();
        dataOfTrackedOppDashboard = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(trackedOppDashboardFilePath, 0, 3, 2);
        for (int rowIndex = 1; rowIndex < 9; rowIndex++){
            String item = dataOfTrackedOppDashboard.get(rowIndex)[0].toString().trim();
            String currentWeek = dataOfTrackedOppDashboard.get(rowIndex)[1].toString();
            System.out.println(String.format("Current Week: %s", currentWeek));
            int countOppsOfCurrentWeek = Integer.parseInt(currentWeek);
            System.out.println(String.format("After Current Week: %d", countOppsOfCurrentWeek));
        }
    }
}
