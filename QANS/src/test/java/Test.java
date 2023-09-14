import Pages.Reports.RSSaleGapAccountAssignmentPage;
import Pages.Reports.RSTrackedOppDashboardPage;


public class Test {
    public static void main(String[] args) {
        String filePath = "C:\\CucumberFramework\\Downloads\\Output.xlsx";
        RSTrackedOppDashboardPage rsTrackedOppDashboardPage = new RSTrackedOppDashboardPage();
        rsTrackedOppDashboardPage.getDataFromTrackedOppDashboard();
        System.out.println("Done");
    }
}
