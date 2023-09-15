import Pages.Reports.RSMarginReportPage;
import Pages.Reports.RSSaleGapAccountAssignmentPage;
import Pages.Reports.RSSaleGapReportPage;
import Pages.Reports.RSTrackedOppDashboardPage;


public class Test {
    public static void main(String[] args) {
        String filePath = "C:\\CucumberFramework\\Downloads\\Output.xlsx";
//        RSTrackedOppDashboardPage rsTrackedOppDashboardPage = new RSTrackedOppDashboardPage();
//        rsTrackedOppDashboardPage.getDataFromTrackedOppDashboard();
        RSMarginReportPage rsMarginReportPage = new RSMarginReportPage();
        rsMarginReportPage.getDataFromMarginReport();
//        RSSaleGapReportPage rsSaleGapReportPage = new RSSaleGapReportPage();
//        rsSaleGapReportPage.getOEMGroupAndMainSalesRepFromSGReport();
        System.out.println("Done");
    }
}
