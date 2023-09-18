import Helpers.DataProvider.ExcelReader;
import Pages.Reports.RSSaleGapReportPage;

public class Test {
    public static void main(String[] args) {
        String filePath = "C:\\CucumberFramework\\Downloads\\Margin Reporting By OEM Part.xlsx";
//        RSTrackedOppDashboardPage rsTrackedOppDashboardPage = new RSTrackedOppDashboardPage();
//        rsTrackedOppDashboardPage.getDataFromTrackedOppDashboard();
//        RSMarginReportPage rsMarginReportPage = new RSMarginReportPage();
//        rsMarginReportPage.getDataFromMarginReport();
        RSSaleGapReportPage rsSaleGapReportPage = new RSSaleGapReportPage();
        rsSaleGapReportPage.getOEMGroupAndMainSalesRepFromSGReport();

        System.out.println("Done");
    }
}
