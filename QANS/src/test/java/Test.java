import Helpers.DataProvider.ExcelReader;
import Pages.CommonPage;
import Pages.Reports.RSSaleGapReportPage;
import org.openqa.selenium.WebDriver;

public class Test {
    public static void main(String[] args) {
        String filePath = "C:\\CucumberFramework\\Downloads\\Margin Reporting By OEM Part.xlsx";
//        RSTrackedOppDashboardPage rsTrackedOppDashboardPage = new RSTrackedOppDashboardPage();
//        rsTrackedOppDashboardPage.getDataFromTrackedOppDashboard();
//        RSMarginReportPage rsMarginReportPage = new RSMarginReportPage();
//        rsMarginReportPage.getDataFromMarginReport();
        RSSaleGapReportPage rsSaleGapReportPage = new RSSaleGapReportPage();
        rsSaleGapReportPage.getDataFromSGReport(2023, 2023, 3, 3);


        System.out.println("Done");
    }
}
