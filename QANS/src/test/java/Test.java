import Helpers.DataProvider.ExcelReader;
import Pages.CommonPage;
import Pages.Reports.RSSaleGapReportPage;
import org.openqa.selenium.WebDriver;

public class Test {
    public static void main(String[] args) {
        String filePath = "C:\\CucumberFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx";
//        ExcelReader excelReader = new ExcelReader();
//        int pos = excelReader.getPosOfCol(filePath, 0, 2, "2023.Q3 B");

//        RSTrackedOppDashboardPage rsTrackedOppDashboardPage = new RSTrackedOppDashboardPage();
//        rsTrackedOppDashboardPage.getDataFromTrackedOppDashboard();
//        RSMarginReportPage rsMarginReportPage = new RSMarginReportPage();
//        rsMarginReportPage.getDataFromMarginReport();
        RSSaleGapReportPage rsSaleGapReportPage = new RSSaleGapReportPage();
        rsSaleGapReportPage.getDataFromSGReport("2023", "3");
        rsSaleGapReportPage.getDataSourceForSGReport("2023", "3");
//        rsSaleGapReportPage.verifySGReport();
        System.out.println("Done");
    }
}
