package Pages.Reports;

import Helpers.KeywordWebUI;
import Helpers.Manager.FileReaderManager;
import org.openqa.selenium.WebDriver;

public class RSLoginPage extends KeywordWebUI {
    private WebDriver driver;
    private String rootUrl = "";
    public RSLoginPage(WebDriver driver){
        super(driver);
        this.driver = driver;
        String email = FileReaderManager.getInstance().getPropertyFileReader("C:\\CucumberFramework\\Config\\Configuration.properties").getValueFromKey("EMAIL");
        String pass = FileReaderManager.getInstance().getPropertyFileReader("C:\\CucumberFramework\\Config\\Configuration.properties").getValueFromKey("OUTLOOK_PASSWORD");
        rootUrl = String.format("http://%s:%s@reports/ReportServer", email, pass);
    }
    public void loginToReport(String title){
        String url = "";
        switch (title){
            case "GAP to SF Part Number Details":
                url = rootUrl + "/Pages/ReportViewer.aspx?/NetSuite+Reports/GAP+to+SF+-+Part+Number+Detail&rs:Command=Render&FromYear=2021&FromTime=1&ToYear=2023&ToTime=3&SalesPerson=29336&SalesPerson=98&SalesPerson=96&SalesPerson=28095";
                break;
            default:
                System.out.println("no match");
                break;
        }
        driver.get(url);
    }
    public void shouldSeeTitleOfReport(String title){

    }



}
