package Pages.Reports;

import Helpers.KeywordWebUI;
import org.openqa.selenium.WebDriver;

public class RSLoginPage extends KeywordWebUI {
    private WebDriver driver;
    public RSLoginPage(WebDriver driver){
        super(driver);
        this.driver = driver;
    }

    public void loginToReportViewer(){

    }
}
