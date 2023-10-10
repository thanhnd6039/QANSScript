package Helpers.Manager;

import Pages.CommonPage;
import Pages.NS.LoginPage;
import Pages.Reports.RSLoginPage;
import Pages.Reports.RSSGReportPage;
import org.openqa.selenium.WebDriver;

public class PageObjectManager {
    private WebDriver driver;
    private CommonPage rsCommonPage;
    private RSLoginPage rsLoginPage;
    private RSSGReportPage rssgReportPage;
    private LoginPage loginPage;

    public PageObjectManager(WebDriver driver){
        this.driver = driver;
    }
    public CommonPage getRsCommonPage(){
        return (rsCommonPage == null) ? rsCommonPage = new CommonPage(driver) : rsCommonPage;
    }
    public RSLoginPage getRsLoginPage(){
        return (rsLoginPage == null) ? rsLoginPage = new RSLoginPage(driver) : rsLoginPage;
    }
    public RSSGReportPage getRssgReportPage(){
        return (rssgReportPage == null) ? rssgReportPage = new RSSGReportPage(driver) : rssgReportPage;
    }
    public LoginPage getLoginPage(){
        return (loginPage == null) ? loginPage = new LoginPage(driver) : loginPage;
    }


}
