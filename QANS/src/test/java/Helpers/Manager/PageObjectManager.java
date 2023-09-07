package Helpers.Manager;

import Pages.CommonPage;
import Pages.Reports.RSGapToSFPartNumberDetailsPage;
import Pages.Reports.RSLoginPage;
import org.openqa.selenium.WebDriver;

public class PageObjectManager {
    private WebDriver driver;
    private CommonPage rsCommonPage;
    private RSLoginPage rsLoginPage;
    private RSGapToSFPartNumberDetailsPage rsGapToSFPartNumberDetailsPage;
    public PageObjectManager(WebDriver driver){
        this.driver = driver;
    }
    public CommonPage getRsCommonPage(){
        return (rsCommonPage == null) ? rsCommonPage = new CommonPage(driver) : rsCommonPage;
    }
    public RSLoginPage getRsLoginPage(){
        return (rsLoginPage == null) ? rsLoginPage = new RSLoginPage(driver) : rsLoginPage;
    }
    public RSGapToSFPartNumberDetailsPage getRsGapToSFPartNumberDetailsPage(){
        return (rsGapToSFPartNumberDetailsPage == null) ? rsGapToSFPartNumberDetailsPage = new RSGapToSFPartNumberDetailsPage(driver) : rsGapToSFPartNumberDetailsPage;
    }

}
