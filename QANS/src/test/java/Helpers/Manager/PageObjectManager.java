package Helpers.Manager;

import Pages.Reports.RSCommonPage;
import Pages.Reports.RSGapToSFPartNumberDetailsPage;
import Pages.Reports.RSLoginPage;
import org.openqa.selenium.WebDriver;

public class PageObjectManager {
    private WebDriver driver;
    private RSCommonPage rsCommonPage;
    private RSLoginPage rsLoginPage;
    private RSGapToSFPartNumberDetailsPage rsGapToSFPartNumberDetailsPage;
    public PageObjectManager(WebDriver driver){
        this.driver = driver;
    }
    public RSCommonPage getRsCommonPage(){
        return (rsCommonPage == null) ? rsCommonPage = new RSCommonPage(driver) : rsCommonPage;
    }
    public RSLoginPage getRsLoginPage(){
        return (rsLoginPage == null) ? rsLoginPage = new RSLoginPage(driver) : rsLoginPage;
    }
    public RSGapToSFPartNumberDetailsPage getRsGapToSFPartNumberDetailsPage(){
        return (rsGapToSFPartNumberDetailsPage == null) ? rsGapToSFPartNumberDetailsPage = new RSGapToSFPartNumberDetailsPage(driver) : rsGapToSFPartNumberDetailsPage;
    }

}
