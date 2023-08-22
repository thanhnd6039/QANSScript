package Helpers.Manager;

import Pages.Reports.RSGapToSFPartNumberDetailsPage;
import Pages.Reports.RSLoginPage;
import org.openqa.selenium.WebDriver;

public class PageObjectManager {
    private WebDriver driver;
    private RSLoginPage rsLoginPage;
    private RSGapToSFPartNumberDetailsPage rsGapToSFPartNumberDetailsPage;
    public PageObjectManager(WebDriver driver){
        this.driver = driver;
    }
    public RSLoginPage getRsLoginPage(){
        return (rsLoginPage == null) ? rsLoginPage = new RSLoginPage(driver) : rsLoginPage;
    }
    public RSGapToSFPartNumberDetailsPage getRsGapToSFPartNumberDetailsPage(){
        return (rsGapToSFPartNumberDetailsPage == null) ? rsGapToSFPartNumberDetailsPage = new RSGapToSFPartNumberDetailsPage(driver) : rsGapToSFPartNumberDetailsPage;
    }

}
