package Helpers.Manager;

import Pages.CommonPage;
import Pages.Reports.RSGapToSFPNDetailsPage;
import Pages.Reports.RSLoginPage;
import org.openqa.selenium.WebDriver;

public class PageObjectManager {
    private WebDriver driver;
    private CommonPage rsCommonPage;
    private RSLoginPage rsLoginPage;
    private RSGapToSFPNDetailsPage rsGapToSFPNDetailsPage;
    public PageObjectManager(WebDriver driver){
        this.driver = driver;
    }
    public CommonPage getRsCommonPage(){
        return (rsCommonPage == null) ? rsCommonPage = new CommonPage(driver) : rsCommonPage;
    }
    public RSLoginPage getRsLoginPage(){
        return (rsLoginPage == null) ? rsLoginPage = new RSLoginPage(driver) : rsLoginPage;
    }
    public RSGapToSFPNDetailsPage getRsGapToSFPNDetailsPage(){
        return (rsGapToSFPNDetailsPage == null) ? rsGapToSFPNDetailsPage = new RSGapToSFPNDetailsPage(driver) : rsGapToSFPNDetailsPage;
    }

}
