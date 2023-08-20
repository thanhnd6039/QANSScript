package Helpers.Manager;

import Pages.Reports.RSLoginPage;
import org.openqa.selenium.WebDriver;

public class PageObjectManager {
    private WebDriver driver;
    private RSLoginPage rsLoginPage;
    public PageObjectManager(WebDriver driver){
        this.driver = driver;
    }
    public RSLoginPage getRsLoginPage(){
        return (rsLoginPage == null) ? rsLoginPage = new RSLoginPage(driver) : rsLoginPage;
    }

}
