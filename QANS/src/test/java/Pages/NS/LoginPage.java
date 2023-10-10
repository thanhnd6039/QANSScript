package Pages.NS;

import Pages.CommonPage;
import org.junit.Assert;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

public class LoginPage extends CommonPage {
    WebDriver driver;
    public LoginPage(WebDriver driver){
        super(driver);
        this.driver = driver;
        PageFactory.initElements(driver, this);
    }
    public void navigateToNS(){
        String url = "https://system.netsuite.com/app/login/secure/enterpriselogin.nl?whence=";
        driver.get(url);
        driver.manage().window().maximize();
        Assert.assertTrue("Cannot navigate to NS", waitForPageLoadComplete());
    }

}
