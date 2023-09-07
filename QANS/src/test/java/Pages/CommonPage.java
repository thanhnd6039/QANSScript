package Pages;

import Helpers.KeywordWebUI;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

public class CommonPage extends KeywordWebUI {
    private WebDriver driver;
    public CommonPage(WebDriver driver){
        super(driver);
        PageFactory.initElements(driver, this);
        this.driver = driver;
    }
}
