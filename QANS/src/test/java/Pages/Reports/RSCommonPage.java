package Pages.Reports;

import Helpers.KeywordWebUI;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

public class RSCommonPage extends KeywordWebUI {
    private WebDriver driver;
    public RSCommonPage(WebDriver driver){
        super(driver);
        PageFactory.initElements(driver, this);
        this.driver = driver;
    }
}
