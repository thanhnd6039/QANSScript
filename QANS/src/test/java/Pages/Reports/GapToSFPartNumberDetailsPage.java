package Pages.Reports;

import Helpers.KeywordWebUI;
import org.openqa.selenium.WebDriver;

public class GapToSFPartNumberDetailsPage extends KeywordWebUI {
    private WebDriver driver;
    public GapToSFPartNumberDetailsPage(WebDriver driver){
        super(driver);
        this.driver = driver;
    }

}
