package Pages.Reports;

import Helpers.KeywordWebUI;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class GapToSFPartNumberDetailsPage extends KeywordWebUI {
    private WebDriver driver;
    @FindBy(xpath = "//*/div[contains(text(),'GAP to SF - Part Number Detail')]")
    private WebElement txtTitle;
    public GapToSFPartNumberDetailsPage(WebDriver driver){
        super(driver);
        this.driver = driver;
    }
    public void shouldSeeTitle(String title){

    }


}
