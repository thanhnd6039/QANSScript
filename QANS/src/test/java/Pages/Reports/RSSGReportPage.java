package Pages.Reports;

import Helpers.KeywordWebUI;
import Pages.CommonPage;
import org.junit.Assert;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class RSSGReportPage extends CommonPage {
    private WebDriver driver;
    @FindBy(xpath = "//span[contains(text(),'Sales Gap Report')]")
    private WebElement txtTitle;
    public RSSGReportPage(WebDriver driver){
        super(driver);
        this.driver = driver;
        PageFactory.initElements(driver, this);
    }
    public void shouldSeeTitleOfReport(String expectedTitle){
        Assert.assertTrue(waitForElementVisibility(txtTitle));
        String actualTitle = getTextFromElement(txtTitle);
        System.out.println(String.format("actualTitle: %s", actualTitle));
        Assert.assertTrue(actualTitle.contains(expectedTitle));

    }
}
