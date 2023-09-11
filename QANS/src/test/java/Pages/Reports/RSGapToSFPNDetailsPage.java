package Pages.Reports;

import Helpers.Manager.FileReaderManager;
import Pages.CommonPage;
import org.apache.commons.collections4.list.SetUniqueList;
import org.junit.Assert;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import java.util.*;

public class RSGapToSFPNDetailsPage extends CommonPage {
    private WebDriver driver;

    @FindBy(xpath = "//*/div[contains(text(),'GAP to SF - Part Number Detail')]")
    private WebElement txtTitle;
    public RSGapToSFPNDetailsPage(WebDriver driver){
        super(driver);
        PageFactory.initElements(driver, this);
        this.driver = driver;
    }
    public void shouldSeeTitle(String title){
        Assert.assertTrue(waitForElementVisibility(txtTitle));
        String actualTitle = getTextFromElement(txtTitle);
        Assert.assertEquals(title, actualTitle);
    }


    public void getSourceDataForGapToSFPNDetailReport(){

    }




}
