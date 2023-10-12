package Pages.NS;

import Helpers.Manager.FileReaderManager;
import Pages.CommonPage;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import java.util.List;

public class LoginPage extends CommonPage {
    private WebDriver driver;
    @FindBy(id = "email")
    private WebElement txtEmail;
    @FindBy(id = "password")
    private WebElement txtPass;
    @FindBy(id = "login-submit")
    private WebElement btnLogIn;
    @FindBy(id = "uif43")
    private WebElement txtTitle;
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
    public void loginToNS(){
        String email = FileReaderManager.getInstance().getPropertyFileReader(CONFIGURE_FILE_PATH).getValueFromKey("EMAIL");
        String pass = FileReaderManager.getInstance().getPropertyFileReader(CONFIGURE_FILE_PATH).getValueFromKey("NS_PASS");
        setTextToElement(txtEmail, email);
        setTextToElement(txtPass, pass);
        Assert.assertTrue("Cannot click to the LogIn button", clickToElement(btnLogIn));
    }
    public void chooseAccountType(String expectedAccountType){
        String accountTableXpath = String.format("//*/table[@class='wideTable']/tbody/tr[2]/td/table/tbody/tr");
        List<WebElement> accountRows = getListOfElementsByTable(accountTableXpath);

        for (int rowIndex = 2; rowIndex <= accountRows.size(); rowIndex++){
            String accountTypeXpath = String.format("//*/table[@class='wideTable']/tbody/tr[2]/td/table/tbody/tr[%d]/td[2]", rowIndex);
            WebElement accountTypeElement = driver.findElement(By.xpath(accountTypeXpath));
            String actualAccountType = getTextFromElement(accountTypeElement);
            if (actualAccountType.equalsIgnoreCase(expectedAccountType)){
                String chooseAccountXpath = String.format("//*/table[@class='wideTable']/tbody/tr[2]/td/table/tbody/tr[%d]/td[3]/a", rowIndex);
                WebElement chooseAccountElement = driver.findElement(By.xpath(chooseAccountXpath));
                Assert.assertTrue("Cannot click to the Choose account link", clickToElement(chooseAccountElement));
                break;
            }
        }
    }
    public void shouldSeeTitle(String expectedTitle){
        Assert.assertTrue(waitForElementVisibility(txtTitle));
        String actualTitle = getTextFromElement(txtTitle);
        Assert.assertTrue(String.format("Cannot see the title %s", expectedTitle), actualTitle.contains(expectedTitle));
    }

}
