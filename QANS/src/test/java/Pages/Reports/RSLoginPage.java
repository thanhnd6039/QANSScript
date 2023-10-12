package Pages.Reports;

import Helpers.Manager.FileReaderManager;
import Pages.CommonPage;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

public class RSLoginPage extends CommonPage {
    private WebDriver driver;
    private String rootUrl = "";
    public RSLoginPage(WebDriver driver){
        super(driver);
        this.driver = driver;
        PageFactory.initElements(driver, this);
        String emailUsername = FileReaderManager.getInstance().getPropertyFileReader(CONFIGURE_FILE_PATH).getValueFromKey("EMAIL_USERNAME");
        String pass = FileReaderManager.getInstance().getPropertyFileReader(CONFIGURE_FILE_PATH).getValueFromKey("EMAIL_PASS");
        rootUrl = String.format("http://%s:%s@reports/ReportServer", emailUsername, pass);
    }
    public void loginToReport(String title){
        String url = "";
        switch (title){
            case "SG":
                url = rootUrl + "/Pages/ReportViewer.aspx?%2fNetsuite+Reports%2fSales+Gap+Report+NS+With+SO+Forecast&rs:Command=Render";
                break;
            default:
                System.out.println(String.format("The report %s is not supported. Please contact admin!"));
                break;
        }
        driver.get(url);
        driver.manage().window().maximize();
    }
    public void inputVerificationCode(){

    }




}
