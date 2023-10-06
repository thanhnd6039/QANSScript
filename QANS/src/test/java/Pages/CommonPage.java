package Pages;

import Helpers.KeywordWebUI;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

import java.util.Calendar;
import java.util.GregorianCalendar;

public class CommonPage extends KeywordWebUI{
    private WebDriver driver;
    protected final String CONFIGURE_FILE_PATH = "C:\\CucumberFramework\\Config\\Configuration.properties";
    public CommonPage(WebDriver driver){
        super(driver);
        PageFactory.initElements(driver, this);
        this.driver = driver;
    }
    public int getCurrentQuarter(){
        int currentQuarter = 0;
        Calendar calendar = new GregorianCalendar();
        currentQuarter = (calendar.get(Calendar.MONTH) / 3) + 1;
        return currentQuarter;
    }

    public int getCurrentYear(){
        int currentYear = 0;
        Calendar calendar = new GregorianCalendar();
        currentYear = calendar.get(Calendar.YEAR);
        return currentYear;
    }

}
