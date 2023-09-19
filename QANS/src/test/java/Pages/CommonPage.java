package Pages;

import Helpers.KeywordWebUI;
import Helpers.Manager.FileReaderManager;
import org.apache.commons.collections4.list.SetUniqueList;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.List;

public class CommonPage extends KeywordWebUI{
    private WebDriver driver;
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
