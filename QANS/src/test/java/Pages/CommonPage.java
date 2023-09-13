package Pages;

import Helpers.KeywordWebUI;
import Helpers.Manager.FileReaderManager;
import org.apache.commons.collections4.list.SetUniqueList;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

import java.util.ArrayList;
import java.util.List;

public class CommonPage extends KeywordWebUI {
    private WebDriver driver;
    public CommonPage(WebDriver driver){
        super(driver);
        PageFactory.initElements(driver, this);
        this.driver = driver;
    }
}
