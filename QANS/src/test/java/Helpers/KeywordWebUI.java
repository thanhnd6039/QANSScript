package Helpers;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

public class KeywordWebUI {
    private WebDriver driver;
    private WebDriverWait wait;

    public KeywordWebUI(WebDriver driver){
        this.driver = driver;
        wait = new WebDriverWait(driver, 60);
    }


}
