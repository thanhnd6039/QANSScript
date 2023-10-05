package Helpers;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


public class KeywordWebUI {
    private WebDriver driver;
    private WebDriverWait wait;

    public KeywordWebUI(WebDriver driver){
        this.driver = driver;
        wait = new WebDriverWait(driver, 60);
    }
    public boolean waitForElementVisibility(WebElement element){
        try {
            wait.until(ExpectedConditions.visibilityOf(element));
            return true;
        }catch (Exception e){
            System.out.println(String.format("Error: %s", e.getMessage()));
            return false;
        }
    }

    public String getTextFromElement(WebElement element){
        String text = null;
        try {
            text = element.getText();
        }catch (Exception e){
            System.out.println(String.format("Error: %s", e.getMessage()));
        }
        return text;
    }

}
