package Helpers;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedCondition;
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
    public boolean waitForPageLoadComplete(){
        try{
            wait.until(new ExpectedCondition<Boolean>() {
                public Boolean apply(WebDriver driver){
                    JavascriptExecutor js = (JavascriptExecutor) driver;
                    Object result = js.executeScript("return document.readyState");
                    if (result.toString().equalsIgnoreCase("complete")){
                        return true;
                    }else {
                        return false;
                    }
                }
            });
            return true;
        }
        catch (Exception e){
//            LogReader.error(String.format("Error: %s", e.getMessage()));
            System.out.println(String.format("Error: %s", e.getMessage()));
            return false;
        }
    }


}
