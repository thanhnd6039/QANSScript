package Helpers.Manager;

import org.openqa.selenium.WebDriver;

public class WebDriverManager {

    private WebDriver driver;
    private String environmentType;

    public WebDriverManager(){

    }

    public WebDriver getDriver(){
        if (driver == null){
            driver = createDriver();
        }
        return driver;
    }

    private WebDriver createDriver(){
        if (environmentType == null || environmentType.equalsIgnoreCase("local")){
            driver = createLocalDriver();
        } else if (environmentType.equalsIgnoreCase("remote")) {
            driver = createRemoteDriver();
        }
        else {
            throw new RuntimeException(String.format("The value of ENVIRONMENT key in the file Configuration is wrong"));
        }
        return driver;
    }

    private WebDriver createLocalDriver(){
        return driver;
    }

    private WebDriver createRemoteDriver(){
        return driver;
    }
}
