package Helpers.Manager;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;

import java.util.HashMap;
import java.util.Map;

public class WebDriverManager {

    private WebDriver driver;
    private String environmentType;
    private String driverType;
    private String driverPath;
    private String downloadFilePath;
    private final String CHROME_DRIVER_PROPERTY = "webdriver.chrome.driver";
    private final String FIREFOX_DRIVER_PROPERTY = "webdriver.gecko.driver";
    private final String INTERNETEXPLORER_DRIVER_PROPERTY = "webdriver.ie.driver";
    private final String EDGE_DRIVER_PROPERTY = "webdriver.edge.driver";

    public WebDriverManager(){
        driverPath = "C:\\CucumberFramework\\Drivers\\Web\\Windows\\";
        driverPath = "C:\\CucumberFramework\\Downloads\\";
        environmentType = FileReaderManager.getInstance().getPropertyFileReader("C:\\CucumberFramework\\Config\\Configuration.properties").getValueFromKey("ENVIRONMENT");
        driverType = FileReaderManager.getInstance().getPropertyFileReader("C:\\CucumberFramework\\Config\\Configuration.properties").getValueFromKey("BROWSER");
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
        Map<String, Object> prefs = new HashMap<String, Object>();
        prefs.put("download.default_directory", downloadFilePath);

        if (driverType == null || driverType.equalsIgnoreCase("chrome")){
            driverPath += "chromedriver.exe";
            System.setProperty(CHROME_DRIVER_PROPERTY, driverPath);
            ChromeOptions options = new ChromeOptions();
            options.setExperimentalOption("prefs", prefs);
            driver = new ChromeDriver(options);
        }
        else if (driverType.equalsIgnoreCase("firefox")){
            driverPath += "geckodriver.exe";
            System.setProperty(FIREFOX_DRIVER_PROPERTY, driverPath);
            driver = new FirefoxDriver();
        }
        else if (driverType.equalsIgnoreCase("internetexplorer")){
            driverPath += "IEDriverServer.exe";
            System.setProperty(INTERNETEXPLORER_DRIVER_PROPERTY, driverPath);
            driver = new InternetExplorerDriver();
        }
        else if (driverType.equalsIgnoreCase("edge")){
            driverPath += "msedgedriver.exe";
            System.setProperty(EDGE_DRIVER_PROPERTY, driverPath);
            driver = new EdgeDriver();
        }
        else {
            throw new RuntimeException(String.format("The value of BROWSER key in the file Configuration.properties is wrong"));
        }
        return driver;
    }

    private WebDriver createRemoteDriver(){
        return driver;
    }
}
