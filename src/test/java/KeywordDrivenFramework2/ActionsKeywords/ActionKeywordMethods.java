package KeywordDrivenFramework2.ActionsKeywords;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import KeywordDrivenFramework2.Utilities.Constants;
public class ActionKeywordMethods {
    static WebDriver driver;
    static WebDriverWait wait;

    public void open_Browser(String browserName) {
        try {
            if (browserName.equalsIgnoreCase("Firefox")) {
                driver = new FirefoxDriver();
            } else if (browserName.equalsIgnoreCase("chrome")) {
                System.setProperty(Constants.chromeWDAtrribute,
                        Constants.chromeDriverPath);
                driver = new ChromeDriver();
            } else if (browserName.equalsIgnoreCase("IE")) {
                driver = new InternetExplorerDriver();
            }
        } catch (WebDriverException e) {
            System.out.println(e.getMessage());
        }
    }

    public void enter_URL(String URL) {
        driver.manage().window().maximize();
        driver.navigate().to(URL);
    }

    public By locatorValue(String locatorTpye, String value) {
        By by;
        switch (locatorTpye) {
            case "id":
                by = By.id(value);
                break;
            case "name":
                by = By.name(value);
                break;
            case "xpath":
                by = By.xpath(value);
                break;
            case "css":
                by = By.cssSelector(value);
                break;
            case "linkText":
                by = By.linkText(value);
                break;
            case "partialLinkText":
                by = By.partialLinkText(value);
                break;
            default:
                by = null;
                break;
        }
        return by;
    }

    public void enter_Text(String locatorType, String value, String text) {
        try {
            By locator;
            locator = locatorValue(locatorType, value);
            WebElement element = driver.findElement(locator);
            element.sendKeys(text);
        } catch (NoSuchElementException e) {
            System.err.format("No Element Found to enter text" + e);
        }
    }

    public void click_On_Link(String locatorType, String value) {
        try {
            By locator;
            locator = locatorValue(locatorType, value);
            WebDriverWait wait = new WebDriverWait(driver,10);
            WebElement element =  wait.until(ExpectedConditions.visibilityOfElementLocated(locator));
            element.click();
        } catch (NoSuchElementException e) {
            System.err.format("No Element Found to click on Link" + e);
        }
    }

    public void click_On_Button(String locatorType, String value) {
        try {
            By locator;
            locator = locatorValue(locatorType, value);
            WebElement element = driver.findElement(locator);
            element.click();
        } catch (NoSuchElementException e) {
            System.err.format("No Element Found to perform click" + e);
        }
    }

    public void close_Browser() {
        driver.quit();
    }
}
