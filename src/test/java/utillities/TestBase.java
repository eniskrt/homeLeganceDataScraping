package utillities;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.time.Duration;

public class TestBase {

    protected static WebDriver driver;

    @BeforeClass
    public static void setup() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(15));

        driver.get(ConfigReader.getProperty("loginUrl"));

        WebElement usernameInput = driver.findElement(By.id("username"));
        usernameInput.sendKeys(ConfigReader.getProperty("userName"));

        WebElement passwordInput = driver.findElement(By.id("password"));
        passwordInput.sendKeys(ConfigReader.getProperty("password"));

        WebElement signInBtn = driver.findElement(By.id("submitBtn"));
        signInBtn.click();
    }

    @AfterClass
    public static void teardown() {
        driver.quit();
    }
}
