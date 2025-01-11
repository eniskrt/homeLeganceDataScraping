package utilities;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.junit.After;
import org.junit.Before;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.time.Duration;

public class TestBase {

    protected static WebDriver driver;

    @Before
    public void setup() {
       /* WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(11));*/

        // ChromeOptions ile tarayıcı ayarları
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--headless=new"); // Yeni headless modu (daha stabil)
        options.addArguments("--disable-popup-blocking"); // Pop-up engelleme devre dışı
        options.addArguments("--remote-allow-origins=*"); // Çapraz köken hatalarını önler

        driver = new ChromeDriver(options);
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(11));https://github.com/user-attachments/assets/8fbe66da-cc51-439f-b61b-a6b035859d47
        // Login işlemi
        driver.get(ConfigReader.getProperty("loginUrl"));

        WebElement usernameInput = driver.findElement(By.id("username"));
        usernameInput.sendKeys(ConfigReader.getProperty("userName"));

        WebElement passwordInput = driver.findElement(By.id("password"));
        passwordInput.sendKeys(ConfigReader.getProperty("password"));

        WebElement signInBtn = driver.findElement(By.id("submitBtn"));
        signInBtn.click();
    }

    @After
    public void teardown() {
        if (driver != null) {
            driver.quit();
        }
    }
}
