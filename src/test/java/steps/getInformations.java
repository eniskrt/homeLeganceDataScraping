package steps;

import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.devtools.v85.page.Page;
import org.openqa.selenium.interactions.Actions;
import utillities.TestBase;

import java.util.List;

public class getInformations extends TestBase {

    @Test
    public void newArrivals() throws InterruptedException {
        WebElement newArrivalsLink = driver.findElement(By.xpath("(//a[@href='/new_arrivals/'])[1]"));
        newArrivalsLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();

        List<WebElement> productList = driver.findElements(By.xpath("//div[@id='styleNumBox']//div[@class='col position-relative']"));

        for (int i = 0; i < productList.size(); i++) {
            WebElement product = productList.get(i);
            product.click();
            Thread.sleep(1000);

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
            System.out.println(isContainsCollection);
            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']//img"));
                    System.out.println(subProductListInLoop.size());
                    WebElement subProduct = subProductListInLoop.get(j);

                    // Tıklayın
                    subProduct.click();
                    Thread.sleep(1000);

                    // Geri dön
                    driver.navigate().back();
                    Thread.sleep(1000);
                }

            }
            driver.navigate().back();
            Thread.sleep(3000);
        }
        System.out.println(productList.size());
    }

    @Test
    public void bedroom() {

    }

    @Test
    public void dining() {

    }

    @Test
    public void youth() {

    }

    @Test
    public void seating() {

    }

    @Test
    public void occasional() {

    }

    @Test
    public void office() {

    }

    @Test
    public void accent() {

    }

    @Test
    public void media() {

    }

    @Test
    public void lighting() {

    }

    @Test
    public void mattress() {

    }

    @Test
    public void onSale() {

    }
}
