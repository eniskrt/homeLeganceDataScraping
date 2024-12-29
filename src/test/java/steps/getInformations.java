package steps;

import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import utillities.TestBase;

import java.util.List;

public class getInformations extends TestBase {

    private static final Logger log = LoggerFactory.getLogger(getInformations.class);
    /*
    @Test
    public void newArrivals() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement newArrivalsLink = driver.findElement(By.xpath("(//a[@href='/new_arrivals/'])[1]"));
        newArrivalsLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();
        actions.scrollByAmount(0,275).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 4 == 0) {
                actions.scrollByAmount(0,410).perform();
            }

            WebElement product = productList.get(i);
            product.click();
            Thread.sleep(1000);

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");

            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j == 0){
                        actions.scrollByAmount(0,750).perform();
                    } else if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                    Thread.sleep(1000);
                }
            } else {
                productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                productPrice =driver.findElement(By.id("price_block")).getText();
                productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                System.out.println(productCode);
                System.out.println(productPrice);
                System.out.println(productStockStatus);
                System.out.println(productRemoteStockStatus);
                System.out.println("----------");
            }
            //new arrival sayfsına geri dönüş
            driver.navigate().back();
            Thread.sleep(1000);
        }
    }*/

    /*
    @Test
    public void bedroom() {

    }

    @Test
    public void dining() {

    }

    @Test
    public void youth() {

    }*/

    @Test
    public void seating() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement seatingLink = driver.findElement(By.xpath("(//a[@href='/seating/'])[1]"));
        actions.moveToElement(seatingLink).perform();
        Thread.sleep(1000);

        WebElement browseAllSeatingLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[5]//li[@class='menu-item-li']"));
        browseAllSeatingLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();

        actions.scrollByAmount(0,400).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,555).perform();
            }

            WebElement product = productList.get(i);
            product.click();
            //Thread.sleep(1000);

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");

            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j == 0){
                        actions.scrollByAmount(0,750).perform();
                    } else if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    //Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                }
            } else {
                productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                productPrice =driver.findElement(By.id("price_block")).getText();
                productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                System.out.println(productCode);
                System.out.println(productPrice);
                System.out.println(productStockStatus);
                System.out.println(productRemoteStockStatus);
                System.out.println("----------");
            }
            //seating sayfsına geri dönüş
            driver.navigate().back();
        }
    }
    /*
    @Test
    public void occasional() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement occasionalLink = driver.findElement(By.xpath("(//a[@href='/occasional/'])[1]"));
        actions.moveToElement(occasionalLink).perform();
        Thread.sleep(1000);

        WebElement browseAllOccasionalLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[6]//li[@class='menu-item-li']"));
        browseAllOccasionalLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();

        actions.scrollByAmount(0,400).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,560).perform();
            }

            WebElement product = productList.get(i);
            product.click();
            //Thread.sleep(1000);

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            String productHeading = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText();
            boolean isContainsCollection =productHeading.contains("Collection");
            if (productHeading.equals("3503BK Occasional-Muriel Collection")){
                // bu üründe resim gelmiyor
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    //Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                }

            } else if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j == 0){
                        actions.scrollByAmount(0,735).perform();
                    } else if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    //Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                }
            } else {
                productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                productPrice =driver.findElement(By.id("price_block")).getText();
                productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                System.out.println(productCode);
                System.out.println(productPrice);
                System.out.println(productStockStatus);
                System.out.println(productRemoteStockStatus);
                System.out.println("----------");
            }
            //occasional sayfsına geri dönüş
            driver.navigate().back();
        }
    }*/
    /*
    @Test
    public void office() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement officeLink = driver.findElement(By.xpath("(//a[@href='/office/'])[1]"));
        actions.moveToElement(officeLink).perform();
        Thread.sleep(1000);

        WebElement browseAllOfficeLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[7]//li[@class='menu-item-li']"));
        browseAllOfficeLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();

        actions.scrollByAmount(0,400).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,555).perform();
            }

            WebElement product = productList.get(i);
            product.click();
            //Thread.sleep(1000);

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");

            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j == 0){
                        actions.scrollByAmount(0,750).perform();
                    } else if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    //Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                }
            } else {
                productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                productPrice =driver.findElement(By.id("price_block")).getText();
                productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                System.out.println(productCode);
                System.out.println(productPrice);
                System.out.println(productStockStatus);
                System.out.println(productRemoteStockStatus);
                System.out.println("----------");
            }
            //office sayfsına geri dönüş
            driver.navigate().back();
        }
    }
    *//*
    @Test
    public void accent() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement accentLink = driver.findElement(By.xpath("(//a[@href='/accent/'])[1]"));
        actions.moveToElement(accentLink).perform();
        Thread.sleep(1000);

        WebElement browseAllAccentLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[8]//li[@class='menu-item-li']"));
        browseAllAccentLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();

        actions.scrollByAmount(0,400).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,555).perform();
            }

            WebElement product = productList.get(i);
            product.click();
            //Thread.sleep(1000);

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");

            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j == 0){
                        actions.scrollByAmount(0,750).perform();
                    } else if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    //Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                }
            } else {
                productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                productPrice =driver.findElement(By.id("price_block")).getText();
                productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                System.out.println(productCode);
                System.out.println(productPrice);
                System.out.println(productStockStatus);
                System.out.println(productRemoteStockStatus);
                System.out.println("----------");
            }
            //accent sayfsına geri dönüş
            driver.navigate().back();
        }
    }*/
    /*
    @Test
    public void media() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement mediaLink = driver.findElement(By.xpath("(//a[@href='/media/'])[1]"));
        actions.moveToElement(mediaLink).perform();
        Thread.sleep(1000);

        WebElement browseAllMediaLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[9]//li[@class='menu-item-li']"));
        browseAllMediaLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();

        actions.scrollByAmount(0,400).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,555).perform();
            }

            WebElement product = productList.get(i);
            product.click();
            //Thread.sleep(1000);

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");

            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j == 0){
                        actions.scrollByAmount(0,750).perform();
                    } else if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    //Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                }
            } else {
                productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                productPrice =driver.findElement(By.id("price_block")).getText();
                productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                System.out.println(productCode);
                System.out.println(productPrice);
                System.out.println(productStockStatus);
                System.out.println(productRemoteStockStatus);
                System.out.println("----------");
            }
            //media sayfsına geri dönüş
            driver.navigate().back();
        }
    }*/
    /*
    @Test
    public void lighting() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement lightingLink = driver.findElement(By.xpath("(//a[@href='/lighting/'])[1]"));
        actions.moveToElement(lightingLink).perform();
        Thread.sleep(1000);

        WebElement browseAllLightingLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem last-child'])[1]//li[@class='menu-item-li']"));
        browseAllLightingLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();

        actions.scrollByAmount(0,400).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,555).perform();
            }

            WebElement product = productList.get(i);
            product.click();
            //Thread.sleep(1000);

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");

            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j == 0){
                        actions.scrollByAmount(0,750).perform();
                    } else if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    //Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                }
            } else {
                productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                productPrice =driver.findElement(By.id("price_block")).getText();
                productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                System.out.println(productCode);
                System.out.println(productPrice);
                System.out.println(productStockStatus);
                System.out.println(productRemoteStockStatus);
                System.out.println("----------");
            }
            //lighting sayfsına geri dönüş
            driver.navigate().back();
        }
    }*/
    /*
    @Test
    public void mattress() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement mattressLink = driver.findElement(By.xpath("(//a[@href='/mattress/'])[1]"));
        actions.moveToElement(mattressLink).perform();
        Thread.sleep(1000);

        WebElement browseAllMattressLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem last-child'])[2]//li[@class='menu-item-li']"));
        browseAllMattressLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();

        actions.scrollByAmount(0,400).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,555).perform();
            }

            WebElement product = productList.get(i);
            product.click();
            //Thread.sleep(1000);

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");

            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j == 0){
                        actions.scrollByAmount(0,750).perform();
                    } else if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    //Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                }
            } else {
                productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                productPrice =driver.findElement(By.id("price_block")).getText();
                productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                System.out.println(productCode);
                System.out.println(productPrice);
                System.out.println(productStockStatus);
                System.out.println(productRemoteStockStatus);
                System.out.println("----------");
            }
            //mattress sayfsına geri dönüş
            driver.navigate().back();
        }
    }*/
    /*
    @Test
    public void onSale() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement onSaleLink = driver.findElement(By.xpath("(//a[@href='/cal_on_sale/'])[1]"));
        actions.moveToElement(onSaleLink).perform();
        Thread.sleep(1000);

        WebElement browseAllOnSaleLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem last-child'])[3]//li[@class='menu-item-li']"));
        browseAllOnSaleLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();

        actions.scrollByAmount(0,400).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,555).perform();
            }

            WebElement product = productList.get(i);
            product.click();
            //Thread.sleep(1000);

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");

            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 0; j < subProductList.size(); j++) {
                    // Listeyi yeniden alın
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                            "//a"));
                    WebElement subProduct = subProductListInLoop.get(j);

                    //Collections ürünlerin alt listeleri için iterator
                    if (j == 0){
                        actions.scrollByAmount(0,750).perform();
                    } else if (j % 5 == 0) {
                        actions.scrollByAmount(0,425).perform();
                    }
                    subProduct.click();
                    //Thread.sleep(1000);
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    System.out.println(productCode);
                    System.out.println(productPrice);
                    System.out.println(productStockStatus);
                    System.out.println(productRemoteStockStatus);
                    System.out.println("----------");

                    // Collections ürüne geri dönüş
                    driver.navigate().back();
                }
            } else {
                productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                productPrice =driver.findElement(By.id("price_block")).getText();
                productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                System.out.println(productCode);
                System.out.println(productPrice);
                System.out.println(productStockStatus);
                System.out.println(productRemoteStockStatus);
                System.out.println("----------");
            }
            //mattress sayfsına geri dönüş
            driver.navigate().back();
        }
    }*/
}
