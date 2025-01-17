package steps;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.AfterClass;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import utilities.TestBase;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class GetInformations extends TestBase {

    private static Set<String[]> allProductsInformation = new HashSet<>(); // Ortak veri listesi
    @Test
    public void newArrivals() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement newArrivalsLink = driver.findElement(By.xpath("(//a[@href='/new_arrivals/'])[1]"));
        newArrivalsLink.click();

        WebElement allBtn = driver.findElement(By.xpath("//span[text()='ALL']"));
        allBtn.click();
        actions.scrollByAmount(0,150).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 1; i <= productList.size(); i++) {
            Thread.sleep(500);
            if (i == 1) {
                // İlk ürün için kaydırma işlemi atlanır
            }

            productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));
            WebElement product;
            try {
                product = productList.get(i - 1);
                product.click();
            } catch (Exception e) {
                System.out.println(e.getMessage() + " i:" + i);
                Thread.sleep(1000);
                try {
                    product = productList.get(i - 1);
                    actions.click(product).perform();
                } catch (Exception ex) {
                    try {
                        product = productList.get(i - 1);
                        JavascriptExecutor jsExecutor = (JavascriptExecutor) driver;
                        jsExecutor.executeScript("arguments[0].click();", product);
                    } catch (Exception exc) {
                        System.out.println("Her halukarda tıklama başarısız");
                    }
                }
            }
            if (i % 4 == 0) {
                actions.scrollByAmount(0, 570).perform();
            }

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = false;
            try {
                isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
            } catch (Exception e) {
                System.out.println("Ürün başlığı okunamadı" + e.getMessage());
            }
            if (isContainsCollection) {
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));
                    Thread.sleep(500);
                    if (j == 1) {
                        actions.scrollByAmount(0, 635).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);

                    if (j % 5 == 0) {
                        actions.scrollByAmount(0, 435).perform();
                        Thread.sleep(515);
                    }

                    try {
                        subProduct.click();
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " item: " + i + " element: " + j );
                        actions.click(subProduct).perform();
                    }

                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                        driver.navigate().back();
                    } catch (Exception ex) {
                        System.out.println(ex.getMessage() + "item:" + i + "element:" + j );
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice = driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    System.out.println(e.getMessage() + "item:" + i);
                }
            }
            driver.navigate().back();
        }
    }

    @Test
    public void bedroom() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement bedroomLink = driver.findElement(By.xpath("(//a[@href='/bedroom/'])[1]"));
        actions.moveToElement(bedroomLink).perform();
        Thread.sleep(1000);

        actions.scrollByAmount(0, 150).perform();

        WebElement browseAllSeatingLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[2]//li[@class='menu-item-li']"));
        browseAllSeatingLink.click();

        actions.scrollByAmount(0, 250).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));

            if (k >= 2) {
                actions.scrollByAmount(0, 250).perform();
            }

            for (int i = 1; i <= productList.size(); i++) {

                //productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                WebElement product;
                try {
                    product = productList.get(i - 1);
                    product.click();
                    if (i %3 == 0){
                        actions.scrollByAmount(0, 530).perform();
                    } else if (i == 34) {
                        actions.scrollByAmount(0, 200).perform();
                    }
                } catch (Exception e) {
                    System.out.println(e.getLocalizedMessage() + " i: " + i);
                    Thread.sleep(1000);
                    product = productList.get(i - 1);
                    actions.click(product).perform();
                    if (i %3 == 0){
                        actions.scrollByAmount(0, 530).perform();
                    } else if (i == 34) {
                        actions.scrollByAmount(0, 200).perform();
                    }
                }

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = false;
                try {
                    isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                } catch (Exception e) {
                    System.out.println("Ürün başlığı okunamadı " + e.getMessage() + " page: " + k);
                }
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        //List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 400).perform(); //todo 600
                        }

                        WebElement subProduct = subProductList.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
                            System.out.println(e.getMessage() + " page: " + k + " item: " + i + " element: " + j );
                            actions.click(subProduct).perform();
                        }

                        if (j % 5 == 0) {
                            actions.scrollByAmount(0, 425).perform(); //todo 440 olmalı sanki
                        }

                        try {
                            productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                            productPrice = driver.findElement(By.id("price_block")).getText();
                            productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                            productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                            String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                            allProductsInformation.add(productInfo);
                            driver.navigate().back();
                        } catch (Exception ex) {
                            System.out.println(ex.getMessage() + " page: " + k + " item: " + i + " element: " + j );
                            actions.sendKeys(Keys.TAB)
                                    .sendKeys(Keys.ENTER)
                                    .perform();
                        }
                    }
                } else {
                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " page: " + k + " item: " + i);
                    }
                }
                driver.navigate().back();
            }

            // Son sayfada değilsek, bir sonraki sayfaya geç
            if (k == pageSize){
                break;
            }
            if (k < pageSize-3) {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[16]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(2000); // Sayfanın yüklenmesi için bekleme süresi
            } else {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[12]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(1000); // Sayfanın yüklenmesi için bekleme süresi
            }
        }
    }

    @Test
    public void dining() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement diningLink = driver.findElement(By.xpath("(//a[@href='/dining/'])[1]"));
        actions.moveToElement(diningLink).perform();
        Thread.sleep(1000);

        actions.scrollByAmount(0, 150).perform();

        WebElement browseAllDiningLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[3]//li[@class='menu-item-li']"));
        browseAllDiningLink.click();

        actions.scrollByAmount(0, 250).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));

            if (k >= 2) {
                actions.scrollByAmount(0, 250).perform();
            }

            for (int i = 1; i <= productList.size(); i++) {

                //productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                WebElement product;
                try {
                    product = productList.get(i - 1);
                    product.click();
                    if (i %3 == 0){
                        actions.scrollByAmount(0, 530).perform();
                    } else if (i == 34) {
                        actions.scrollByAmount(0, 200).perform();
                    }
                } catch (Exception e) {
                    System.out.println(e.getLocalizedMessage() + " i: " + i);
                    Thread.sleep(1000);
                    product = productList.get(i - 1);
                    actions.click(product).perform();
                    if (i %3 == 0){
                        actions.scrollByAmount(0, 530).perform();
                    } else if (i == 34) {
                        actions.scrollByAmount(0, 200).perform();
                    }
                }

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = false;
                try {
                    isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                } catch (Exception e) {
                    System.out.println("Ürün başlığı okunamadı" + e.getMessage() + "page:" + k);
                }
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 400).perform();
                        }

                        WebElement subProduct = subProductListInLoop.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
                            System.out.println(e.getMessage() + " page: " + k + " item: " + i + " element: " + j );
                            actions.click(subProduct).perform();
                        }

                        if (j % 5 == 0) {
                            actions.scrollByAmount(0, 400).perform();
                        }

                        try {
                            productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                            productPrice = driver.findElement(By.id("price_block")).getText();
                            productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                            productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                            String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                            allProductsInformation.add(productInfo);
                            driver.navigate().back();
                        } catch (Exception ex) {
                            System.out.println(ex.getMessage() + "page:" + k + "item:" + i + "element:" + j );
                            actions.sendKeys(Keys.TAB)
                                    .sendKeys(Keys.ENTER)
                                    .perform();
                        }
                    }
                } else {
					try {
						productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
						productPrice = driver.findElement(By.id("price_block")).getText();
						productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
						productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
						String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
						allProductsInformation.add(productInfo);
					} catch (Exception e) {
                        System.out.println(e.getMessage() + "page:" + k + "item:" + i);
					}
				}
                driver.navigate().back();
            }

            // Son sayfada değilsek, bir sonraki sayfaya geç
            if (k == pageSize){
                break;
            }
            if (k < pageSize-3) {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[16]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(2000); // Sayfanın yüklenmesi için bekleme süresi
            } else {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[12]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(1000); // Sayfanın yüklenmesi için bekleme süresi
            }
        }
    }

    @Test
    public void youth() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement youthLink = driver.findElement(By.xpath("(//a[@href='/youth/'])[1]"));
        actions.moveToElement(youthLink).perform();
        Thread.sleep(1000);

        actions.scrollByAmount(0,250).perform();

        WebElement browseAllYouthLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[4]//li[@class='menu-item-li']"));
        browseAllYouthLink.click();

        actions.scrollByAmount(0, 250).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));

            if (k >= 2) {
                actions.scrollByAmount(0, 250).perform();
            }

            for (int i = 1; i <= productList.size(); i++) {

                //productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                WebElement product;
                try {
                    product = productList.get(i - 1);
                    product.click();
                    if (i %3 == 0){
                        actions.scrollByAmount(0, 530).perform();
                    } else if (i == 34) {
                        actions.scrollByAmount(0, 200).perform();
                    }
                } catch (Exception e) {
                    System.out.println(e.getLocalizedMessage() + " i: " + i);
                    Thread.sleep(1000);
                    product = productList.get(i - 1);
                    actions.click(product).perform();
                    if (i %3 == 0){
                        actions.scrollByAmount(0, 530).perform();
                    } else if (i == 34) {
                        actions.scrollByAmount(0, 200).perform();
                    }
                }

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = false;
                try {
                    isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                } catch (Exception e) {
                    System.out.println("Ürün başlığı okunamadı " + e.getMessage() + " page: " + k);
                }
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        //List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 400).perform(); //todo 600
                        }

                        WebElement subProduct = subProductList.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
                            System.out.println(e.getMessage() + " page: " + k + " item: " + i + " element: " + j );
                            actions.click(subProduct).perform();
                        }

                        if (j % 5 == 0) {
                            actions.scrollByAmount(0, 425).perform(); //todo 440 olmalı sanki
                        }

                        try {
                            productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                            productPrice = driver.findElement(By.id("price_block")).getText();
                            productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                            productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                            String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                            allProductsInformation.add(productInfo);
                            driver.navigate().back();
                        } catch (Exception ex) {
                            System.out.println(ex.getMessage() + " page: " + k + " item: " + i + " element: " + j );
                            actions.sendKeys(Keys.TAB)
                                    .sendKeys(Keys.ENTER)
                                    .perform();
                        }
                    }
                } else {
                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " page: " + k + " item: " + i);
                    }
                }
                driver.navigate().back();
            }

            // Son sayfada değilsek, bir sonraki sayfaya geç
            if (k == pageSize){
                break;
            }
            if (k < pageSize-3) {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[16]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(2000); // Sayfanın yüklenmesi için bekleme süresi
            } else {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[12]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(1000); // Sayfanın yüklenmesi için bekleme süresi
            }
        }
    }

    @Test
    public void seating() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement seatingLink = driver.findElement(By.xpath("(//a[@href='/seat/'])[1]"));
        actions.moveToElement(seatingLink).perform();
        Thread.sleep(1000);

        actions.scrollByAmount(0, 150).perform();

        WebElement browseAllSeatingLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[5]//li[@class='menu-item-li']"));
        browseAllSeatingLink.click();

        actions.scrollByAmount(0, 250).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productListFirst = driver.findElements(By.xpath("//div[@class='col position-relative']"));
            System.out.println("İlk alınan ürün listesi sayısı : " +  productListFirst.size());

            if (k >= 2) {
                actions.scrollByAmount(0, 250).perform();
            }

            for (int i = 1; i <= productListFirst.size(); i++) {

                WebElement product;
                try {
                    product = productListFirst.get(i - 1);
                    product.click();
                    if (i %3 == 0){
                        actions.scrollByAmount(0, 530).perform();
                    }
                } catch (Exception e) {
                    System.out.println(e.getLocalizedMessage() + " i: " + i);
                    Thread.sleep(1000);
                    try {
                        List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                        System.out.println("İkinci alınan ürün listesi sayısı : " +  productList.size());
                        product = productList.get(i - 1);
                        actions.click(product).perform();
                        if (i %3 == 0){
                            actions.scrollByAmount(0, 530).perform();
                        }
                    } catch (Exception ex) {
                        Thread.sleep(1000);
                        List<WebElement> productListLast = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                        System.out.println("\u001B[34mÜçüncü alınan ürün listesi sayısı : " + productListLast.size() + "\u001B[0m");
                        product = productListLast.get(i - 1);
                        actions.click(product).perform();
                        if (i %3 == 0){
                            actions.scrollByAmount(0, 530).perform();
                        }
                    }
                }

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = false;
                try {
                    isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                } catch (Exception e) {
                    System.out.println("Ürün başlığı okunamadı " + e.getMessage() + " page: " + k);
                }
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        //List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 400).perform(); //todo 600
                        }

                        WebElement subProduct = subProductList.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
                            System.out.println(e.getMessage() + " page: " + k + " item: " + i + " element: " + j );
                            try {
                                if (j == 1) {
                                    actions.scrollByAmount(0, 100).perform();
                                }
                                actions.click(subProduct).perform();
                            } catch (Exception ex) {
                                driver.navigate().refresh();
                                Thread.sleep(2000);
                                actions.click(subProduct).perform();
                            }
                        }

                        if (j % 5 == 0) {
                            actions.scrollByAmount(0, 425).perform(); //todo 440 olmalı sanki
                        }

                        try {
                            productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                            productPrice = driver.findElement(By.id("price_block")).getText();
                            productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                            productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                            String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                            allProductsInformation.add(productInfo);
                            driver.navigate().back();
                        } catch (Exception ex) {
                            System.out.println(ex.getMessage() + " page: " + k + " item: " + i + " element: " + j );
                            actions.sendKeys(Keys.TAB)
                                    .sendKeys(Keys.ENTER)
                                    .perform();
                        }
                    }
                } else {
                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " page: " + k + " item: " + i);
                    }
                }
                driver.navigate().back();
            }

            // Son sayfada değilsek, bir sonraki sayfaya geç
            if (k == pageSize){
                break;
            }
            if (k < pageSize-3) {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[16]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    Thread.sleep(1000);
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(2000); // Sayfanın yüklenmesi için bekleme süresi
            } else {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[12]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(1000); // Sayfanın yüklenmesi için bekleme süresi
            }
        }
    }

    @Test
    public void occasional() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement occasionalLink = driver.findElement(By.xpath("(//a[@href='/occasional/'])[1]"));
        actions.moveToElement(occasionalLink).perform();
        Thread.sleep(1000);

        WebElement browseAllOccasionalLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[6]//li[@class='menu-item-li']"));
        browseAllOccasionalLink.click();

        actions.scrollByAmount(0, 250).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));

            if (k >= 2) {
                actions.scrollByAmount(0, 250).perform();
            }

            for (int i = 1; i <= productList.size(); i++) {

                //productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                WebElement product;
                try {
                    product = productList.get(i - 1);
                    product.click();
                    if (i %3 == 0){
                        actions.scrollByAmount(0, 530).perform();
                    } else if (i == 34) {
                        actions.scrollByAmount(0, 200).perform();
                    }
                } catch (Exception e) {
                    System.out.println(e.getLocalizedMessage() + " i: " + i);
                    Thread.sleep(1000);
                    product = productList.get(i - 1);
                    actions.click(product).perform();
                    if (i %3 == 0){
                        actions.scrollByAmount(0, 530).perform();
                    } else if (i == 34) {
                        actions.scrollByAmount(0, 200).perform();
                    }
                }

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = false;
                try {
                    isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                } catch (Exception e) {
                    System.out.println("Ürün başlığı okunamadı " + e.getMessage() + " page: " + k);
                }
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        //List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 400).perform(); //todo 600'den 400'e indirdim
                        }

                        WebElement subProduct = subProductList.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
                            System.out.println(e.getMessage() + " page: " + k + " item: " + i + " element: " + j );
                            actions.click(subProduct).perform();
                        }

                        if (j % 5 == 0) {
                            actions.scrollByAmount(0, 400).perform(); //todo 425'ten 400'e indirdim
                        }

                        try {
                            productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                            productPrice = driver.findElement(By.id("price_block")).getText();
                            productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                            productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                            String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                            allProductsInformation.add(productInfo);
                            driver.navigate().back();
                        } catch (Exception ex) {
                            System.out.println(ex.getMessage() + "page:" + k + "item:" + i + "element:" + j );
                            actions.sendKeys(Keys.TAB)
                                    .sendKeys(Keys.ENTER)
                                    .perform();
                        }
                    }
                } else {
                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " page: " + k + " item: " + i);
                    }
                }
                driver.navigate().back();
            }

            // Son sayfada değilsek, bir sonraki sayfaya geç
            if (k == pageSize){
                break;
            }
            if (k < pageSize-3) {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[16]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(2000); // Sayfanın yüklenmesi için bekleme süresi
            } else {
                WebElement nextPageBtn = driver.findElement(By.xpath("(//li[@class='page-item'])[12]//a"));
                try {
                    nextPageBtn.click();
                } catch (Exception e) {
                    actions.click(nextPageBtn).perform();
                }
                Thread.sleep(1000); // Sayfanın yüklenmesi için bekleme süresi
            }
        }
    }

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

        actions.scrollByAmount(0,150).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 1; i <= productList.size(); i++) {

            productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
            WebElement product;
            try {
                product = productList.get(i - 1);
                product.click();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            } catch (Exception e) {
                System.out.println(e.getLocalizedMessage() + " i: " + i);
                Thread.sleep(1000);
                product = productList.get(i - 1);
                actions.click(product).perform();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            }

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = false;
            try {
                isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
            } catch (Exception e) {
                System.out.println("Ürün başlığı okunamadı" + e.getMessage());
            }
            if (isContainsCollection) {
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 635).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);

                    try {
                        subProduct.click();
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " item: " + i + " element: " + j );
                        actions.click(subProduct).perform();
                    }

                    if (j % 5 == 0) {
                        actions.scrollByAmount(0, 425).perform();
                    }

                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                        driver.navigate().back();
                    } catch (Exception ex) {
                        System.out.println(ex.getMessage() + "item:" + i + "element:" + j );
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice = driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    System.out.println(e.getMessage() + "item:" + i);
                }
            }
            driver.navigate().back();
        }
    }

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

        actions.scrollByAmount(0,250).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 1; i<= productList.size(); i++) {

            //productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
            WebElement product;
            try {
                product = productList.get(i - 1);
                product.click();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            } catch (Exception e) {
                System.out.println(e.getLocalizedMessage() + " i: " + i);
                Thread.sleep(1000);
                product = productList.get(i - 1);
                actions.click(product).perform();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            }

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = false;
            try {
                isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
            } catch (Exception e) {
                System.out.println("Ürün başlığı okunamadı" + e.getMessage());
            }
            if (isContainsCollection) {
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 600).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " item: " + i + " element: " + j );
                        actions.click(subProduct).perform();
                    }

                    if (j % 5 == 0) {
                        actions.scrollByAmount(0, 425).perform();
                    }

                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                        driver.navigate().back();
                    } catch (Exception ex) {
                        System.out.println(ex.getMessage() + "item:" + i + "element:" + j );
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice = driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    System.out.println(e.getMessage() + "item:" + i);
                }
            }
            driver.navigate().back();
        }
    }

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

        actions.scrollByAmount(0,250).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 1; i <= productList.size(); i++) {

            //productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
            WebElement product;
            try {
                product = productList.get(i - 1);
                product.click();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            } catch (Exception e) {
                System.out.println(e.getLocalizedMessage() + " i: " + i);
                Thread.sleep(1000);
                product = productList.get(i - 1);
                actions.click(product).perform();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            }

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = false;
            try {
                isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
            } catch (Exception e) {
                System.out.println("Ürün başlığı okunamadı" + e.getMessage());
            }
            if (isContainsCollection) {
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 600).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " item: " + i + " element: " + j );
                        actions.click(subProduct).perform();
                    }

                    if (j % 5 == 0) {
                        actions.scrollByAmount(0, 425).perform();
                    }

                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                        driver.navigate().back();
                    } catch (Exception ex) {
                        System.out.println(ex.getMessage() + "item:" + i + "element:" + j );
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice = driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    System.out.println(e.getMessage() + "item:" + i);
                }
            }
            driver.navigate().back();
        }
    }

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

        actions.scrollByAmount(0,250).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 1; i <= productList.size(); i++) {

            //productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
            WebElement product;
            try {
                product = productList.get(i - 1);
                product.click();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            } catch (Exception e) {
                System.out.println(e.getLocalizedMessage() + " i: " + i);
                Thread.sleep(1000);
                product = productList.get(i - 1);
                actions.click(product).perform();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            }

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = false;
            try {
                isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
            } catch (Exception e) {
                System.out.println("Ürün başlığı okunamadı" + e.getMessage());
            }
            if (isContainsCollection) {
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 600).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " item: " + i + " element: " + j );
                        actions.click(subProduct).perform();
                    }

                    if (j % 5 == 0) {
                        actions.scrollByAmount(0, 425).perform();
                    }

                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                        driver.navigate().back();
                    } catch (Exception ex) {
                        System.out.println(ex.getMessage() + "item:" + i + "element:" + j );
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice = driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    System.out.println(e.getMessage() + "item:" + i);
                }
            }
            driver.navigate().back();
        }
    }

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

        actions.scrollByAmount(0,250).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 1; i <= productList.size(); i++) {
            if (i == 1) {
                // İlk ürün için kaydırma işlemi atlanır
            }

            productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));
            WebElement product;
            try {
                product = productList.get(i -1);
                product.click();
            } catch (Exception e) {
                System.out.println("\u001B[34m1239. sıra\u001B[0m");
                System.out.println(e.getMessage() + " i: " + i);
                Thread.sleep(1000);
                product = productList.get(i - 1);
                actions.click(product).perform();
            }

            if (i % 3 == 0) {
                actions.scrollByAmount(0, 570).perform();
            }

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = false;
            try {
                isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
            } catch (Exception e) {
                System.out.println("\u001B[32m1259. sıra\u001B[0m");
                System.out.println("Ürün başlığı okunamadı " + e.getMessage());
            }
            if (isContainsCollection) {
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    //List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 600).perform();
                    }

                    WebElement subProduct = subProductList.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
                        System.out.println("\u001B[33m1277. sıra\u001B[0m");
                        System.out.println(e.getMessage() + " item: " + i + " element: " + j );
                        actions.click(subProduct).perform();
                    }

                    if (j % 5 == 0) {
                        actions.scrollByAmount(0, 425).perform();
                    }

                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                        driver.navigate().back();
                    } catch (Exception ex) {
                        System.out.println("\u001B[36m1295. sıra\u001B[0m");
                        System.out.println(ex.getMessage() + " item: " + i + " element: " + j );
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice = driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    System.out.println("\u001B[31m1311. sıra\u001B[0m");
                    System.out.println(e.getMessage() + " item: " + i);
                }
            }
            driver.navigate().back();
        }
    }

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

        actions.scrollByAmount(0,250).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 1; i <= productList.size(); i++) {

            //productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
            WebElement product;
            try {
                product = productList.get(i - 1);
                product.click();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            } catch (Exception e) {
                System.out.println(e.getLocalizedMessage() + " i: " + i);
                Thread.sleep(1000);
                product = productList.get(i - 1);
                actions.click(product).perform();
                if (i %3 == 0){
                    actions.scrollByAmount(0, 530).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
            }

            String productCode;
            String productPrice;
            String productStockStatus;
            String productRemoteStockStatus;

            boolean isContainsCollection = false;
            try {
                isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
            } catch (Exception e) {
                System.out.println("Ürün başlığı okunamadı" + e.getMessage());
            }
            if (isContainsCollection) {
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 600).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
                        System.out.println(e.getMessage() + " item: " + i + " element: " + j );
                        actions.click(subProduct).perform();
                    }

                    if (j % 5 == 0) {
                        actions.scrollByAmount(0, 425).perform();
                    }

                    try {
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        productPrice = driver.findElement(By.id("price_block")).getText();
                        productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                        productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                        String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                        allProductsInformation.add(productInfo);
                        driver.navigate().back();
                    } catch (Exception ex) {
                        System.out.println(ex.getMessage() + "item:" + i + "element:" + j );
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice = driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    System.out.println(e.getMessage() + "item:" + i);
                }
            }
            driver.navigate().back();
        }
    }

    @AfterClass
    public static void writeDataToExcel() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Product List");

        // Başlık stili
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex()); // Arka plan rengi
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // Dolgu tipi
        Font headerFont = workbook.createFont();
        headerFont.setBold(true); // Kalın yazı
        headerStyle.setFont(headerFont);

        // Başlıklar
        String[] headers = {"SKU NUMBER", "PRICE", "SANTA FE WAREHOUSE", "REMOTE WAREHOUSE"};
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(headers[i]);
            headerCell.setCellStyle(headerStyle); // Stili uygula
        }

        // Verileri ekle
        int rowCount = 1;
        for (String[] rowData : allProductsInformation) {
            Row row = sheet.createRow(rowCount++);
            for (int i = 0; i < rowData.length; i++) {
                row.createCell(i).setCellValue(rowData[i]);
            }
        }

        // Sütun genişliklerini otomatik ayarla
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Excel dosyasını kaydet
        String currentDate = new SimpleDateFormat("ddMMyyyy").format(new Date());
        String currentTime = new SimpleDateFormat("HHmm").format(new Date()); // Saati HHmm formatında al
        String fileName = "HomelegancePriceList_" + currentDate + "_" + currentTime + ".xlsx";
        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
            workbook.write(outputStream);
        }
        workbook.close();
        System.out.println("Excel dosyası başarıyla kaydedildi!");
    }
}
