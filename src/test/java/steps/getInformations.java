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
import utillities.TestBase;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class getInformations extends TestBase {

    private static List<String[]> allProductsInformation = new ArrayList<>(); // Ortak veri listesi

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
                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 750).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
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
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
				try {
					productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
					productPrice =driver.findElement(By.id("price_block")).getText();
					productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
					productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
					String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
					allProductsInformation.add(productInfo);
				} catch (Exception e) {
					productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    System.out.println(productCode + " Ürün bilgileri alınamadı");
				}
			}
            //new arrival sayfsına geri dönüş
            driver.navigate().back();
            Thread.sleep(1000);
        }
    }

    @Test
    public void bedroom() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement bedroomLink = driver.findElement(By.xpath("(//a[@href='/bedroom/'])[1]"));
        actions.moveToElement(bedroomLink).perform();
        Thread.sleep(1000);

        actions.scrollByAmount(0, 250).perform();

        WebElement browseAllSeatingLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[2]//li[@class='menu-item-li']"));
        browseAllSeatingLink.click();

        actions.scrollByAmount(0, 450).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));

            if (k >= 2) {
                actions.scrollByAmount(0, 450).perform();
            }

            for (int i = 0; i < productList.size(); i++) {
                System.out.println("i ==== " + i);
                if (i == 0) {
                    // İlk ürün için kaydırma işlemi atlanır
                } else if (i % 3 == 0) {
                    actions.scrollByAmount(0, 570).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
                productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                WebElement product = productList.get(i);
                product.click();

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 750).perform();
                        }

                        WebElement subProduct = subProductListInLoop.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
                            actions.click(subProduct).perform();
                        }

                        if (j % 5 == 0) {
                            System.out.println( j + "if else içinde");
                            actions.scrollByAmount(0, 425).perform();
                            System.out.println("perform gerçekleşti");
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
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        System.out.println(productCode + " Ürün bilgileri alınamadı.");
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

        actions.scrollByAmount(0, 250).perform();

        WebElement browseAllDiningLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[3]//li[@class='menu-item-li']"));
        browseAllDiningLink.click();

        actions.scrollByAmount(0, 450).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));

            if (k >= 2) {
                actions.scrollByAmount(0, 450).perform();
            }

            for (int i = 0; i < productList.size(); i++) {
                System.out.println("i ==== " + i);
                if (i == 0) {
                    // İlk ürün için kaydırma işlemi atlanır
                } else if (i % 3 == 0) {
                    actions.scrollByAmount(0, 570).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
                productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                WebElement product = productList.get(i);
                product.click();

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 750).perform();
                        }

                        WebElement subProduct = subProductListInLoop.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
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
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        System.out.println(productCode + " Ürün bilgileri alınamadı.");
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

        actions.scrollByAmount(0,350).perform();

        WebElement browseAllYouthLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[4]//li[@class='menu-item-li']"));
        browseAllYouthLink.click();

        actions.scrollByAmount(0, 450).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));

            if (k >= 2) {
                actions.scrollByAmount(0, 450).perform();
            }

            for (int i = 0; i < productList.size(); i++) {
                System.out.println("i ==== " + i);
                if (i == 0) {
                    // İlk ürün için kaydırma işlemi atlanır
                } else if (i % 3 == 0) {
                    actions.scrollByAmount(0, 570).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
                productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                WebElement product = productList.get(i);
                product.click();

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 750).perform();
                        }

                        WebElement subProduct = subProductListInLoop.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
                            actions.click(subProduct).perform();
                        }

                        if (j % 5 == 0) {
                            System.out.println( j + "if else içinde");
                            actions.scrollByAmount(0, 425).perform();
                            System.out.println("perform gerçekleşti");
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
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        System.out.println(productCode + " Ürün bilgileri alınamadı.");
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

        actions.scrollByAmount(0, 250).perform();

        WebElement browseAllSeatingLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[5]//li[@class='menu-item-li']"));
        browseAllSeatingLink.click();

        actions.scrollByAmount(0, 450).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));

            if (k >= 2) {
                actions.scrollByAmount(0, 450).perform();
            }

            for (int i = 0; i < productList.size(); i++) {
                System.out.println("i ==== " + i);
                if (i == 0) {
                    // İlk ürün için kaydırma işlemi atlanır
                } else if (i % 3 == 0) {
                    actions.scrollByAmount(0, 570).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
                productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                WebElement product = productList.get(i);
                product.click();

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 750).perform();
                        }

                        WebElement subProduct = subProductListInLoop.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
                            actions.click(subProduct).perform();
                        }

                        if (j % 5 == 0) {
                            System.out.println( j + "if else içinde");
                            actions.scrollByAmount(0, 425).perform();
                            System.out.println("perform gerçekleşti");
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
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        System.out.println(productCode + " Ürün bilgileri alınamadı.");
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
    public void occasional() throws InterruptedException {
        Actions actions = new Actions(driver);
        WebElement occasionalLink = driver.findElement(By.xpath("(//a[@href='/occasional/'])[1]"));
        actions.moveToElement(occasionalLink).perform();
        Thread.sleep(1000);

        WebElement browseAllOccasionalLink = driver.findElement(By.xpath("(//li[@class='main-menu-li categoryImageItem '])[6]//li[@class='menu-item-li']"));
        browseAllOccasionalLink.click();

        actions.scrollByAmount(0, 450).perform();

        // Pagination bilgisi
        String lastPageCountAsString = driver.findElement(By.xpath("(//li[@class='page-item'])[7]")).getText();
        int pageSize = Integer.parseInt(lastPageCountAsString);

        for (int k = 1; k <= pageSize; k++) {
            // Mevcut sayfadaki ürünleri bul
            List<WebElement> productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));

            if (k >= 2) {
                actions.scrollByAmount(0, 450).perform();
            }

            for (int i = 0; i < productList.size(); i++) {
                System.out.println("i ==== " + i);
                if (i == 0) {
                    // İlk ürün için kaydırma işlemi atlanır
                } else if (i % 3 == 0) {
                    actions.scrollByAmount(0, 570).perform();
                } else if (i == 34) {
                    actions.scrollByAmount(0, 200).perform();
                }
                productList = driver.findElements(By.xpath("//div[@class='col position-relative']"));
                WebElement product = productList.get(i);
                product.click();

                String productCode;
                String productPrice;
                String productStockStatus;
                String productRemoteStockStatus;

                boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");
                if (isContainsCollection) {
                    List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));

                    for (int j = 1; j <= subProductList.size(); j++) {
                        // Listeyi yeniden al
                        List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                        if (j == 1) {
                            actions.scrollByAmount(0, 750).perform();
                        }

                        WebElement subProduct = subProductListInLoop.get(j - 1);
                        try {
                            subProduct.click();
                        } catch (Exception e) {
                            actions.click(subProduct).perform();
                        }

                        if (j % 5 == 0) {
                            System.out.println( j + "if else içinde");
                            actions.scrollByAmount(0, 425).perform();
                            System.out.println("perform gerçekleşti");
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
                        productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                        System.out.println(productCode + " Ürün bilgileri alınamadı.");
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

        actions.scrollByAmount(0,375).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,550).perform();
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
                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 750).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
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
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    System.out.println(productCode + " Ürün bilgileri alınamdı.");
                }
            }
            //office sayfsına geri dönüş
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

        actions.scrollByAmount(0,450).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,570).perform();
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
                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 750).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
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
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    System.out.println(productCode + " Ürün bilgileri alınamdı.");
                }
            }
            //accent sayfsına geri dönüş
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

        actions.scrollByAmount(0,450).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,570).perform();
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
                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 750).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
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
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    System.out.println(productCode + " Ürün bilgileri alınamdı.");
                }
            }
            //media sayfsına geri dönüş
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

        actions.scrollByAmount(0,450).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,570).perform();
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
                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 750).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
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
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    System.out.println(productCode + " Ürün bilgileri alınamdı.");
                }
            }
            //lighting sayfsına geri dönüş
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

        actions.scrollByAmount(0,450).perform();

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

            boolean isContainsCollection = driver.findElement(By.xpath("//li[@class='breadcrumb-item active']")).getText().contains("Collection");

            if (isContainsCollection){
                List<WebElement> subProductList = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']" +
                        "//div[@class='mt-3 text-center thumb-item-box thumb-border quick-view-box lh-sm']"));
                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 750).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
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
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
                try {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    productPrice =driver.findElement(By.id("price_block")).getText();
                    productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
                    productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
                    String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
                    allProductsInformation.add(productInfo);
                } catch (Exception e) {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    System.out.println(productCode + " Ürün bilgileri alınamdı.");
                }
            }
            //mattress sayfsına geri dönüş
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

        actions.scrollByAmount(0,450).perform();

        List<WebElement> productList = driver.findElements(By.xpath("//a[@class='flex-fill']"));

        for (int i = 0; i < productList.size(); i++) {
            if (i == 0){
                //buradaki scroll'ü yukarıda kaydık burada boş geçmemiz gerekiyor
            } else if (i % 3 == 0) {
                actions.scrollByAmount(0,570).perform();
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
                for (int j = 1; j <= subProductList.size(); j++) {
                    // Listeyi yeniden al
                    List<WebElement> subProductListInLoop = driver.findElements(By.xpath("//div[@class='row mt-3 row-cols-2 row-cols-md-3 row-cols-lg-4 row-cols-xl-5']//a"));

                    if (j == 1) {
                        actions.scrollByAmount(0, 750).perform();
                    }

                    WebElement subProduct = subProductListInLoop.get(j - 1);
                    try {
                        subProduct.click();
                    } catch (Exception e) {
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
                        actions.sendKeys(Keys.TAB)
                                .sendKeys(Keys.ENTER)
                                .perform();
                    }
                }
            } else {
				try {
					productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
					productPrice =driver.findElement(By.id("price_block")).getText();
					productStockStatus = driver.findElement(By.id("AvailabiltySpan")).getText();
					productRemoteStockStatus = driver.findElement(By.xpath("(//span//b)[4]")).getText();
					String[] productInfo = {productCode, productPrice, productStockStatus, productRemoteStockStatus};
					allProductsInformation.add(productInfo);
				} catch (Exception e) {
                    productCode = driver.findElement(By.xpath("//span[@id='bgitem_name']")).getText();
                    System.out.println(productCode + " Ürün bilgileri alınamdı.");
				}

			}
            //mattress sayfsına geri dönüş
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
        String fileName = "ProductPrices_" + currentDate + ".xlsx";
        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
            workbook.write(outputStream);
        }
        workbook.close();
        System.out.println("Excel dosyası başarıyla kaydedildi!");
    }
}
