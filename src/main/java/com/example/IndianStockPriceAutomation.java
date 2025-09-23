package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
//import org.openqa.selenium.firefox.FirefoxDriver;
//import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import io.github.bonigarcia.wdm.WebDriverManager;
import java.io.*;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.nio.file.Paths;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class IndianStockPriceAutomation {

    private static final String CSV_FILE_PATH = "src/main/resources/stocks_list.csv";
    private static final String OUTPUT_FILE_PATH = "stock_prices_output.xlsx";
    private static final int TIMEOUT_SECONDS = 15;

    public static void main(String[] args) {
        IndianStockPriceAutomation automation = new IndianStockPriceAutomation();
        automation.runStockPriceScript();
    }

    public void runStockPriceScript() {
        List<String> stockSymbols = readStockSymbolsFromCSV();
        if (stockSymbols.isEmpty()) {
            System.out.println("No stock symbols found in CSV file!");
            return;
        }

        WebDriver driver = setupWebDriver();
        List<StockData> stockDataList = new ArrayList<>();

        try {
            for (String symbol : stockSymbols) {
                try {
                    StockData stockData = fetchStockPrice(driver, symbol);
                    stockDataList.add(stockData);
                    System.out.println("Fetched data for: " + symbol + " - Price: ₹" + stockData.price);
                    Thread.sleep(3000); // Wait between requests
                } catch (Exception e) {
                    System.err.println("Error fetching data for " + symbol + ": " + e.getMessage());
                    stockDataList.add(new StockData(symbol, "Error", getCurrentTimestamp()));
                }
            }

            writeStockDataToExcel(stockDataList);
            System.out.println("\nStock price data has been written to: " + OUTPUT_FILE_PATH);

        } catch (Exception e) {
            System.err.println("An error occurred: " + e.getMessage());
        } finally {
            driver.quit();
        }
    }

    // NEW CSV READER METHOD
    private List<String> readStockSymbolsFromCSV() {
        List<String> stockSymbols = new ArrayList<>();
        
        try (BufferedReader br = new BufferedReader(new FileReader(CSV_FILE_PATH))) {
            String line;
            boolean firstLine = true;
            
            while ((line = br.readLine()) != null) {
                line = line.trim();
                
                if (firstLine) {
                    // Skip header if it contains "Symbol"
                    if (line.toLowerCase().contains("symbol")) {
                        firstLine = false;
                        continue;
                    }
                    firstLine = false;
                }
                
                if (!line.isEmpty()) {
                    stockSymbols.add(line);
                }
            }
            
        } catch (IOException e) {
            System.err.println("Error reading CSV file: " + e.getMessage());
            System.out.println("Please ensure the CSV file '" + CSV_FILE_PATH + "' exists with stock symbols.");
        }
        
        return stockSymbols;
    }

    private WebDriver setupWebDriver() {
    WebDriverManager.chromedriver().setup();
    
    ChromeOptions options = new ChromeOptions();
    
    // Chrome headless options
    options.addArguments("--headless=new");
    options.addArguments("--no-sandbox");
    options.addArguments("--disable-dev-shm-usage");
    options.addArguments("--disable-gpu");
    options.addArguments("--window-size=1920,1080");
    options.addArguments("--disable-extensions");
    options.addArguments("--disable-web-security");
    options.addArguments("--allow-running-insecure-content");
    
    // Set user agent
    options.addArguments("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
    
    return new ChromeDriver(options);

   /* private WebDriver setupWebDriver() {
        WebDriverManager.firefoxdriver().setup();

        FirefoxOptions options = new FirefoxOptions();

        // Firefox performance optimizations
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("--disable-gpu");
        options.addArguments("--width=1920");
        options.addArguments("--height=1080");
        options.addArguments("--headless");

        // Set user agent
        options.addPreference("general.useragent.override",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/109.0");

        // Disable images to improve performance
        options.addPreference("permissions.default.image", 2);

        // Disable Flash
        options.addPreference("dom.ipc.plugins.enabled.libflashplayer.so", false);

        // Add timeout preferences
        options.addPreference("network.http.connection-timeout", 60);
        options.addPreference("network.http.response.timeout", 60);

        return new FirefoxDriver(options);
       */
    }

    // PRIMARY: NSE India as main data source - SIMPLIFIED to get only price
    private StockData fetchStockPrice(WebDriver driver, String symbol) {
        try {
            String url = "https://www.nseindia.com/get-quotes/equity?symbol=" + symbol.toUpperCase();
            System.out.println("Fetching from NSE India: " + url);

            driver.get(url);
            Thread.sleep(5000); // Wait for page to load

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(TIMEOUT_SECONDS));

            // NSE India selectors - try multiple options for price only
            String[] nseSelectors = {
                    "#quoteLtp",
                    ".trading_price",
                    ".overview-eq .equity-price",
                    "span[id*='ltp']",
                    ".equity-ltp",
                    "#priceInfoData span[id*='ltp']"
            };

            WebElement priceElement = null;
            for (String selector : nseSelectors) {
                try {
                    priceElement = wait.until(
                            ExpectedConditions.presenceOfElementLocated(By.cssSelector(selector))
                    );
                    if (priceElement != null && !priceElement.getText().trim().isEmpty()) {
                        System.out.println("NSE selector worked: " + selector);
                        break;
                    }
                } catch (Exception e) {
                    System.out.println("NSE selector failed: " + selector);
                }
            }

            if (priceElement != null) {
                String price = priceElement.getText().trim();
                System.out.println("NSE India price found: " + price);

                // SIMPLIFIED: Return only symbol, price, and timestamp
                return new StockData(symbol, price, getCurrentTimestamp());
            }

        } catch (Exception e) {
            System.err.println("NSE India failed for " + symbol + ": " + e.getMessage());
        }

        // Fallback to MoneyControl only for price
        return fetchFromMoneyControl(driver, symbol);
    }

    // FALLBACK: MoneyControl - SIMPLIFIED to get only price
    private StockData fetchFromMoneyControl(WebDriver driver, String symbol) {
        try {
            String url = "https://www.moneycontrol.com/india/stockpricequote/" + symbol.toLowerCase();
            System.out.println("Fallback to MoneyControl: " + url);

            driver.get(url);
            Thread.sleep(5000);

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(TIMEOUT_SECONDS));

            // MoneyControl selectors for price only
            String[] mcSelectors = {
                    "#Bse_Prc_tick .span_price_wrap",
                    "#nsecp",
                    ".inprice1",
                    ".price_overview .span_price_wrap",
                    "div[id*='price'] .span_price_wrap",
                    ".overview .inprice",
                    ".stockprc"
            };

            WebElement priceElement = null;
            for (String selector : mcSelectors) {
                try {
                    priceElement = wait.until(
                            ExpectedConditions.presenceOfElementLocated(By.cssSelector(selector))
                    );
                    if (priceElement != null && !priceElement.getText().trim().isEmpty()) {
                        System.out.println("MoneyControl selector worked: " + selector);
                        break;
                    }
                } catch (Exception e) {
                    System.out.println("MoneyControl selector failed: " + selector);
                }
            }

            if (priceElement != null) {
                String price = priceElement.getText().trim();
                System.out.println("MoneyControl price found: " + price);

                // SIMPLIFIED: Return only symbol, price, and timestamp
                return new StockData(symbol, price, getCurrentTimestamp());
            }

        } catch (Exception e) {
            System.err.println("MoneyControl also failed for " + symbol + ": " + e.getMessage());
        }

        // FINAL FALLBACK: Try BSE India for price only
        return fetchFromBSEIndia(driver, symbol);
    }

    // BSE India fallback - SIMPLIFIED to get only price
    private StockData fetchFromBSEIndia(WebDriver driver, String symbol) {
        try {
            String url = "https://www.bseindia.com/stock-share-price/" + symbol.toLowerCase() + "/";
            System.out.println("Final fallback to BSE India: " + url);

            driver.get(url);
            Thread.sleep(5000);

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(TIMEOUT_SECONDS));

            // BSE India selectors for price only
            String[] bseSelectors = {
                    ".curr-price",
                    ".stock-price",
                    ".price-current",
                    "span[id*='price']"
            };

            WebElement priceElement = null;
            for (String selector : bseSelectors) {
                try {
                    priceElement = wait.until(
                            ExpectedConditions.presenceOfElementLocated(By.cssSelector(selector))
                    );
                    if (priceElement != null && !priceElement.getText().trim().isEmpty()) {
                        System.out.println("BSE selector worked: " + selector);
                        break;
                    }
                } catch (Exception ignored) {
                    System.out.println("BSE selector failed: " + selector);
                }
            }

            if (priceElement != null) {
                String price = priceElement.getText().trim();
                System.out.println("BSE India price found: " + price);

                // SIMPLIFIED: Return only symbol, price, and timestamp
                return new StockData(symbol, price, getCurrentTimestamp());
            }

        } catch (Exception e) {
            System.err.println("All sources failed for " + symbol + ": " + e.getMessage());
        }

        return new StockData(symbol, "Error", getCurrentTimestamp());
    }

    // SIMPLIFIED: Excel output with only 3 columns
    private void writeStockDataToExcel(List<StockData> stockDataList) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Stock Prices");

            // Create header row - ONLY 3 columns
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Stock Symbol");
            headerRow.createCell(1).setCellValue("Current Price");
            headerRow.createCell(2).setCellValue("Timestamp");

            // Create data rows - ONLY 3 columns
            for (int i = 0; i < stockDataList.size(); i++) {
                Row row = sheet.createRow(i + 1);
                StockData data = stockDataList.get(i);

                row.createCell(0).setCellValue(data.symbol);
                row.createCell(1).setCellValue(data.price);
                row.createCell(2).setCellValue(data.timestamp);
            }

            // Auto-size columns - ONLY 3 columns
            for (int i = 0; i < 3; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write to file
            try (FileOutputStream fos = new FileOutputStream(OUTPUT_FILE_PATH)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            System.err.println("Error writing to Excel file: " + e.getMessage());
        }
    }

    private String getCurrentTimestamp() {
        return LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
    }

    // SIMPLIFIED: StockData class with only 3 fields
    static class StockData {
        String symbol;
        String price;
        String timestamp;

        public StockData(String symbol, String price, String timestamp) {
            this.symbol = symbol;
            this.price = price;
            this.timestamp = timestamp;
        }
    }
}

