# Indian Stock Price Automation

This project automates the process of fetching stock prices of Indian companies from multiple online sources (NSE India, MoneyControl, BSE India) and exports the results into an Excel file.

## Features

Reads stock symbols from an input Excel file.

Fetches real-time stock prices using Selenium WebDriver.

Attempts multiple data sources in order: NSE India → MoneyControl → BSE India.

Exports results (Stock Symbol, Price, Timestamp) into a new Excel file.

Handles missing/invalid stock data gracefully.

## Tech Stack

Java (Core logic)

Selenium WebDriver (Web scraping)

Apache POI (Excel handling)

WebDriverManager (Driver management)

Firefox (GeckoDriver) (Browser automation)

## Output

Results are saved in stock_prices_output.xlsx with:

Stock Symbol

Current Price

Timestamp

### 🚀 Use this script to quickly check multiple stock prices and maintain updated records.
