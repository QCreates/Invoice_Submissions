const { Builder, By, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const { JSDOM } = require('jsdom');
const fs = require('fs');
const XLSX = require('xlsx');
const { time } = require('console');

// Set Chrome options to connect to the existing Chrome session
let chromeOptions = new chrome.Options();
chromeOptions = chromeOptions.addArguments("--remote-debugging-port=9222");  // Use the correct port where Chrome is running
chromeOptions = chromeOptions.debuggerAddress('localhost:9222');  // Connect to the existing Chrome session on this port

// Initialize WebDriver and connect to the existing Chrome session
let driver = new Builder().forBrowser('chrome')
    .setChromeOptions(chromeOptions)
    .build();

// Main loop to scrape pages and navigate using the "Next" button
async function main() {
    try {
        
        // Use the existing Chrome instance to navigate to the URL
        await driver.get('https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/shippingqueue');
    } finally {
        // Optionally, you can leave the Chrome session open
        // driver.quit();  // Comment this out if you want to keep Chrome open
    }
}

main();
