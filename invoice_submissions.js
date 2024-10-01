const { Builder, By, until, Key } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const { JSDOM } = require('jsdom');
const fs = require('fs');
const XLSX = require('xlsx');
const { Select } = require('selenium-webdriver');


// Set Chrome options to connect to the existing Chrome session
let chromeOptions = new chrome.Options();
chromeOptions = chromeOptions.addArguments("--remote-debugging-port=9222");  // Use the correct port where Chrome is running
chromeOptions = chromeOptions.debuggerAddress('localhost:9222');  // Connect to the existing Chrome session on this port

let outputData = [];

// Initialize WebDriver and connect to the existing Chrome session
let driver = new Builder().forBrowser('chrome')
    .setChromeOptions(chromeOptions)
    .build();

// Helper function to parse dates in M/D/YYYY format
function parseDate(dateStr) {
    const parts = dateStr.split('/');
    return new Date(parts[2], parts[0] - 1, parts[1]);  // Year, Month (0-indexed), Day
}

function excelDateToJSDate(excelDate) {
    const date = XLSX.SSF.parse_date_code(excelDate);
    if (date) {
        // Create a new Date object using parsed date parts
        const jsDate = new Date(date.y, date.m - 1, date.d);  // Month is 0-indexed in JS
        return `${jsDate.getMonth() + 1}/${jsDate.getDate()}/${jsDate.getFullYear()}`;
    }
    return excelDate;  // Return original value if it's not a date
}

// Main loop to scrape pages and navigate using the "Next" button
async function main() {
    try {
        // Use the existing Chrome instance to navigate to the URL
        await driver.get('https://vendorcentral.amazon.com/hz/vendor/members/invoice-creation/search-shipments?date-range-option=DEFAULT_DATE_RANGE&payee-name-selection=allPayeeCodes&poSearchTable-JSON=%7B%22sortColumnId%22:%22shipped_date%22%7D');

        // Wait until the dropdown is present in the DOM
        let dropdown = await driver.wait(until.elementLocated(By.id('shipment-search-key')), 10000);
        // Select an option from the dropdown
        let selectElement = await new Select(dropdown);
        await selectElement.selectByVisibleText('Purchase Order Number(s)'); // or 'Purchase Order Number(s)'
        console.log('Selected Purchase Order Number(s) from the dropdown.');

        // Wait until the input field is present in the DOM
        let inputField = await driver.wait(until.elementLocated(By.id('po-number')), 10000);
        // Clear the field if necessary
        await inputField.clear();

        const workbook = XLSX.readFile("invoices.xlsx");
        const sheetName = workbook.SheetNames[0];  // Assuming the first sheet is the one you want
        const worksheet = workbook.Sheets[sheetName];

        // Convert sheet to JSON array of arrays (rows)
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });  // Array of arrays (rows)

        // Loop through each row
        var firstRow = 1;
        var invoiceArray = [];  
        rows.forEach((row, rowIndex) => {
            if (firstRow > 1){
                const cellValue = row[0];
                if (!cellValue) {
                    console.log(`Row ${rowIndex + 1}: PO is empty, stopping.`);
                    return;
                }
                invoiceArray.push([row[0], excelDateToJSDate(row[1]), row[2], row[3]])
            }
            firstRow++;
        });

        // Loop through each invoiceArray and send the PO numbers
        for (let i = 0; i < invoiceArray.length; i++) {
            // Wait until the input field is present and re-locate it each time to avoid stale elements
            let inputField = await driver.wait(until.elementLocated(By.id('po-number')), 10000);
            await inputField.clear();  // Clear the field before entering the new value
            await inputField.sendKeys(invoiceArray[i][0]);  // Send the PO number

            // Wait until the "Search" button is present and re-locate it
            let searchButton = await driver.wait(until.elementLocated(By.id('shipmentSearchTableForm-submit')), 10000);
            await searchButton.click();  // Click the "Search" button

            // Wait for the search results to load (adjust the wait time if needed)
            await driver.sleep(500);  // Wait for results to load

            // Now process the table rows and compare dates
            let tableRows = await driver.findElements(By.css('.mt-row'));  // Locate all rows with class "mt-row"
        
            for (let rowIndex = 0; rowIndex < tableRows.length; rowIndex++) {
                // Locate the shipped date for each row
                let shippedDateElement = await driver.wait(until.elementLocated(By.id(`r${rowIndex + 1}-shipped_date`)), 2000);
                let shippedDateText = await shippedDateElement.getText();  // Get the shipped date from the web page

                let shippedDateOnPage = parseDate(shippedDateText);
                //console.log(shippedDateOnPage);
                let shippedDateInInvoice = parseDate(invoiceArray[i][1]);  // Assuming invoiceArray[i][1] is your ship date
                //console.log(shippedDateInInvoice);
                
                // Compare the dates and check if the date on the web page is the same or later than the invoice date
                if (shippedDateOnPage >= shippedDateInInvoice) {
                    // If the condition is met, locate the checkbox and click it
                    
                    let checkbox = await driver.wait(until.elementLocated(By.css(`#r${rowIndex + 1}-asn_checkbox-input-harmonic-checkbox ~ i.a-icon.a-icon-checkbox`)), 10000);
                    await checkbox.click();
                    console.log(`Checkbox for item ${invoiceArray[i][0]} row ${rowIndex + 1} clicked.`);

                    /*
                    SELECTING THE SPECIFIC PO OF ASN
                    */
                    let specificPOButton = await driver.wait(
                        until.elementLocated(By.id('create-inv-asn-po-toggle')),
                        10000  // Wait up to 10 seconds
                    );
                    await driver.executeScript('arguments[0].scrollIntoView(true);', specificPOButton);
                    await specificPOButton.click();
                    let ASNcheckbox = await driver.wait(
                        until.elementLocated(By.css('input[type="checkbox"][data-asn-check="true"]')),
                        10000  // Wait up to 10 seconds
                    );
                    let isChecked = await ASNcheckbox.isSelected();
                    if (isChecked) {
                        await ASNcheckbox.click();
                        console.log(`${invoiceArray[i][0]}`);
                        let POcheckbox = await driver.wait(
                            until.elementLocated(By.css(`input[type="checkbox"][data-po-check="true"][value="${invoiceArray[i][0]}"]`)),
                            10000  // Wait up to 10 seconds
                        );
                        await POcheckbox.click();
                        console.log('Checkbox deselected.');
                            
                    } else {
                        console.log('Checkbox is already deselected.');
                    }


                    /*
                    SELECTING THE SPECIFIC PO OF ASN
                    */

                    let submitButton = await driver.wait(
                        until.elementLocated(By.css('input.a-button-input[aria-labelledby="create-invoice-submit-announce"]')),
                        10000  // Wait up to 10 seconds
                    );
                    await submitButton.click();

                    let totalAmountElement = await driver.wait(
                        until.elementLocated(By.id('inv-total-amount-data')),
                        10000  // Wait up to 10 seconds
                    );
                    let totalAmountText = await totalAmountElement.getText();
                    let totalAmount = parseFloat(totalAmountText.replace(/[$,]/g, ''));
                    console.log(totalAmount);
                    let invoiceAmount = parseFloat(invoiceArray[i][3]);
                    console.log(invoiceAmount);
                    
                    // Define an epsilon (tolerance for floating-point comparison)
                    let epsilon = 0.001;
                    
                    // Compare the two values with the epsilon tolerance
                    if (Math.abs(totalAmount - invoiceAmount) < epsilon) {
                        console.log("USD Match")
                        let invoiceNumberInput = await driver.wait(
                            until.elementLocated(By.id('invoice-number')),
                            500
                        );
                        await invoiceNumberInput.sendKeys(invoiceArray[i][2]);
                        
                        
                        let checkboxIcon = await driver.wait(
                            until.elementIsEnabled(driver.findElement(By.css('#inv-agree-checkbox ~ i.a-icon.a-icon-checkbox'))),
                            10000  // Wait up to 10 seconds
                        );
                        // Scroll the checkbox into view and click it
                        await driver.executeScript("arguments[0].scrollIntoView(true);", checkboxIcon);
                        await driver.executeScript("arguments[0].click();", checkboxIcon);  // Use JavaScript to click

                        let submitButton = await driver.wait(
                            until.elementLocated(By.css('input.a-button-input[aria-labelledby="inv-submit-announce"]')),
                            10000  // Wait up to 10 seconds
                        );
                        await driver.executeScript("arguments[0].scrollIntoView(true);", submitButton);
                        await driver.executeScript("arguments[0].click();", submitButton);
                        
                        /*
                        PROCEED TO NEXT INVOICE
                        */
                        await driver.sleep(1000);  // Wait for results to load
                        let redirectButton = await driver.wait(
                            until.elementLocated(By.css('input.a-button-input[aria-labelledby="inv-crt-redirect-announce"]')),
                            10000  // Wait up to 10 seconds
                        );
                        await driver.executeScript("arguments[0].click();", redirectButton);
                        outputData.push([invoiceArray[i][0], invoiceArray[i][2], invoiceAmount, invoiceAmount, "Submitted"]);
                    } else if (totalAmount == 0){
                        outputData.push([invoiceArray[i][0], invoiceArray[i][2], invoiceAmount, invoiceAmount, "Submitted"]);
                        console.log("Already Submitted")
                        await driver.get('https://vendorcentral.amazon.com/hz/vendor/members/invoice-creation/search-shipments?date-range-option=DEFAULT_DATE_RANGE&payee-name-selection=allPayeeCodes&poSearchTable-JSON=%7B%22sortColumnId%22:%22shipped_date%22%7D');
                        let dropdown = await driver.wait(until.elementLocated(By.id('shipment-search-key')), 10000);
                        // Select an option from the dropdown
                        let selectElement = await new Select(dropdown);
                        await selectElement.selectByVisibleText('Purchase Order Number(s)'); // or 'Purchase Order Number(s)'
                        continue;
                    } else{
                        console.log("ERROR MATCHING USD")
                        await driver.get('https://vendorcentral.amazon.com/hz/vendor/members/invoice-creation/search-shipments?date-range-option=DEFAULT_DATE_RANGE&payee-name-selection=allPayeeCodes&poSearchTable-JSON=%7B%22sortColumnId%22:%22shipped_date%22%7D');
                        // Wait until the dropdown is present in the DOM
                        let dropdown = await driver.wait(until.elementLocated(By.id('shipment-search-key')), 10000);
                        // Select an option from the dropdown
                        let selectElement = await new Select(dropdown);
                        await selectElement.selectByVisibleText('Purchase Order Number(s)'); // or 'Purchase Order Number(s)'
                        outputData.push([invoiceArray[i][0], invoiceArray[i][2], invoiceAmount, totalAmount, "Price error"]);
                        continue;
                    }

                } else {
                    console.log(`Checkbox for item ${invoiceArray[i][0]} row ${rowIndex + 1} not clicked (date condition not met).`);
                    outputData.push([invoiceArray[i][0], invoiceArray[i][3], "Date condition not met"]);
                }
            }
        }

    } finally {
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.aoa_to_sheet(outputData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
        XLSX.writeFile(newWorkbook, 'invoices_status.xlsx');

        console.log('Invoice statuses have been written to invoices_status.xlsx');
        // Optionally, you can leave the Chrome session open
        // driver.quit();  // Comment this out if you want to keep Chrome open
        /*

        NOTES:
        inv-total-amount-data must equal invoiceArray[3]

        Create another invoice button:
        <div class="a-column a-span2"><span id="inv-crt-redirect" class="a-button a-button-groupfirst a-button-primary"><span class="a-button-inner"><input class="a-button-input" type="submit" aria-labelledby="inv-crt-redirect-announce"><span id="inv-crt-redirect-announce" class="a-button-text" aria-hidden="true">Create another invoice</span></span></span></div>
        */
    
    }
}

main();
