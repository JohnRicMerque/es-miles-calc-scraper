const puppeteer = require('puppeteer');
const XLSX = require('xlsx');

// Function to handle cookies window
async function handleCookiesPopup(page) {
    const cookiesButton = await page.$('#onetrust-accept-btn-handler');
    if (cookiesButton) {
        await cookiesButton.click();
        console.log('Clicked the cookies accept button...');
    }
}

async function enterDataIntoCombobox(page, dataTestId, inputData) {
    // Select the element using data-testid
    const comboboxSelector = `div[data-testid="${dataTestId}"]`;

    // Wait for the element to appear on the page
    await page.waitForSelector(comboboxSelector);

    // Click to open the combobox
    await page.click(comboboxSelector);
    const inputSelector = `${comboboxSelector} input.input-field__input`;

    // Focus and type into the input
    await page.type(inputSelector, inputData);
    await delay(500);
    await page.keyboard.press('Enter');
    console.log(`${dataTestId}: ${inputData}`);
    await delay(1000);
}

// Function to select an option from a combobox
async function selectComboboxOption(page, comboboxTestId, optionText) {
    // Trigger the dropdown to open
    await page.click(`div[data-testid="${comboboxTestId}"]`);

    // Wait for the options to appear
    await page.waitForSelector('button.auto-suggest__item', { visible: true });

    // Select the desired option based on the optionText
    await page.evaluate((text) => {
        // Find the option with the text specified by optionText and click it
        const options = Array.from(document.querySelectorAll('button.auto-suggest__item'));
        const option = options.find(el => el.textContent.trim() === text);
        if (option) {
            option.click();
        }
    }, optionText);

    await delay(1000);
}

// Function to handle delays
function delay(time) {
    return new Promise(function (resolve) {
        setTimeout(resolve, time);
    });
}

(async () => {
    const browser = await puppeteer.launch({
        headless: false,
        defaultViewport: null,
    });

    // Open a new page
    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(60000);

    // Variables for Input 
    const flyingWith = 'Emirates';
    const leavingFrom = 'ABJ';
    const goingTo = 'ABV';
    const cabinClass = 'Economy';
    const emiratesSkywardsTier = 'Blue'


    // Navigate to Emirates base fare page
    const homeUrl = 'https://www.emirates.com/ph/english/skywards/miles-calculator/';
    await page.goto(homeUrl);
    console.log('Navigated to Emirates mileage calculator...');

    await delay(1000);

    await handleCookiesPopup(page);

    // Click Oneway
    await page.click('input.radio-button__input#OW0');
    await delay(500);
    console.log('Clicked Oneway...');

    await selectComboboxOption(page, "combobox_Flying with", flyingWith);
    await enterDataIntoCombobox(page, "combobox_Leaving from", leavingFrom);
    await enterDataIntoCombobox(page, "combobox_Going to", goingTo);
    await selectComboboxOption(page, "combobox_Cabin class", cabinClass);
    await selectComboboxOption(page, "combobox_Emirates Skywards tier", emiratesSkywardsTier);

    await page.click('button[aria-label="Calculate"]');
    console.log('Submitted the form...');
    console.log('Waiting for results...');



    // NEXT PAGE
    const waitForSelectors = Promise.race([
        page.waitForSelector('.tabs', { timeout: 90000 }),
        page.waitForSelector('h1:not([class]):not([id])', { timeout: 90000 })
    ]);

    await waitForSelectors;

    let flattenedData = [];

    const isResults = await page.$('.tabs');
    const isAccessDenied = await page.$('h1:not([class]):not([id])');

    if (isResults) {
        const pageCardData = await page.$$eval('.tabs', (cards) => {
            return cards.map(card => {
                // Access the selected tab directly in the browser context
                const selectedTab = card.querySelector('a[aria-selected="true"]');
                const action = selectedTab ? selectedTab.querySelector('.miles-calculator-result__tab-button--text').textContent.trim() : null;
                return { action };
            });
        });
    
        // Ensure pageCardData is defined and has the expected structure
        if (pageCardData) {
            flattenedData = pageCardData.map(card => ({
                action: card.action,
            }));
        }
    }
    
    
    else if (isAccessDenied) {
        const errorData = {
            action: "Access Denied"
        };

        flattenedData.push(errorData);
    }

    if (flattenedData.length > 0) {
        const worksheet = XLSX.utils.json_to_sheet(flattenedData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Flight Data');
        XLSX.writeFile(workbook, 'flightData.xlsx');
        console.log('Data successfully written to flightData.xlsx');
    } else {
        console.log('No data found to write.');
    }
    






















    // const comboboxSelector = 'div[data-testid="combobox_Going to"]';

    // // Wait for the element to appear on the page
    // await page.waitForSelector(comboboxSelector);

    // // Click to open the combobox
    // await page.click(comboboxSelector);
    // console.log('Clicked the combobox...');
    // const inputSelector = `${comboboxSelector} input.input-field__input`;

    // await page.click(inputSelector)
    // await page.keyboard.down('Control');
    // await page.keyboard.press('A');
    // await page.keyboard.up('Control');
    // await page.keyboard.press('Backspace');
    // await page.type(inputSelector, leavingFrom);
    // await delay(500);
    // await page.keyboard.press('Enter');
    // console.log('Flying with: ' + leavingFrom);

    // await delay(1000);







    // await page.keyboard.down('ArrowDown');
    // await page.keyboard.press('Enter');
    // await delay(500);

    // await page.keyboard.down('Control');
    // await page.keyboard.press('A');
    // await page.keyboard.up('Control');
    // await page.keyboard.press('Backspace');
    // await page.type('.js-origin-dropdown input', flyingWith);
    // await delay(500);
    // await page.keyboard.press('Enter');
    // console.log('Flying with: ' + flyingWith);

    // await delay(1000);

    // // leaving from input
    // await page.click('.js-origin-dropdown input');
    // await delay(500);
    // await page.keyboard.down('Control');
    // await page.keyboard.press('A');
    // await page.keyboard.up('Control');
    // await page.keyboard.press('Backspace');
    // await page.type('.js-origin-dropdown input', leavingFrom);
    // await delay(500);
    // await page.keyboard.press('Enter');
    // console.log('Leaving from: ' + leavingFrom);

    // await delay(1000);

    // Select the parent div
    // const parentDivSelector = '.skywards-miles-calculator__search-widget__wrapper';
    // const dropdownSelectors = await page.$$(`${parentDivSelector} select`);

    // const dropdownValues = ['Option1', 'Option2', 'Option3']; // Replace with the actual values you want to select

    // // Loop through each dropdown and select a value
    // for (let i = 0; i < dropdownSelectors.length; i++) {
    //     const dropdown = dropdownSelectors[i];
    //     const valueToSelect = dropdownValues[i]; // Assume each dropdown has a corresponding value in dropdownValues

    //     await dropdown.select(valueToSelect);
    //     console.log(`Selected ${valueToSelect} in dropdown ${i + 1}`);
    //     await delay(500); // Add a slight delay between selections if necessary
    // }

    await browser.close();

})();