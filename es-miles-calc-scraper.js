const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());
const XLSX = require('xlsx');

// Function to handle cookies window
async function handleCookiesPopup(page) {
    const cookiesButton = await page.$('#onetrust-accept-btn-handler');
    if (cookiesButton) {
        await cookiesButton.click();
        console.log('Clicked the cookies accept button...');
    }
}

// Function to enter data into input fields
async function enterDataIntoCombobox(page, dataTestId, inputData, maxRetries = 3) {
    const comboboxSelector = `div[data-testid="${dataTestId}"]`;
    const inputSelector = `${comboboxSelector} input.input-field__input`;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            // Wait for the combobox and click it
            await page.locator(inputSelector).click({ clickCount: 3 });
            await page.keyboard.press('Backspace');

            // Type the input data
            await page.type(inputSelector, inputData);
            await delay(500);
            await page.keyboard.press('Enter');

            // Wait and validate if the input text is contained within the value
            // await delay(1000);
            const enteredText = await page.$eval(inputSelector, el => el.value);

            if (enteredText.includes(inputData)) {
                console.log(`Success: ${dataTestId}: ${inputData}`);
                return; // Exit the function if the text is successfully entered
            } else {
                console.warn(`Validation failed on attempt ${attempt}: ${enteredText} does not include ${inputData}`);
            }
        } catch (error) {
            console.error(`Attempt ${attempt} failed: ${error.message}`);
        }

        if (attempt < maxRetries) {
            console.log(`Retrying (${attempt + 1}/${maxRetries})...`);
            // await delay(1000); // Optional: wait before retrying
        }
    }

    throw new Error(`Failed to enter ${inputData} into ${dataTestId} after ${maxRetries} attempts`);
}



// Function to select an option from a dropdown menu
// async function selectComboboxOption(page, comboboxTestId, optionText) {
//     await page.click(`div[data-testid="${comboboxTestId}"]`);
//     await page.waitForSelector('button.auto-suggest__item', { visible: true });
//     await page.evaluate((text) => {
//         const options = Array.from(document.querySelectorAll('button.auto-suggest__item'));
//         const option = options.find(el => el.textContent.trim() === text);
//         if (option) {
//             option.click();
//         }
//     }, optionText);
//     await delay(1000);
// }
async function selectComboboxOption(page, comboboxTestId, optionText) {
    await page.click(`div[data-testid="${comboboxTestId}"]`);
    await page.waitForSelector('button.auto-suggest__item', { visible: true });
    await page.evaluate((text) => {
        const options = Array.from(document.querySelectorAll('button.auto-suggest__item'));
        const cleanedText = text.trim();
        const option = options.find(el => {
            let optionText = el.textContent.trim();
            if (optionText.includes('Class')) {
                optionText = optionText.replace(/Class/g, '').replace(/\s+/g, '');
            }
            return optionText === cleanedText;
        });
        if (option) {
            option.click();
        }
    }, optionText);
    await delay(1000);
}


function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
}

// Function to read and validate Excel data returns Json
function readExcelData(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    const requiredColumns = ['flyingWith', 'leavingFrom', 'goingTo', 'cabinClass', 'emiratesSkywardsTier', 'oneWayOrRoundtrip'];
    const columnNames = Object.keys(jsonData[0]);

    // Check if required columns are present
    const missingColumns = requiredColumns.filter(col => !columnNames.includes(col));
    if (missingColumns.length > 0) {
        throw new Error(`Missing columns in Excel file: ${missingColumns.join(', ')}`);
    }

    return jsonData;
}


(async () => {
    const browser = await puppeteer.launch({
        headless: false,
        // devtools: true,
        defaultViewport: null,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-accelerated-2d-canvas',
            '--disable-gpu',
            '--window-size=1920x1080',
            '--disable-background-timer-throttling',
            '--disable-backgrounding-occluded-windows',
            '--disable-renderer-backgrounding',
            '--disable-web-security',
            '--disable-features=IsolateOrigins',
            '--disable-site-isolation-trials',
            '--disable-features=BlockInsecurePrivateNetworkRequests'
        ],
    });

    try {
        const excelFilePath = 'inputData.xlsx'; // Replace with your actual Excel file path
        const excelData = readExcelData(excelFilePath);
        const page = await browser.newPage();
        await page.setDefaultNavigationTimeout(60000);
        const homeUrl = 'https://www.emirates.com/ph/english/skywards/miles-calculator/';
        await page.goto(homeUrl);
        await page.waitForSelector('.skywards-miles-calculator__search-widget__wrapper', { visible: true });
        console.log('Navigated to Emirates mileage calculator...');
        await delay(250);
        await handleCookiesPopup(page);

        let allData = []; // Array to store all results

        for (const rowData of excelData) {
            const flyingWith = rowData.flyingWith;
            const leavingFrom = rowData.leavingFrom;
            const goingTo = rowData.goingTo;
            const cabinClass = rowData.cabinClass;
            const emiratesSkywardsTier = rowData.emiratesSkywardsTier;
            const oneWayOrRoundtrip = rowData.oneWayOrRoundtrip;

            // Click the appropriate radio button based on the oneWayOrRoundtrip variable
            if (oneWayOrRoundtrip === "One Way") {
                await page.click('input.radio-button__input#OW0');
                await delay(500);
                console.log('Clicked One Way...');
            } else if (oneWayOrRoundtrip === "Round Trip") {
                await page.click('input.radio-button__input#RT1');
                await delay(500);
                console.log('Clicked Round Trip...');
            }

            await selectComboboxOption(page, "combobox_Flying with", flyingWith);
            await enterDataIntoCombobox(page, "combobox_Leaving from", leavingFrom);
            await enterDataIntoCombobox(page, "combobox_Going to", goingTo);
            await selectComboboxOption(page, "combobox_Cabin class", cabinClass);
            await selectComboboxOption(page, "combobox_Emirates Skywards tier", emiratesSkywardsTier);
            await page.click('button[aria-label="Calculate"]');
            console.log('Submitted the form...');
            console.log('Waiting for results...');

            const waitForSelectors = Promise.race([
                page.waitForSelector('.tabs', { timeout: 90000 }),
                page.waitForSelector('h1:not([class]):not([id])', { timeout: 90000 })
            ]);

            await waitForSelectors;

            let flattenedData = [];

            const isResults = await page.$('.tabs');
            const isAccessDenied = await page.$('h1:not([class]):not([id])');

            // Wait for the results page to load
            await page.waitForSelector('.skywards-miles-calculator__search-result.miles-calculator-result__section', { visible: true });

            // Data Scraping
            if (isResults) {
                const pageCardData = await page.$$eval('.tabs', (cards, cabinClass) => {
                    return cards.map(card => {
                        const actionElement = card.querySelector('a[aria-selected="true"]');
                        const action = actionElement ? actionElement.querySelector('.miles-calculator-result__tab-button--text').textContent.trim() : null;

                        const fareDetails = Array.from(document.querySelectorAll('.miles-calculator-result__card-wrapper')).flatMap(wrapper => {
                            const fareTypes = ['special', 'saver', 'flex', 'flexplus'];
                            return fareTypes.map(fareType => {
                                // Here we strictly use input data for cabin class i.e. premiumeconomy must have no space thus the regex, input must not have Class words in excel, it was handled in the input for Cabin Class in the form page
                                const fareDiv = wrapper.querySelector(`.miles-card__card.miles-card__ek-${cabinClass.toLowerCase().replace(/\s+/g, '')}-${fareType}`);
                                if (fareDiv) {
                                    const milesContent = fareDiv.querySelector('.miles-card__content__miles');
                                    let skywardMiles = 'N/A';
                                    let tierMiles = 'N/A';

                                    // Extract Skywards Miles
                                    const skywardsMilesTitle = milesContent.querySelector('.miles-card__skywards-title');
                                    if (skywardsMilesTitle && skywardsMilesTitle.textContent.trim() === 'Skywards Miles') {
                                        const nextSkywardsMilesDiv = skywardsMilesTitle.nextElementSibling;
                                        if (nextSkywardsMilesDiv && nextSkywardsMilesDiv.classList.contains('miles-card__skywards-miles')) {
                                            const skywardsMilesSpan = nextSkywardsMilesDiv.querySelector('span');
                                            if (skywardsMilesSpan) {
                                                skywardMiles = skywardsMilesSpan.textContent.trim();
                                            }
                                        }
                                    }

                                    // Extract Tier Miles
                                    const tierMilesDiv = milesContent.querySelector('.miles-card__tier-miles .miles-card__skywards-miles span');
                                    if (tierMilesDiv) {
                                        tierMiles = tierMilesDiv.textContent.trim();
                                    }

                                    return {
                                        brandedFare: fareType.charAt(0).toUpperCase() + fareType.slice(1),
                                        skywardMiles,
                                        tierMiles
                                    };
                                } else {
                                    return {
                                        brandedFare: fareType.charAt(0).toUpperCase() + fareType.slice(1),
                                        skywardMiles: 'N/A',
                                        tierMiles: 'N/A'
                                    };
                                }
                            });
                        });

                        return {
                            action,
                            fareDetails
                        };
                    });
                }, cabinClass);

                if (pageCardData) {
                    pageCardData.forEach(card => {
                        card.fareDetails.forEach(fareDetail => {
                            flattenedData.push({
                                action: card.action,
                                flyingWith: flyingWith,
                                leavingFrom: leavingFrom,
                                goingTo: goingTo,
                                date: oneWayOrRoundtrip,
                                cabinClass: cabinClass,
                                brandedFare: fareDetail.brandedFare,
                                skywardTier: emiratesSkywardsTier,
                                skywardMiles: fareDetail.skywardMiles,
                                tierMiles: fareDetail.tierMiles
                            });
                        });
                    });
                }
            } else if (isAccessDenied) {
                console.log('Access denied');
                flattenedData.push({
                    action: "Access Denied",
                    leavingFrom: leavingFrom,
                    goingTo: goingTo,
                    brandedFare: 'N/A',
                    skywardMiles: 'N/A',
                    tierMiles: 'N/A'
                });
            }

            // Accumulate data for each row
            allData = allData.concat(flattenedData);
            await page.locator('a[data-id="pagebody_link"][data-link="Back to results"]').click();
        }

        // Write all accumulated data to a single Excel file

        if (allData.length > 0) {
            const workbook = XLSX.utils.book_new();

            const actions = [...new Set(allData.map(item => item.action))];

            actions.forEach(action => {
                const dataForSheet = allData.filter(item => item.action === action);
                const sheetData = dataForSheet.map(item => ({
                    'Action': item.action,
                    'Flying With': item.flyingWith,
                    'Leaving from': item.leavingFrom,
                    'Going to': item.goingTo,
                    'Date (OW/RT)': item.date,
                    'Cabin Class': item.cabinClass,
                    'Emirates Skywards Tier': item.skywardTier,
                    'Branded Fare': item.brandedFare,
                    'Skyward Miles': item.skywardMiles,
                    'Tier Miles': item.tierMiles
                }));
                const worksheet = XLSX.utils.json_to_sheet(sheetData);
                XLSX.utils.book_append_sheet(workbook, worksheet, action);
            });

            const fileName = `skywardsMilesData.xlsx`;
            XLSX.writeFile(workbook, fileName);
            console.log(`Excel file written to ${fileName}`);
            
            
        }
    }
     catch (error) {
        console.error('Error:', error);
        await browser.close();
    } finally {
        await browser.close();
    }
})();
