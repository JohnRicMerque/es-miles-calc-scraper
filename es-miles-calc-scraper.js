const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());
const XLSX = require('xlsx');
const anonymizeUaPlugin = require('puppeteer-extra-plugin-anonymize-ua');
puppeteer.use(anonymizeUaPlugin());

// Function to handle cookies window
async function handleCookiesPopup(page) {
    const cookiesButton = await page.$('#onetrust-accept-btn-handler');
    if (cookiesButton) {
        await cookiesButton.click();
        console.log('Clicked the cookies accept button...');
    }
}

// Function to enter data into input fields
// async function enterDataIntoCombobox(page, dataTestId, inputData, maxRetries = 3) {
//     const comboboxSelector = `div[data-testid="${dataTestId}"]`;
//     const inputSelector = `${comboboxSelector} input.input-field__input`;

//     for (let attempt = 1; attempt <= maxRetries; attempt++) {
//         try {
//             // Wait for the combobox and click it
//             await page.locator(inputSelector).click({ clickCount: 2 });
//             await page.keyboard.press('Backspace');

//             // Type the input data
//             await page.type(inputSelector, inputData);
//             await delay(500);
//             await page.keyboard.press('Enter');

//             // Wait and validate if the input text is contained within the value
//             // await delay(1000);
//             const enteredText = await page.$eval(inputSelector, el => el.value);

//             if (enteredText.includes(inputData)) {
//                 console.log(`Success: ${dataTestId}: ${inputData}`);
//                 return; // Exit the function if the text is successfully entered
//             } else {
//                 console.warn(`Validation failed on attempt ${attempt}: ${enteredText} does not include ${inputData}`);
//             }
//         } catch (error) {
//             console.error(`Attempt ${attempt} failed: ${error.message}`);
//         }

//         if (attempt < maxRetries) {
//             await delay(500);
//             console.log(`Retrying (${attempt + 1}/${maxRetries})...`);
//             // await delay(1000); // Optional: wait before retrying
//         }
//     }

//     throw new Error(`Failed to enter ${inputData} into ${dataTestId} after ${maxRetries} attempts`);
// }

async function enterDataIntoCombobox(page, dataTestId, inputData, maxRetries = 2) {
    const comboboxSelector = `div[data-testid="${dataTestId}"]`;
    const inputSelector = `${comboboxSelector} input.input-field__input`;

    // Function to check if the inputData is already in the value within parentheses
    const isDataAlreadyIncluded = async (inputData) => {
        const currentValue = await page.$eval(inputSelector, el => el.value);
        const regex = /\((\w{3})\)/; // Regex to match 3-letter strings inside parentheses
        const match = regex.exec(currentValue);
        return match && inputData.includes(match[1]);
    };

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const alreadyIncluded = await isDataAlreadyIncluded(inputData);

            if (!alreadyIncluded) {
                // Wait for the combobox and click it
                await page.locator(inputSelector).click({ clickCount: 2 });
                await page.keyboard.press('Backspace');

                // Type the input data
                await page.type(inputSelector, inputData);
                await delay(500);
                await page.keyboard.press('Enter');
            } else {
                console.log(`Input data "${inputData}" is already included in the value.`);
            }

            // Wait and validate if the input text is contained within the value
            await delay(1000); // Adjust if necessary
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
            await delay(500);
            console.log(`Retrying (${attempt + 1}/${maxRetries})...`);
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
    await delay(500);
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

function formatDateTime(date) {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0'); // Months are zero-based
    const day = date.getDate().toString().padStart(2, '0');
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');
    const seconds = date.getSeconds().toString().padStart(2, '0');
    return `${month}${day}${year}-${hours}${minutes}`;
}


(async () => {
    // Start time
    const startTime = new Date();
    
    // const fs = require('fs');
    let randomProxy

    // Read the JSON file
    // fs.readFile('proxies_free.json', 'utf8', (err, data) => {
    // if (err) {
    //     console.error('Error reading the file:', err);
    //     return;
    // }
    // const entries = JSON.parse(data);
    // const proxyList = entries.map(entry => entry.ip);
    // const randomProxy = proxyList[Math.floor(Math.random() * proxyList.length)];

    // });

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
            // `--proxy-server=${randomProxy}`,
            '--disable-features=BlockInsecurePrivateNetworkRequests'
        ],
    });

    try {
        const excelFilePath = 'EBB/Input/inputData_EBB2.xlsx'; // Replace with your actual Excel file path
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
            let flattenedData = [];
            const flyingWith = rowData.flyingWith;
            const leavingFrom = rowData.leavingFrom;
            const goingTo = rowData.goingTo;
            const cabinClass = rowData.cabinClass;
            const emiratesSkywardsTier = rowData.emiratesSkywardsTier;
            const oneWayOrRoundtrip = rowData.oneWayOrRoundtrip;

            // EXTRACTION HAPPENS HERE
            try{
            // Click the appropriate radio button based on the oneWayOrRoundtrip variable
            // if (oneWayOrRoundtrip === "One Way") {
                // console.log('Clicking One Way...');
                await page.locator('input.radio-button__input#OW0').click();
                // await page.click('input.radio-button__input#OW0');
                await delay(500);
                console.log('Clicked One Way...');
            // } else if (oneWayOrRoundtrip === "Round Trip") {
            //     await page.click('input.radio-button__input#RT1');
            //     await delay(500);
            //     console.log('Clicked Round Trip...');
            // }

            // FILL FORMS
            await selectComboboxOption(page, "combobox_Flying with", flyingWith);
            await enterDataIntoCombobox(page, "combobox_Leaving from", leavingFrom);
            await enterDataIntoCombobox(page, "combobox_Going to", goingTo);
            await selectComboboxOption(page, "combobox_Cabin class", cabinClass);
            await selectComboboxOption(page, "combobox_Emirates Skywards tier", emiratesSkywardsTier);
            await page.click('button[aria-label="Calculate"]');
            console.log('Submitted the form...');
            console.log('Waiting for results...');

            const waitForSelectors = Promise.race([
                page.waitForSelector('.tabs', { timeout: 40000 }),
                page.waitForSelector('h1:not([class]):not([id])', { timeout: 40000 })
            ]);

            await waitForSelectors;

            

            const isResults = await page.$('.tabs');
            const isAccessDenied = await page.$('h1:not([class]):not([id])');

            // Wait for the results page to load
            await page.waitForSelector('.skywards-miles-calculator__search-result.miles-calculator-result__section', { visible: true });

            // Data Scraping
            if (isResults) {
                const pageCardData = await page.$$eval('.tabs', (cards, cabinClass) => {
                    return cards.map(card => {
                        const actionElement = card.querySelector('a[title="Earn"]');
                        const action = actionElement ? actionElement.querySelector('.miles-calculator-result__tab-button--text').textContent.trim() : null;
                        

                        const fareDetails = Array.from(document.querySelectorAll('.miles-calculator-result__card-wrapper')).flatMap(wrapper => {
                            const fareTypes = ['flexplus','flex','saver', 'special'];
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

                                    // Add Fare Type Prefix
                                    let prefix = '';
                                    switch (cabinClass) {
                                        case 'Economy':
                                            prefix = 'Economy ';
                                            break;
                                        case 'Premium':
                                            prefix = '';
                                            break;
                                        case 'Business':
                                            prefix = 'Business ';
                                            break;
                                        case 'First':
                                            prefix = 'First ';
                                            break;
                                        default:
                                            prefix = ''; // No prefix if cabinClass is unrecognized
                                            break;}
                                    
                                    const brandedFare = prefix + fareType.charAt(0).toUpperCase() + fareType.slice(1);

                                    return {
                                        brandedFare: brandedFare,
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
            } catch {
            // ADD ERROR OCCURED
            flattenedData.push({
                action: "Earn",
                flyingWith: flyingWith,
                leavingFrom: leavingFrom,
                goingTo: goingTo,
                date: oneWayOrRoundtrip,
                cabinClass: cabinClass,
                brandedFare: "None",
                skywardTier: "None",
                skywardMiles: "None",
                tierMiles: "None"
            })

            allData = allData.concat(flattenedData);
            
        }
        }

        // Write all accumulated data to a single Excel file

        if (allData.length > 0) {
            // const workbook = XLSX.utils.book_new();

            // const actions = [...new Set(allData.map(item => item.action))];

            // actions.forEach(action => {
            //     const dataForSheet = allData.filter(item => item.action === action);
            //     const sheetData = dataForSheet.map(item => ({
            //         'Action': item.action,
            //         'Flying With': item.flyingWith,
            //         'Leaving from': item.leavingFrom,
            //         'Going to': item.goingTo,
            //         'Date (OW/RT)': item.date,
            //         'Cabin Class': item.cabinClass,
            //         'Emirates Skywards Tier': item.skywardTier,
            //         'Branded Fare': item.brandedFare,
            //         'Skyward Miles': item.skywardMiles,
            //         'Tier Miles': item.tierMiles
            //     }));
            //     const worksheet = XLSX.utils.json_to_sheet(sheetData);
            //     XLSX.utils.book_append_sheet(workbook, worksheet, action);
            // });

            // const fileName = `skywardsMilesData.xlsx`;
            // XLSX.writeFile(workbook, fileName);
            // console.log(`Excel file written to ${fileName}`);

            // Define the single action
            const action = 'Earn';

            // Filter data for the single action
            const dataForSheet = allData.filter(item => item.action === action);

            // Transform the data into the format
            const sheetData = dataForSheet.map(item => ({
                // 'Action': item.action,
                'Direction': item.date,
                'Airline': item.flyingWith,
                'Leaving from': item.leavingFrom,
                'Going to': item.goingTo,
                'Cabin Class': item.cabinClass,
                'Skywards Tier': item.skywardTier,
                'Branded Fare': item.brandedFare,
                'Skyward Miles': item.skywardMiles,
                'Tier Miles': item.tierMiles
            }));

            // Create a new workbook and add the sheet
            const workbook = XLSX.utils.book_new();
            const worksheet = XLSX.utils.json_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(workbook, worksheet, action);
            const formattedDateTime = formatDateTime(startTime);
            const filename = `EKS_Route_PER_${formattedDateTime}.xlsx`;
            // Write the workbook to a file
            XLSX.writeFile(workbook, filename);
            console.log(`Excel file written to ${filename}`);
            
            
        }

            // End time
        const endTime = new Date();
        
        // Calculate time elapsed
        const elapsedTimeSeconds = (endTime - startTime) / 1000; // Time in seconds
        const elapsedTimeMinutes = (endTime - startTime) / 60000; // Time in minutes

        // Get the number of entries
        const totalEntries = allData.filter(item => item.action === 'Earn').length;

        // console.log(allData)
        console.log(`Time elapsed: ${elapsedTimeSeconds} seconds`);
        console.log(`Time elapsed: ${elapsedTimeMinutes} minutes`);
        console.log(`Total number of entries: ${totalEntries}`);

    }
     catch (error) {
        console.error('Error:', error);
        await browser.close();
    } finally {
        await browser.close();
    }
})();
