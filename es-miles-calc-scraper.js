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

// Function to enter data into input fields
async function enterDataIntoCombobox(page, dataTestId, inputData) {
    const comboboxSelector = `div[data-testid="${dataTestId}"]`;
    await page.waitForSelector(comboboxSelector);
    await page.click(comboboxSelector);
    const inputSelector = `${comboboxSelector} input.input-field__input`;
    await page.type(inputSelector, inputData);
    await delay(500);
    await page.keyboard.press('Enter');
    console.log(`${dataTestId}: ${inputData}`);
    await delay(1000);
}

// Function to select an option from a dropdown menu
async function selectComboboxOption(page, comboboxTestId, optionText) {
    await page.click(`div[data-testid="${comboboxTestId}"]`);
    await page.waitForSelector('button.auto-suggest__item', { visible: true });
    await page.evaluate((text) => {
        const options = Array.from(document.querySelectorAll('button.auto-suggest__item'));
        const option = options.find(el => el.textContent.trim() === text);
        if (option) {
            option.click();
        }
    }, optionText);
    await delay(1000);
}

function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
}

(async () => {
    const browser = await puppeteer.launch({
        headless: false,
        defaultViewport: null,
    });

    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(60000);

    const flyingWith = 'Emirates';
    const leavingFrom = 'ABJ';
    const goingTo = 'ADD';
    const cabinClass = 'Economy';
    const emiratesSkywardsTier = 'Blue';
    const oneWayOrRoundtrip = "One Way"

    const homeUrl = 'https://www.emirates.com/ph/english/skywards/miles-calculator/';
    await page.goto(homeUrl);
    console.log('Navigated to Emirates mileage calculator...');
    await delay(1000);

    await handleCookiesPopup(page);
    // Click the appropriate radio button based on the oWorRoundtrip variable
    if (oneWayOrRoundtrip === "One Way") {
        await page.click('input.radio-button__input#OW0');
        await delay(500);
        console.log('Clicked Oneway...');
    } else if (oneWayOrRoundtrip === "Round Trip") {
        await page.click('input.radio-button__input#RT1');
        await delay(500);
        console.log('Clicked Oneway...');
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
    await page.waitForSelector('.skywards-miles-calculator__search-result.miles-calculator-result__section', { visible: true }); // Wait for the specific results section to be visible
    

    // Data Scraping
    if (isResults) {
        const pageCardData = await page.$$eval('.tabs', (cards) => {
            return cards.map(card => {
                const actionElement = card.querySelector('a[aria-selected="true"]');
                const action = actionElement ? actionElement.querySelector('.miles-calculator-result__tab-button--text').textContent.trim() : null;
                console.log('Action:', action);

                const fareDetails = Array.from(document.querySelectorAll('.miles-calculator-result__card-wrapper')).flatMap(wrapper => {
                    const fareTypes = ['special', 'saver', 'flex', 'flexplus'];
                    return fareTypes.map(fareType => {
                        const fareDiv = wrapper.querySelector(`.miles-card__card.miles-card__ek-economy-${fareType}`);
                        if (fareDiv) {
                            console.log(`${fareType} fareDiv exists`);
                            const milesContent = fareDiv.querySelector('.miles-card__content__miles');
                            console.log('milesContent exists');
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
                            console.log(`Skywards Miles (${fareType}): ${skywardMiles}`);
                
                            // Extract Tier Miles
                            const tierMilesDiv = milesContent.querySelector('.miles-card__tier-miles .miles-card__skywards-miles span');
                            if (tierMilesDiv) {
                                tierMiles = tierMilesDiv.textContent.trim();
                            }
                            console.log(`Tier Miles (${fareType}): ${tierMiles}`);
                
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
                
                console.log(fareDetails);
                

                return {
                    action,
                    fareDetails
                };
            });
        });

        if (pageCardData) {
            pageCardData.forEach(card => {
                card.fareDetails.forEach(fareDetail => {
                    flattenedData.push({
                        action: card.action,
                        leavingFrom: leavingFrom,
                        goingTo: goingTo,
                        brandedFare: fareDetail.brandedFare,
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

    // Write to Excel
    if (flattenedData.length > 0) {
        const workbook = XLSX.utils.book_new();

        const actions = [...new Set(flattenedData.map(item => item.action))];

        actions.forEach(action => {
            const dataForSheet = flattenedData.filter(item => item.action === action);
            const sheetData = dataForSheet.map(item => ({
                'Action': item.action,
                'Flying With': flyingWith,
                'Leaving from': item.leavingFrom,
                'Going to': item.goingTo,
                'Date (OW/Rt)': oneWayOrRoundtrip,
                'Cabin Class': cabinClass,
                'Emirates Skywards Tier': emiratesSkywardsTier,
                'Branded Fare': item.brandedFare,
                'Skyward Miles': item.skywardMiles,
                'Tier Miles': item.tierMiles
            }));
            const worksheet = XLSX.utils.json_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(workbook, worksheet, action);
        });

        XLSX.writeFile(workbook, 'flightData.xlsx');
        console.log('Data successfully written to flightData.xlsx');
    } else {
        console.log('No data found to write.');
    }

    await browser.close();
})();
