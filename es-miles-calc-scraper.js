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
    const goingTo = 'ABV';
    const cabinClass = 'Economy';
    const emiratesSkywardsTier = 'Blue';

    const homeUrl = 'https://www.emirates.com/ph/english/skywards/miles-calculator/';
    await page.goto(homeUrl);
    console.log('Navigated to Emirates mileage calculator...');
    await delay(1000);

    await handleCookiesPopup(page);
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
                const selectedTab = card.querySelector('a[aria-selected="true"]');
                const action = selectedTab ? selectedTab.querySelector('.miles-calculator-result__tab-button--text').textContent.trim() : null;
                const fareTypes = ['special', 'saver', 'flex', 'flexplus'];
                const fareDetails = fareTypes.map(fareType => {
                    const fareDiv = card.querySelector(`.miles-card__card.miles-card__ek-premiumeconomy-${fareType}`);
                    if (fareDiv) {
                        const milesContent = fareDiv.querySelector('.miles-card__content__miles');
                        const skywardMiles = milesContent ? milesContent.querySelector('.miles-card__skywards-miles span')?.textContent.trim() : 'N/A';
                        const tierMiles = milesContent ? milesContent.querySelector('.miles-card__tier-miles span')?.textContent.trim() : 'N/A';
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
        flattenedData.push({
            action: "Access Denied",
            leavingFrom: leavingFrom,
            goingTo: goingTo,
            brandedFare: 'N/A',
            skywardMiles: 'N/A',
            tierMiles: 'N/A'
        });
    }

    if (flattenedData.length > 0) {
        const workbook = XLSX.utils.book_new();

        const actions = [...new Set(flattenedData.map(item => item.action))];

        actions.forEach(action => {
            const dataForSheet = flattenedData.filter(item => item.action === action);
            const sheetData = dataForSheet.map(item => ({
                'Leaving from': item.leavingFrom,
                'Going to': item.goingTo,
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
