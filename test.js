const puppeteer = require('puppeteer');
const XLSX = require('xlsx'); // Import the xlsx library

// Function to handle cookies window
async function handleCookiesPopup(page) {
    const cookiesButton = await page.$('#onetrust-accept-btn-handler');
    if (cookiesButton) {
        await cookiesButton.click();
        console.log('Clicked the cookies accept button...');
    }
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

    const airportFrom = 'FRA'; // Enter IATA code of origin airport
    const airportTo = 'JFK'; // Enter IATA code of destination airport
    const departDate = '102025'; // Enter departure date where "102025" stands for 10th February 2025
    const returnDate = '112025'; // Enter return date where "112025" stands for 11th February 2025

    // Navigate to Emirates base fare page
    const homeUrl = 'https://www.emirates.com/us/english/';
    await page.goto(homeUrl);
    console.log('Navigated to Emirates website...');

    await delay(1000);

    await handleCookiesPopup(page);

    await page.click('.js-origin-dropdown input');

    await page.keyboard.down('Control');
    await page.keyboard.press('A');
    await page.keyboard.up('Control');
    await page.keyboard.press('Backspace');
    await page.type('.js-origin-dropdown input', airportFrom);
    await delay(500);
    await page.keyboard.press('Enter');
    console.log('Entered the origin airport...');

    await delay(1000);

    await page.click('.destination-dropdown input');

    await page.keyboard.down('Control');
    await page.keyboard.press('A');
    await page.keyboard.up('Control');
    await page.keyboard.press('Backspace');
    await page.type('.destination-dropdown input', airportTo);
    await delay(500);
    await page.keyboard.press('Enter');
    console.log('Entered the destination airport...');

    await delay(1000);

    await page.click('#search-flight-date-picker--depart');
    await delay(1000);

    async function selectDate(targetDate) {
        let foundTargetDate = false;
        let disabledArrow = false;

        while (!foundTargetDate && !disabledArrow) {
            const button = await page.$(`button[data-string="${targetDate}"]`);

            if (button) {
                foundTargetDate = true;
                await button.click();
                console.log('Selected a date for departure or return...');
            } else {
                const isDisabled = await page.$eval('button.icon-arrow-right', (btn) =>
                    btn.hasAttribute('disabled')
                );

                if (isDisabled) {
                    console.error('Error: Right arrow is not clickable.');
                    await browser.close();
                    return;
                } else {
                    await page.click('button.icon-arrow-right');
                }
            }
        }
    }

    await selectDate(departDate);

    await page.click('#search-flight-date-picker--return');
    await delay(1000);

    await selectDate(returnDate);
    await delay(1000);

    await page.click('button[type="submit"]');
    console.log('Submitted the form...');
    await page.waitForNavigation({ timeout: 30000 });
    console.log('Waiting for results...');

    const waitForSelectors = Promise.race([
        page.waitForSelector('.ts-fbr-flight-list__body', { timeout: 90000 }),
        page.waitForSelector('.flights-row', { timeout: 90000 }),
        page.waitForSelector('h1:not([class]):not([id])', { timeout: 90000 })
    ]);

    await waitForSelectors;

    let flattenedData = [];

    const isFlightListBody = await page.$('.ts-fbr-flight-list__body');
    const isFlightsRow = await page.$('.flights-row');
    const isAccessDenied = await page.$('h1:not([class]):not([id])');

    if (isFlightListBody) {
        console.log('Scraping flight data...');

        const pageCardData = await page.$$eval('.ts-fbr-flight-list__body > div', (cards) => {
            return cards.map((card) => {

                const originAirport = card.querySelector('.ts-fie__place > p').textContent.trim();
                const destinationAirport = card.querySelector('.ts-fie__place.ts-fie__right-side > p > span:nth-child(2)').textContent.trim();
                const departureTime = card.querySelector('.ts-fie__departure').textContent.trim();
                const arrivalTime = card.querySelector('.ts-fie__arrival').textContent.trim();
                const travelTime = card.querySelector('.ts-fie__infographic time span:nth-child(2)').textContent.trim();
                const supDaysElement = card.querySelector('.ts-fie__arrival > sup');
                const supDays = supDaysElement ? parseInt(supDaysElement.textContent.trim()) : 0;
                const connectionsElement = card.querySelector('.ts-fie__infographic a span[aria-hidden="true"]');
                const connections = connectionsElement ? parseInt(connectionsElement.textContent.match(/\d+/)[0]) : 0;

                const headerSelector = card.closest('.ts-fbr-flight-list');
                const dateElement = headerSelector.querySelector('.ts-fbr-flight-list__header-date');
                const date = dateElement ? dateElement.textContent.trim() : "N/A";

                const priceData = Array.from(card.querySelectorAll('.ts-fbr-flight-list-row__options > div')).map((priceElement) => {
                    const priceClass = priceElement.querySelector('.ts-fbr-option__class').textContent.trim();
                    const priceValueElement = priceElement.querySelector('.ts-fbr-option__price');
                    const priceCurrencyElement = priceElement.querySelector('.ts-fbr-option__currency');
                    const priceValue = priceValueElement ? priceValueElement.textContent.trim() : "N/A";
                    const currencyRegex = /[A-Z]{3}/;
                    const currencyMatches = priceCurrencyElement ? priceCurrencyElement.getAttribute('data-from').match(currencyRegex) : null;
                    const priceCurrency = currencyMatches ? currencyMatches[0] : "N/A";

                    return {
                        priceClass,
                        priceValue,
                        priceCurrency
                    };
                });

                return {
                    result: "fare",
                    originAirport,
                    destinationAirport,
                    date,
                    departureTime,
                    arrivalTime,
                    travelTime,
                    supDays,
                    connections,
                    prices: priceData
                };
            });
        });

        // Flatten the data
        flattenedData = pageCardData.flatMap(card => {
            return card.prices.map(price => ({
                result: card.result,
                originAirport: card.originAirport,
                destinationAirport: card.destinationAirport,
                date: card.date,
                departureTime: card.departureTime,
                arrivalTime: card.arrivalTime,
                travelTime: card.travelTime,
                supDays: card.supDays,
                connections: card.connections,
                class: price.priceClass,
                price: price.priceValue,
                currency: price.priceCurrency
            }));
        });

    } else if (isFlightsRow) {
        console.log('Scraping flight data...');

        const pageCardData = await page.$$eval('.flights-row', (cards) => {
            return cards.map((card) => {
                const priceValueElement = card.querySelector('.ts-ifl-row__footer-price > span');
                const priceValue = priceValueElement ? priceValueElement.textContent.trim() : "N/A";
                const priceCurrencyElement = card.querySelector('.ts-ifl-row__footer-price');
                const currencyRegex = /[A-Z]{3}/;
                const currencyMatches = priceCurrencyElement ? priceCurrencyElement.textContent.trim().match(currencyRegex) : null;
                const priceCurrency = currencyMatches ? currencyMatches[0] : "N/A";

                const flightData = Array.from(card.querySelectorAll('.ts-ifl-row__body-item')).map((flightElement) => {

                    const originAirport = flightElement.querySelector('.ts-fie__place > p').textContent.trim();
                    const destinationAirport = flightElement.querySelector('.ts-fie__place.ts-fie__right-side > p > span:nth-child(2)').textContent.trim();
                    const dateElement = flightElement.querySelector('.ts-fip__date-container time');
                    const date = dateElement ? dateElement.textContent.trim().replace(/\n/g, '') : "N/A";
                    const departureTime = flightElement.querySelector('.ts-fie__departure').textContent.trim();
                    const arrivalTime = flightElement.querySelector('.ts-fie__arrival').textContent.trim();
                    const travelTime = flightElement.querySelector('.ts-fie__infographic time span:nth-child(2)').textContent.trim();
                    const supDaysElement = flightElement.querySelector('.ts-fie__arrival > sup');
                    const supDays = supDaysElement ? parseInt(supDaysElement.textContent.trim()) : 0;
                    const connectionsElement = flightElement.querySelector('.ts-fie__infographic a span[aria-hidden="true"]');
                    const connections = connectionsElement ? parseInt(connectionsElement.textContent.match(/\d+/)[0]) : 0;

                    return {
                        result: "fare",
                        originAirport,
                        destinationAirport,
                        date,
                        departureTime,
                        arrivalTime,
                        travelTime,
                        supDays,
                        connections,
                        class: "N/A", // No specific class in this case
                        price: priceValue,
                        currency: priceCurrency
                    };
                });

                return flightData;
            });
        });

        // Flatten the data
        flattenedData = pageCardData.flat();

    } else if (isAccessDenied) {
        const errorData = {
            result: "access denied",
            originAirport,
            destinationAirport,
            date: departDate,
            errorMsg: "Access Denied"
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

    await browser.close();
})();
