// This Puppeteer script enables you to scrape flight details from Emirates website.
// It returns scraping results to external JSON file (flight_data.json).
// Remember to change parameters (airportFrom, airportTo, departDate, returnDate) according to your needs.

// VIDEO TUTORIAL
// Complete video guide available on my YouTube: https://youtu.be/cMxD-LxEd04


const puppeteer = require('puppeteer');
const fs = require('fs');

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
    const departDate = '102025'; // Enter departure date where "1522024" stands for 15th March 2024 (month January has value "0")
    const returnDate = '112025'; // Enter departure date where "1822024" stands for 18th March 2024 (month January has value "0")

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

    const cardData = [];

    const isFlightListBody = await page.$('.ts-fbr-flight-list__body');
    const isFlightsRow = await page.$('.flights-row');
    const isAccessDenied = await page.$('h1:not([class]):not([id])');

    if (isFlightListBody) {
        // Handling the first type of result page with '.ts-fbr-flight-list__body'

        console.log('Scraping flight data...');

        const pageCardData = await page.$$eval('.ts-fbr-flight-list__body > div', (cards) => {
            return cards.map((card) => {

                // Extract flight information from each card

                // Origin airport
                const originAirportElement = card.querySelector('.ts-fie__place > p').textContent.trim();
                const iataRegex = /[A-Z]{3}/;
                const originAirportMatches = originAirportElement.match(iataRegex);
                const originAirport = originAirportMatches ? originAirportMatches[0] : "N/A";

                // Destination airport
                const destinationAirport = card.querySelector('.ts-fie__place.ts-fie__right-side > p > span:nth-child(2)').textContent.trim();

                // Departure time
                const departureTime = card.querySelector('.ts-fie__departure').textContent.trim();

                // Arrival time
                const arrivalTimeElement = card.querySelector('.ts-fie__arrival');
                const arrivalTimeMatch = arrivalTimeElement ? arrivalTimeElement.textContent.match(/\d+:\d+/) : null;
                const arrivalTime = arrivalTimeMatch ? arrivalTimeMatch[0] : "N/A";

                // Extract the number from the <sup> tag for supDays
                const supDaysElement = card.querySelector('.ts-fie__arrival > sup');
                const supDays = supDaysElement ? parseInt(supDaysElement.textContent.trim()) : 0;

                // Travel date
                const headerSelector = card.closest('.ts-fbr-flight-list');
                const dateElement = headerSelector.querySelector('.ts-fbr-flight-list__header-date');
                const date = dateElement ? dateElement.textContent.trim() : "N/A";

                // Travel time
                const travelTimeElement = card.querySelector('.ts-fie__infographic time span:nth-child(2)');
                const travelTime = travelTimeElement ? travelTimeElement.textContent.trim() : "N/A";

                // Extract the number of connections
                const connectionElement = card.querySelector('.ts-fie__infographic a span[aria-hidden="true"]');
                const connectionsText = connectionElement ? connectionElement.textContent : "";

                // Determine the number of connections using a regular expression
                const connectionsMatch = connectionsText.match(/\d+/);
                const connections = connectionsMatch ? parseInt(connectionsMatch[0]) : 0;

                // Extract prices associated with this card
                const priceData = Array.from(card.querySelectorAll('.ts-fbr-flight-list-row__options > div')).map((priceElement) => {
                    const priceClass = priceElement.querySelector('.ts-fbr-option__class').textContent.trim();
                    const priceValueElement = priceElement.querySelector('.ts-fbr-option__price');
                    const priceCurrencyElement = priceElement.querySelector('.ts-fbr-option__currency');

                    if (priceValueElement && priceCurrencyElement) {
                        const priceValue = priceValueElement.textContent.trim();
                        const priceCurrencySelector = priceCurrencyElement.getAttribute('data-from');
                        const currencyRegex = /[A-Z]{3}/;
                        const currencyMatches = priceCurrencySelector ? priceCurrencySelector.match(currencyRegex) : null;
                        const priceCurrency = currencyMatches ? currencyMatches[0] : "N/A";

                        return {
                            class: priceClass,
                            price: priceValue,
                            currency: priceCurrency
                        };
                    }

                    return null; // Handle cases where price data is missing
                }).filter(price => price !== null);

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

        // Save the data to a JSON file
        fs.writeFileSync('flight_data.json', JSON.stringify(pageCardData, null, 2));
        console.log('Saved flight data to flight_data.json.');

    } else if (isFlightsRow) {
        // Handling the second type of result page with '.flights-row'

        console.log('Scraping flight data...');

        const pageCardData = await page.$$eval('.flights-row', (cards) => {
            return cards.map((card) => {

                // Extract flight information from each card

                // Price
                const priceValueElement = card.querySelector('.ts-ifl-row__footer-price > span');
                const priceValue = priceValueElement ? priceValueElement.textContent.trim() : "N/A";

                // Currency
                const priceCurrencyElement = card.querySelector('.ts-ifl-row__footer-price');
                const currencyRegex = /[A-Z]{3}/;
                const currencyMatches = priceCurrencyElement ? priceCurrencyElement.textContent.trim().match(currencyRegex) : null;
                const priceCurrency = currencyMatches ? currencyMatches[0] : "N/A";

                // Extract prices associated with this card
                const flightData = Array.from(card.querySelectorAll('.ts-ifl-row__body-item')).map((flightElement) => {

                    // Origin airport
                    const originAirportElement = flightElement.querySelector('.ts-fie__place > p').textContent.trim();
                    const iataRegex = /[A-Z]{3}/;
                    const originAirportMatches = originAirportElement.match(iataRegex);
                    const originAirport = originAirportMatches ? originAirportMatches[0] : "N/A";

                    // Destination airport
                    const destinationAirport = flightElement.querySelector('.ts-fie__place.ts-fie__right-side > p > span:nth-child(2)').textContent.trim();

                    // Travel date
                    const dateElement = flightElement.querySelector('.ts-fip__date-container time');
                    const date = dateElement ? dateElement.textContent.trim().replace(/\n/g, '') : "N/A";

                    // Departure time
                    const departureTime = flightElement.querySelector('.ts-fie__departure').textContent.trim();

                    // Arrival time
                    const arrivalTimeElement = flightElement.querySelector('.ts-fie__arrival');
                    const arrivalTimeMatch = arrivalTimeElement ? arrivalTimeElement.textContent.match(/\d+:\d+/) : null;
                    const arrivalTime = arrivalTimeMatch ? arrivalTimeMatch[0] : "N/A";

                    // Extract the number from the <sup> tag for supDays
                    const supDaysElement = flightElement.querySelector('.ts-fie__arrival > sup');
                    const supDays = supDaysElement ? parseInt(supDaysElement.textContent.trim()) : 0;

                    // Travel time
                    const travelTimeElement = flightElement.querySelector('.ts-fie__infographic time span:nth-child(2)');
                    const travelTime = travelTimeElement ? travelTimeElement.textContent.trim() : "N/A";

                    // Extract the number of connections
                    const connectionElement = flightElement.querySelector('.ts-fie__infographic a');
                    const connectionsText = connectionElement ? connectionElement.textContent : "";

                    // Determine the number of connections using a regular expression
                    const connectionsMatch = connectionsText.match(/\d+/);
                    const connections = connectionsMatch ? parseInt(connectionsMatch[0]) : 0;

                    return {
                        originAirport: originAirport,
                        destinationAirport: destinationAirport,
                        date: date,
                        departureTime: departureTime,
                        arrivalTime: arrivalTime,
                        supDays: supDays,
                        travelTime: travelTime,
                        connections: connections
                    };

                });

                return {
                    result: "trip",
                    price: priceValue,
                    currency: priceCurrency,
                    flightData: flightData
                };
            });
        });

        cardData.push(...pageCardData);

        // Save the data to a JSON file
        fs.writeFileSync('flight_data.json', JSON.stringify(pageCardData, null, 2));
        console.log('Saved flight data to flight_data.json.');

    } else if (isAccessDenied) {
        throw new Error('Access denied.');
    } else {
        // Handle the case where the page structure doesn't match either selector
        console.log('Unknown result page structure. Unable to extract data.');
    }

    await browser.close();

})();
