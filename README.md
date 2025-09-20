# Emirates Skywards Miles Calculator Scraper

A Node.js web scraper that automates the collection of Emirates Skywards miles data from the official Emirates website. This tool processes flight route data from Excel files and extracts comprehensive mileage information for different fare types and cabin classes.

## 🚀 Features

- **Automated Web Scraping**: Uses Puppeteer with stealth plugins to avoid detection
- **Excel Integration**: Reads input data from Excel files and exports results to Excel
- **Multi-Route Support**: Organized by airport codes (LHR, EBB, EDI, EWR, EZE, LIS, PER)
- **Comprehensive Data Extraction**: Collects Skywards miles and tier miles for all fare types
- **Error Handling**: Robust error handling with retry mechanisms
- **Batch Processing**: Processes multiple flight routes in sequence
- **Data Validation**: Validates input data and handles missing information gracefully

## 📋 Prerequisites

- Node.js (version 14 or higher)
- npm (Node Package Manager)
- Windows/macOS/Linux operating system

## 🛠️ Installation

1. **Clone or download the project**
   ```bash
   git clone <repository-url>
   cd es-miles-calc-scraper
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Verify installation**
   ```bash
   npm start
   ```


## 📊 Input Data Format

The scraper expects Excel files with the following columns:

| Column | Description | Example Values |
|--------|-------------|----------------|
| `flyingWith` | Airline partner | "Emirates", "Qantas", "Virgin Atlantic" |
| `leavingFrom` | Origin airport | "LHR", "DXB", "JFK" |
| `goingTo` | Destination airport | "DXB", "SYD", "LAX" |
| `cabinClass` | Cabin class | "Economy", "Premium", "Business", "First" |
| `emiratesSkywardsTier` | Skywards tier | "Blue", "Silver", "Gold", "Platinum" |
| `oneWayOrRoundtrip` | Trip type | "One Way", "Round Trip" |

## 🔧 Configuration

### Route and Batch Settings

Modify these variables in the main script:

```javascript
let route = "LHR"        // Airport code for the route
let batch = "test"       // Batch identifier
```

### Browser Configuration

The scraper uses the following browser settings:

```javascript
const browser = await puppeteer.launch({
    headless: false,     // Set to true for headless mode
    args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--window-size=1920x1080',
        // ... additional security and performance args
    ],
});
```

## �� Usage

### Basic Usage

1. **Prepare your input data**:
   - Create an Excel file in the appropriate route folder
   - Follow the naming convention: `inputData_[ROUTE][BATCH].xlsx`
   - Ensure all required columns are present

2. **Configure the scraper**:
   - Update the `route` and `batch` variables in the script
   - Adjust any other settings as needed

3. **Run the scraper**:
   ```bash
   npm start
   # or
   node es-miles-calc-scraper.js
   ```

### Advanced Usage

#### Processing Multiple Routes

To process different routes, modify the script for each run:

```javascript
// For LHR route
let route = "LHR"
let batch = "batch1"

// For EBB route  
let route = "EBB"
let batch = "batch2"
```

#### Headless Mode

For production runs, enable headless mode:

```javascript
const browser = await puppeteer.launch({
    headless: true,  // Enable headless mode
    // ... other options
});
```

## 📈 Output Data

The scraper generates Excel files with the following structure:

| Column | Description |
|--------|-------------|
| `Direction` | Trip type (One Way/Round Trip) |
| `Airline` | Airline partner |
| `Leaving from` | Origin airport |
| `Going to` | Destination airport |
| `Cabin Class` | Cabin class |
| `Skywards Tier` | Emirates Skywards tier |
| `Branded Fare` | Fare type (e.g., "Economy Saver", "Business Flex") |
| `Skyward Miles` | Skywards miles earned |
| `Tier Miles` | Tier miles earned |


## 🛡️ Error Handling

The scraper includes comprehensive error handling:

- **Input Validation**: Validates Excel file structure and required columns
- **Retry Logic**: Retries failed operations up to 2 times
- **Browser Recovery**: Handles browser disconnections
- **Data Fallback**: Provides default values for missing data
- **Access Denied**: Handles blocked requests gracefully

## 📝 Logging

The scraper provides detailed logging:

- Progress updates for each step
- Success/failure notifications
- Performance metrics (execution time, entry count)
- Error messages with context

## ⚠️ Important Notes

1. **Rate Limiting**: The scraper includes delays to avoid overwhelming the target website
2. **Stealth Mode**: Uses stealth plugins to avoid detection
3. **Data Accuracy**: Results depend on the current state of the Emirates website
4. **Legal Compliance**: Ensure compliance with website terms of service
5. **Resource Usage**: Browser automation can be resource-intensive

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is for educational and research purposes. Please ensure compliance with Emirates' terms of service and applicable laws.

## 🆘 Support

For issues and questions:
1. Check the troubleshooting section
2. Review error logs
3. Verify input data format
4. Check website accessibility


**Disclaimer**: Users are responsible for ensuring compliance with website terms of service and applicable laws. The authors are not responsible for any misuse of this software.