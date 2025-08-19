# Google Apps Script for Automated Order Delivery Email System


## Prerequisites

- Google Account with access to Google Sheets and Apps Script
- SparkPost account and API key

## Setup Instructions

### 1. Google Sheets Configuration

Your spreadsheet must contain these worksheets with the exact names:
- `Offline Orders`
- `ebay, Amazon & Walmart`
- `NNC NES & non-wire`

#### Required Column Headers
Each worksheet needs these columns (case insensitive):
- **PO**: Purchase order number
- **Date**: Order date
- **Type**: Order type
- **Contact**: Customer name
- **Email**: Customer email address
- **Qty**: Quantity ordered
- **Product**: Product description
- **Status**: Order status (the system triggers when this becomes "delivered")
- **Tracking Number** or **Tracing Number**: Shipment tracking info

### 2. Google Apps Script Setup

1. **Open Google Apps Script**:
   - Click "New Project"

2. **Add the Code**:
   - Delete the default `myFunction`
   - Paste the entire contents of `index.js`
   - Save the project

3. **Enable Google Sheets API**:
   - Click on "Services" in the left sidebar
   - Click "Add a service"
   - Select "Google Sheets API"
   - Click "Add"

### 3. SparkPost Configuration

1. **Configure the Token in Apps Script**:
   - Uncomment the `setSparkPostToken()` function (lines 428-432)
   - Replace `'YOUR_TOKEN_HERE'` with the SparkPost API key
   - Save the script
   - Run the `setSparkPostToken`
   - **Important**: Comment out this function after running it once, and **Delete** the token

### 4. Set Up the Trigger

1. **Manual Setup**:
   - In Apps Script, click triggers in the left sidebar
   - Click "Add Trigger"
   - Configure:
     - Function: `onEdit`
     - Event source: "From spreadsheet"
     - Event type: "On edit"
   - Click "Save"


### 5. Testing Mode

There is an inbuilt testing mode

#### Enable Testing Mode:
```javascript
const TESTING_MODE = true; // Set to true for testing
const TEST_EMAIL = "test-email@company.com"; // an email for tests
```

#### In Testing Mode:
- **Email redirection**: All emails go to `TEST_EMAIL` instead of customers
- **Clear indicators**: Testing emails include banners show customer email
- **Safe marking**: Orders are marked as "TEST-SENT" instead of "YES"
- **Console logging**: more logs for debugging

#### Deployment Mode:
```javascript
const TESTING_MODE = false; // Set to false for final deployment
```
