# Sheets Agent for QuickBooks

A Google Sheets add-on that allows you to interact with QuickBooks Online data through a natural language chat interface.

## Features

- Chat-based interface for QuickBooks data queries
- Natural language processing to interpret user queries
- Fetch reports (Profit & Loss, Balance Sheet)
- Query QuickBooks entities (Invoices, Expenses, Customers)
- Import data directly into Google Sheets
- OAuth 2.0 authentication with QuickBooks

## Setup Instructions

### 1. Create QuickBooks Developer Account

1. Sign up for a QuickBooks Developer account at [developer.intuit.com](https://developer.intuit.com/)
2. Create a new app
3. Get your Client ID and Client Secret
4. Add the following redirect URI to your app:
   ```
   https://script.google.com/macros/d/{SCRIPT_ID}/usercallback
   ```
   Replace `{SCRIPT_ID}` with your Google Apps Script ID

### 2. Configure the Add-on

1. Open the Sheets Agent sidebar in Google Sheets
2. Click the settings (gear) icon
3. Enter your QuickBooks Client ID, Client Secret, and Company ID
4. Click "Save"
5. Click "Connect" to authenticate with QuickBooks

## Usage

Once connected to QuickBooks, you can use natural language queries such as:

- "Fetch profit and loss report for last month"
- "Get invoices from last quarter into Sheet1"
- "Show me all customers"
- "Fetch balance sheet year to date"
- "Get expenses for this month"

## Development

This project is built with:

- Google Apps Script
- CLASP (Command Line Apps Script Projects)
- QuickBooks Online API
- OAuth2 Library for Google Apps Script

### Local Development

1. Install CLASP:
   ```
   npm install -g @google/clasp
   ```

2. Clone the project:
   ```
   clasp clone {SCRIPT_ID}
   ```

3. Make changes and push:
   ```
   clasp push
   ```

## Data Privacy

Sheets Agent only accesses the data you specifically request through queries. All API calls are logged in a hidden "Action Log" sheet for transparency.

## Support

For support or feature requests, please file an issue on the GitHub repository. 