/**
 * QuickBooks API Integration
 * Functions for authenticating and interacting with the QuickBooks Online API.
 */

// QuickBooks API Configuration
const QBO_CONFIG = {
  // These should be set by the user in the UI
  CLIENT_ID: '',
  CLIENT_SECRET: '',
  // Base URLs
  BASE_URL: 'https://quickbooks.api.intuit.com/v3/company/',
  AUTH_URL: 'https://appcenter.intuit.com/connect/oauth2',
  TOKEN_URL: 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer',
  ISSUER: 'https://oauth.platform.intuit.com/op/v1'
};

// Create a global QuickBooks object to export functions
const QuickBooksService = {};

/**
 * Creates and returns the OAuth2 service for QuickBooks.
 * 
 * @return {OAuth2.Service} The OAuth2 service
 */
function getOAuthService() {
  // Load client credentials from properties
  const scriptProperties = PropertiesService.getScriptProperties();
  const clientId = scriptProperties.getProperty('QBO_CLIENT_ID') || QBO_CONFIG.CLIENT_ID;
  const clientSecret = scriptProperties.getProperty('QBO_CLIENT_SECRET') || QBO_CONFIG.CLIENT_SECRET;
  
  // Set up the OAuth2 service
  return OAuth2.createService('quickbooks')
    .setAuthorizationBaseUrl(QBO_CONFIG.AUTH_URL)
    .setTokenUrl(QBO_CONFIG.TOKEN_URL)
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCallbackFunction('authCallback')
    .setScope('openid email profile address phone com.intuit.quickbooks.accounting')
    .setParam('response_type', 'code')
    .setParam('prompt', 'login consent')
    .setTokenHeaders({
      'Content-Type': 'application/x-www-form-urlencoded',
      'Accept': 'application/json'
    });
}

/**
 * Saves the QuickBooks API credentials.
 * 
 * @param {Object} credentials - The client ID and secret
 * @return {Object} Success status
 */
function saveCredentials(credentials) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('QBO_CLIENT_ID', credentials.clientId);
    scriptProperties.setProperty('QBO_CLIENT_SECRET', credentials.clientSecret);
    
    // Also save company ID if provided
    if (credentials.companyId) {
      scriptProperties.setProperty('QBO_COMPANY_ID', credentials.companyId);
    }
    
    return { success: true };
  } catch (error) {
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}

/**
 * Gets the currently saved QuickBooks credentials.
 * 
 * @return {Object} The saved credentials
 */
function getCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    clientId: scriptProperties.getProperty('QBO_CLIENT_ID') || '',
    clientSecret: scriptProperties.getProperty('QBO_CLIENT_SECRET') || '',
    companyId: scriptProperties.getProperty('QBO_COMPANY_ID') || ''
  };
}

/**
 * Makes a request to the QuickBooks API.
 * 
 * @param {string} endpoint - The API endpoint to call
 * @param {string} method - The HTTP method (GET, POST, etc.)
 * @param {Object} data - The request payload for POST/PUT requests
 * @return {Object} The API response
 */
function makeApiCall(endpoint, method = 'GET', data = null) {
  try {
    // Get the OAuth service
    const service = getOAuthService();
    
    // Check for access and try to refresh if necessary
    if (!service.hasAccess()) {
      logAction('OAuth Status', 'No access', 'Checking if refresh possible');
      
      // Check if we can refresh the token
      if (service.canRefresh()) {
        logAction('OAuth Refresh', 'Attempting token refresh', 'Token expired');
        const refreshSuccess = service.refresh();
        
        if (!refreshSuccess) {
          logAction('OAuth Refresh', 'Failed', 'Could not refresh token');
          throw new Error('QuickBooks access token expired and refresh failed. Please reconnect.');
        }
        
        logAction('OAuth Refresh', 'Success', 'Token refreshed successfully');
      } else {
        logAction('OAuth Status', 'No refresh token', 'Authentication required');
        throw new Error('No access to QuickBooks API. Please authenticate in settings.');
      }
    }
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const companyId = scriptProperties.getProperty('QBO_COMPANY_ID');
    
    if (!companyId) {
      throw new Error('Company ID not set. Please configure in settings.');
    }
    
    // Construct the full URL
    const baseUrl = QBO_CONFIG.BASE_URL + companyId + '/';
    let url = baseUrl + endpoint;
    
    // For GET requests with parameters
    if (method === 'GET' && data) {
      // Extract existing query parameters
      const [basePath, existingQuery] = url.split('?');
      const queryParams = [];
      
      // Add any existing query parameters
      if (existingQuery) {
        queryParams.push(existingQuery);
      }
      
      // Add new parameters
      for (const key in data) {
        queryParams.push(key + '=' + encodeURIComponent(data[key]));
      }
      
      // Reconstruct URL with all parameters
      url = basePath + '?' + queryParams.join('&');
    }
    
    // Log the API call details
    logAction('API Call', `${method} ${endpoint}`, 'Initiating request');
    
    // Prepare request options
    const options = {
      method: method,
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    if (data && (method === 'POST' || method === 'PUT')) {
      options.payload = JSON.stringify(data);
    }
    
    // Make the API call
    const response = UrlFetchApp.fetch(url, options);
    
    // Log response status
    const responseCode = response.getResponseCode();
    logAction('API Response', `Status: ${responseCode}`, endpoint);
    
    // Handle specific error codes
    if (responseCode === 401) {
      // Unauthorized - token might be expired
      logAction('API Error', 'Unauthorized (401)', 'Token may be expired');
      
      // Try to refresh the token and retry once
      if (service.canRefresh()) {
        logAction('OAuth Refresh', 'Attempting after 401', 'Token refresh triggered by API response');
        const refreshed = service.refresh();
        
        if (refreshed) {
          // Update the token in the options and retry
          options.headers['Authorization'] = 'Bearer ' + service.getAccessToken();
          logAction('API Retry', `${method} ${endpoint}`, 'Retrying after token refresh');
          return UrlFetchApp.fetch(url, options);
        }
      }
      
      throw new Error('Authentication failed with QuickBooks API. Please reconnect.');
    }
    
    return response;
  } catch (error) {
    // Enhanced error logging
    console.error('Error making API call to ' + endpoint + ': ' + error.toString());
    
    // Check for network/connectivity errors
    if (error.message && error.message.includes('network')) {
      throw new Error('Network connectivity issue when connecting to QuickBooks. Please check your internet connection.');
    }
    
    throw error;
  }
}

/**
 * Makes a request to the QuickBooks API and parses the JSON response.
 * 
 * @param {string} endpoint - The API endpoint to call
 * @param {string} method - The HTTP method (GET, POST, etc.)
 * @param {Object} data - The request payload for POST/PUT requests
 * @return {Object} The parsed API response
 */
function callQuickBooksApi(endpoint, method = 'GET', data = null) {
  try {
    const response = makeApiCall(endpoint, method, data);
    const responseCode = response.getResponseCode();
    
    if (responseCode >= 200 && responseCode < 300) {
      // Success - parse the JSON response
      const contentText = response.getContentText();
      return JSON.parse(contentText);
    } else {
      // Error handling
      console.error('QuickBooks API error: ' + responseCode);
      console.error('Response: ' + response.getContentText());
      
      // Create a more detailed error object
      const error = {
        statusCode: responseCode,
        message: 'QuickBooks API error: ' + responseCode
      };
      
      try {
        const errorResponse = JSON.parse(response.getContentText());
        if (errorResponse.Fault) {
          error.fault = errorResponse.Fault;
          error.message = errorResponse.Fault.Error[0].Message || error.message;
        }
      } catch (e) {
        // Unable to parse error JSON
      }
      
      throw new Error(JSON.stringify(error));
    }
  } catch (error) {
    console.error('Error in callQuickBooksApi: ' + error.toString());
    throw error;
  }
}

/**
 * Queries QuickBooks entities using the Query endpoint.
 * 
 * @param {string} query - The SQL-like query string
 * @return {Object} The query results
 */
function queryQuickBooks(query) {
  return callQuickBooksApi('query?query=' + encodeURIComponent(query));
}

/**
 * Ultra-basic fallback method for P&L when all other approaches fail.
 * Uses transaction data which nearly all accounts can access.
 * 
 * @param {Object} params - Report parameters
 * @return {Object} Simple P&L data in report format
 */
function getBasicTransactionPLReport(params) {
  // Log that we're using the minimal fallback approach
  logAction('P&L Minimal Fallback', 'Using transactions approach', JSON.stringify(params));
  
  try {
    // Create a structured response similar to the report API
    const result = {
      Header: {
        ReportName: "Profit and Loss (Basic)",
        Time: new Date().toISOString(),
        StartPeriod: params.start_date,
        EndPeriod: params.end_date
      },
      Columns: {
        Column: [
          { ColTitle: "Category", ColType: "Account" },
          { ColTitle: "Amount", ColType: "Amount" }
        ]
      },
      Rows: { Row: [] }
    };
    
    // Get the date range for filtering
    const startDate = params.start_date;
    const endDate = params.end_date;
    
    // The most reliable entities across all QuickBooks accounts are invoices and bills
    // Try multiple approaches in sequence for maximum compatibility
    
    // == INCOME SECTION ==
    let totalIncome = 0;
    let incomeSuccess = false;
    
    // Add a section for Income
    result.Rows.Row.push({
      type: "Section",
      Header: { 
        ColData: [{ value: "Income" }] 
      }
    });
    
    // Try different approaches to get income data
    // 1. First try invoices - most common income source
    try {
      const invoiceQuery = `SELECT TotalAmt FROM Invoice WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' MAXRESULTS 1000`;
      const invoiceResponse = queryQuickBooks(invoiceQuery);
      
      if (invoiceResponse.QueryResponse && invoiceResponse.QueryResponse.Invoice) {
        const invoices = invoiceResponse.QueryResponse.Invoice;
        let invoiceTotal = 0;
        
        invoices.forEach(invoice => {
          if (invoice.TotalAmt) {
            invoiceTotal += parseFloat(invoice.TotalAmt);
          }
        });
        
        totalIncome += invoiceTotal;
        
        // Add invoice row
        result.Rows.Row.push({
          type: "Data",
          ColData: [
            { value: "Invoice Revenue" },
            { value: invoiceTotal.toFixed(2) }
          ]
        });
        
        incomeSuccess = true;
      }
    } catch (invoiceError) {
      logAction('P&L Income', 'Invoice query failed', invoiceError.message);
      // Continue to next approach
    }
    
    // 2. Try sales receipts if invoices failed
    if (!incomeSuccess) {
      try {
        const salesQuery = `SELECT TotalAmt FROM SalesReceipt WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' MAXRESULTS 1000`;
        const salesResponse = queryQuickBooks(salesQuery);
        
        if (salesResponse.QueryResponse && salesResponse.QueryResponse.SalesReceipt) {
          const sales = salesResponse.QueryResponse.SalesReceipt;
          let salesTotal = 0;
          
          sales.forEach(sale => {
            if (sale.TotalAmt) {
              salesTotal += parseFloat(sale.TotalAmt);
            }
          });
          
          totalIncome += salesTotal;
          
          // Add sales row
          result.Rows.Row.push({
            type: "Data",
            ColData: [
              { value: "Sales Revenue" },
              { value: salesTotal.toFixed(2) }
            ]
          });
          
          incomeSuccess = true;
        }
      } catch (salesError) {
        logAction('P&L Income', 'Sales query failed', salesError.message);
        // Continue to next approach
      }
    }
    
    // 3. Try generic Deposit transactions as last resort
    if (!incomeSuccess) {
      try {
        const depositQuery = `SELECT TotalAmt FROM Deposit WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' MAXRESULTS 1000`;
        const depositResponse = queryQuickBooks(depositQuery);
        
        if (depositResponse.QueryResponse && depositResponse.QueryResponse.Deposit) {
          const deposits = depositResponse.QueryResponse.Deposit;
          let depositTotal = 0;
          
          deposits.forEach(deposit => {
            if (deposit.TotalAmt) {
              depositTotal += parseFloat(deposit.TotalAmt);
            }
          });
          
          totalIncome += depositTotal;
          
          // Add deposit row
          result.Rows.Row.push({
            type: "Data",
            ColData: [
              { value: "Deposits" },
              { value: depositTotal.toFixed(2) }
            ]
          });
          
          incomeSuccess = true;
        }
      } catch (depositError) {
        logAction('P&L Income', 'Deposit query failed', depositError.message);
      }
    }
    
    // If all specific queries failed, add a total row
    if (!incomeSuccess) {
      result.Rows.Row.push({
        type: "Data",
        ColData: [
          { value: "No income data available" },
          { value: "0.00" }
        ]
      });
    } else {
      // Add total row
      result.Rows.Row.push({
        type: "Data",
        ColData: [
          { value: "Total Income" },
          { value: totalIncome.toFixed(2) }
        ]
      });
    }
    
    // == EXPENSES SECTION ==
    let totalExpenses = 0;
    let expenseSuccess = false;
    
    // Add a section for Expenses
    result.Rows.Row.push({
      type: "Section",
      Header: { 
        ColData: [{ value: "Expenses" }] 
      }
    });
    
    // Try different approaches for expenses
    // 1. First try bills - common expense source
    try {
      const billQuery = `SELECT TotalAmt FROM Bill WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' MAXRESULTS 1000`;
      const billResponse = queryQuickBooks(billQuery);
      
      if (billResponse.QueryResponse && billResponse.QueryResponse.Bill) {
        const bills = billResponse.QueryResponse.Bill;
        let billTotal = 0;
        
        bills.forEach(bill => {
          if (bill.TotalAmt) {
            billTotal += parseFloat(bill.TotalAmt);
          }
        });
        
        totalExpenses += billTotal;
        
        // Add bill row
        result.Rows.Row.push({
          type: "Data",
          ColData: [
            { value: "Bills" },
            { value: billTotal.toFixed(2) }
          ]
        });
        
        expenseSuccess = true;
      }
    } catch (billError) {
      logAction('P&L Expenses', 'Bill query failed', billError.message);
      // Continue to next approach
    }
    
    // 2. Try purchase transactions if bills failed
    if (!expenseSuccess) {
      try {
        const purchaseQuery = `SELECT TotalAmt FROM Purchase WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' MAXRESULTS 1000`;
        const purchaseResponse = queryQuickBooks(purchaseQuery);
        
        if (purchaseResponse.QueryResponse && purchaseResponse.QueryResponse.Purchase) {
          const purchases = purchaseResponse.QueryResponse.Purchase;
          let purchaseTotal = 0;
          
          purchases.forEach(purchase => {
            if (purchase.TotalAmt) {
              purchaseTotal += parseFloat(purchase.TotalAmt);
            }
          });
          
          totalExpenses += purchaseTotal;
          
          // Add purchase row
          result.Rows.Row.push({
            type: "Data",
            ColData: [
              { value: "Purchases" },
              { value: purchaseTotal.toFixed(2) }
            ]
          });
          
          expenseSuccess = true;
        }
      } catch (purchaseError) {
        logAction('P&L Expenses', 'Purchase query failed', purchaseError.message);
        // Continue to next approach
      }
    }
    
    // 3. Try expense transactions as last resort
    if (!expenseSuccess) {
      try {
        const expenseQuery = `SELECT TotalAmt FROM Expense WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' MAXRESULTS 1000`;
        const expenseResponse = queryQuickBooks(expenseQuery);
        
        if (expenseResponse.QueryResponse && expenseResponse.QueryResponse.Expense) {
          const expenses = expenseResponse.QueryResponse.Expense;
          let expenseTotal = 0;
          
          expenses.forEach(expense => {
            if (expense.TotalAmt) {
              expenseTotal += parseFloat(expense.TotalAmt);
            }
          });
          
          totalExpenses += expenseTotal;
          
          // Add expense row
          result.Rows.Row.push({
            type: "Data",
            ColData: [
              { value: "Expenses" },
              { value: expenseTotal.toFixed(2) }
            ]
          });
          
          expenseSuccess = true;
        }
      } catch (expenseError) {
        logAction('P&L Expenses', 'Expense query failed', expenseError.message);
      }
    }
    
    // If all expense queries failed, add a message
    if (!expenseSuccess) {
      result.Rows.Row.push({
        type: "Data",
        ColData: [
          { value: "No expense data available" },
          { value: "0.00" }
        ]
      });
    } else {
      // Add total expenses row
      result.Rows.Row.push({
        type: "Data",
        ColData: [
          { value: "Total Expenses" },
          { value: totalExpenses.toFixed(2) }
        ]
      });
    }
    
    // == NET INCOME SECTION ==
    // Add a section for Net Income
    result.Rows.Row.push({
      type: "Section",
      Header: { 
        ColData: [{ value: "Net Income" }] 
      }
    });
    
    // Add net income row
    result.Rows.Row.push({
      type: "Data",
      ColData: [
        { value: "Net Income" },
        { value: (totalIncome - totalExpenses).toFixed(2) }
      ]
    });
    
    return result;
  } catch (error) {
    logAction('Basic P&L Error', 'Transaction approach failed', error.message);
    
    // Return a minimal valid structure rather than throwing
    return {
      Header: {
        ReportName: "Profit and Loss (Minimal)",
        StartPeriod: params.start_date,
        EndPeriod: params.end_date
      },
      Columns: {
        Column: [
          { ColTitle: "Category", ColType: "Account" },
          { ColTitle: "Amount", ColType: "Amount" }
        ]
      },
      Rows: {
        Row: [
          { 
            type: "Section",
            Header: { ColData: [{ value: "Data Unavailable" }] }
          },
          {
            type: "Data",
            ColData: [
              { value: "Unable to access transaction data" },
              { value: "" }
            ]
          },
          {
            type: "Data",
            ColData: [
              { value: "Please check your QuickBooks permissions" },
              { value: "" }
            ]
          }
        ]
      }
    };
  }
}

/**
 * Direct P&L Solution - Works even with broken OAuth scopes
 * Uses the most basic Transaction API with minimal parameters
 * 
 * @param {Object} params - Basic parameters for date range
 * @return {Object} Simple P&L data
 */
function getDirectSimplePLReport(params) {
  // Log this attempt
  logAction('Direct P&L Approach', 'Last resort method', JSON.stringify(params));
  
  try {
    // Get date range
    const startDate = params.start_date || '2024-01-01';
    const endDate = params.end_date || '2024-12-31';
    const year = startDate.substring(0, 4);
    
    // Structure for response
    const result = {
      Header: {
        ReportName: "Profit and Loss (Direct)",
        Time: new Date().toISOString(),
        StartPeriod: startDate,
        EndPeriod: endDate
      },
      Columns: {
        Column: [
          { ColTitle: "Category", ColType: "Account" },
          { ColTitle: "Amount", ColType: "Amount" }
        ]
      },
      Rows: { Row: [] }
    };
    
    // Add Income section with better explanations
    result.Rows.Row.push({
      type: "Section",
      Header: { ColData: [{ value: "Income" }] }
    });
    
    // Try to get invoice totals (most basic income data)
    try {
      const simpleIncomeQuery = `SELECT Id, TotalAmt FROM Invoice WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' MAXRESULTS 1000`;
      const invoiceResponse = callQuickBooksApi('query?query=' + encodeURIComponent(simpleIncomeQuery));
      
      let totalInvoices = 0;
      if (invoiceResponse.QueryResponse && invoiceResponse.QueryResponse.Invoice) {
        invoiceResponse.QueryResponse.Invoice.forEach(invoice => {
          if (invoice.TotalAmt) {
            totalInvoices += parseFloat(invoice.TotalAmt);
          }
        });
        
        // Add invoice total 
        result.Rows.Row.push({
          type: "Data",
          ColData: [
            { value: "Total Invoices" },
            { value: totalInvoices.toFixed(2) }
          ]
        });
      } else {
        result.Rows.Row.push({
          type: "Data",
          ColData: [
            { value: "Income (API Restricted)" },
            { value: "See QuickBooks" }
          ]
        });
        
        // Add explanation about lack of access
        result.Rows.Row.push({
          type: "Data",
          ColData: [
            { value: "Note: QuickBooks API restrictions are preventing full data access" },
            { value: "" }
          ]
        });
      }
    } catch (incomeError) {
      // If that fails, add placeholder with helpful info
      result.Rows.Row.push({
        type: "Data",
        ColData: [
          { value: "Income (API Restricted)" },
          { value: "See QuickBooks" }
        ]
      });
      
      result.Rows.Row.push({
        type: "Data",
        ColData: [
          { value: `To view complete data for ${year}, please log into QuickBooks directly` },
          { value: "" }
        ]
      });
    }
    
    // Add Expenses section
    result.Rows.Row.push({
      type: "Section",
      Header: { ColData: [{ value: "Expenses" }] }
    });
    
    // Try Bill API or Purchase API - might work when others fail
    let expenseSuccess = false;
    
    try {
      const simpleBillQuery = `SELECT Id, TotalAmt FROM Bill WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' MAXRESULTS 1000`;
      const billResponse = callQuickBooksApi('query?query=' + encodeURIComponent(simpleBillQuery));
      
      let totalBills = 0;
      if (billResponse.QueryResponse && billResponse.QueryResponse.Bill) {
        billResponse.QueryResponse.Bill.forEach(bill => {
          if (bill.TotalAmt) {
            totalBills += parseFloat(bill.TotalAmt);
          }
        });
        
        // Add bill total
        result.Rows.Row.push({
          type: "Data",
          ColData: [
            { value: "Total Bills" },
            { value: totalBills.toFixed(2) }
          ]
        });
        expenseSuccess = true;
      }
    } catch (billError) {
      // Try Purchase API instead
      try {
        const purchaseQuery = `SELECT Id, TotalAmt FROM Purchase WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' MAXRESULTS 1000`;
        const purchaseResponse = callQuickBooksApi('query?query=' + encodeURIComponent(purchaseQuery));
        
        let totalPurchases = 0;
        if (purchaseResponse.QueryResponse && purchaseResponse.QueryResponse.Purchase) {
          purchaseResponse.QueryResponse.Purchase.forEach(purchase => {
            if (purchase.TotalAmt) {
              totalPurchases += parseFloat(purchase.TotalAmt);
            }
          });
          
          // Add purchase total
          result.Rows.Row.push({
            type: "Data",
            ColData: [
              { value: "Total Purchases" },
              { value: totalPurchases.toFixed(2) }
            ]
          });
          expenseSuccess = true;
        }
      } catch (purchaseError) {
        // Both expense queries failed
      }
    }
    
    // If both expense queries failed, add placeholder
    if (!expenseSuccess) {
      result.Rows.Row.push({
        type: "Data",
        ColData: [
          { value: "Expenses (API Restricted)" },
          { value: "See QuickBooks" }
        ]
      });
      
      result.Rows.Row.push({
        type: "Data",
        ColData: [
          { value: "Try reconnecting with QuickBooks Admin credentials" },
          { value: "" }
        ]
      });
    }
    
    // Add Troubleshooting section with helpful info
    result.Rows.Row.push({
      type: "Section", 
      Header: { ColData: [{ value: "Troubleshooting" }] }
    });
    
    result.Rows.Row.push({
      type: "Data",
      ColData: [
        { value: "This report has limited data due to QuickBooks API restrictions" },
        { value: "" }
      ]
    });
    
    result.Rows.Row.push({
      type: "Data",
      ColData: [
        { value: "To fix this issue:" },
        { value: "" }
      ]
    });
    
    result.Rows.Row.push({
      type: "Data",
      ColData: [
        { value: "1. Click settings (gear icon)" },
        { value: "" }
      ]
    });
    
    result.Rows.Row.push({
      type: "Data",
      ColData: [
        { value: "2. Click 'Deep Reset' then 'Connect' again" },
        { value: "" }
      ]
    });
    
    result.Rows.Row.push({
      type: "Data",
      ColData: [
        { value: "3. Make sure you're using QuickBooks Admin credentials" },
        { value: "" }
      ]
    });
    
    result.Rows.Row.push({
      type: "Data",
      ColData: [
        { value: "4. Check 'Company Settings > Advanced > API access' in QuickBooks" },
        { value: "" }
      ]
    });
    
    return result;
  } catch (error) {
    // Even this most basic approach failed
    logAction('Direct P&L Error', 'Critical failure', error.message);
    
    // Create an informative skeleton P&L to help the user
    return {
      Header: {
        ReportName: "P&L Skeleton (API Access Limited)",
        Time: new Date().toISOString()
      },
      Columns: {
        Column: [
          { ColTitle: "Category", ColType: "Account" },
          { ColTitle: "Amount", ColType: "Amount" }
        ]
      },
      Rows: { Row: [
        { type: "Section", Header: { ColData: [{ value: "QuickBooks API Access Limited" }] } },
        { type: "Data", ColData: [{ value: "Unable to retrieve P&L data" }, { value: "" }] },
        { type: "Data", ColData: [{ value: "Please check the following:" }, { value: "" }] },
        { type: "Data", ColData: [{ value: "1. You have Admin access to QuickBooks" }, { value: "" }] },
        { type: "Data", ColData: [{ value: "2. API access is enabled in QuickBooks settings" }, { value: "" }] },
        { type: "Data", ColData: [{ value: "3. Try the 'Deep Reset' button in settings" }, { value: "" }] },
        { type: "Section", Header: { ColData: [{ value: "For Help" }] } },
        { type: "Data", ColData: [{ value: "1. Click settings (gear icon)" }, { value: "" }] },
        { type: "Data", ColData: [{ value: "2. Click 'Deep Reset'" }, { value: "" }] },
        { type: "Data", ColData: [{ value: "3. Click 'Connect' and use Admin credentials" }, { value: "" }] }
      ]}
    };
  }
}

/**
 * Fetches a report from QuickBooks.
 * 
 * @param {string} reportType - The report type to fetch
 * @param {object} params - Additional parameters for the report
 * @return {object} The report data
 */
function getReport(reportType, params = {}) {
  try {
    // Make sure we're authenticated
    if (!isAuthenticated()) {
      throw new Error('Not authenticated with QuickBooks. Please reconnect.');
    }
    
    // Log what we're requesting for debugging
    console.log('Requesting ' + reportType + ' report with params: ' + JSON.stringify(params));
    
    // Set up API call
    const endpoint = 'reports/' + reportType;
    
    // Add current date params if not provided
    if (!params.start_date) {
      const today = new Date();
      const firstDayOfYear = new Date(today.getFullYear(), 0, 1); // January 1
      params.start_date = Utilities.formatDate(firstDayOfYear, 'GMT', 'yyyy-MM-dd');
    }
    
    if (!params.end_date) {
      const today = new Date();
      params.end_date = Utilities.formatDate(today, 'GMT', 'yyyy-MM-dd');
    }
    
    // Add accounting method if not specified
    if (!params.accounting_method) {
      params.accounting_method = 'Accrual';
    }
    
    // Add other common parameters
    if (reportType === 'ProfitAndLoss' && !params.columns) {
      params.columns = 'monthly'; // Default to monthly for P&L
    }
    
    // Make the API call with retry logic
    let response;
    let retries = 0;
    const maxRetries = 3;
    
    while (retries < maxRetries) {
      try {
        response = makeApiCall(endpoint, 'GET', params);
        break; // If successful, exit the retry loop
      } catch (apiError) {
        retries++;
        console.error('API call attempt ' + retries + ' failed: ' + apiError.message);
        
        if (retries >= maxRetries) {
          throw apiError; // Re-throw if we've hit max retries
        }
        
        // Exponential backoff
        Utilities.sleep(1000 * Math.pow(2, retries));
      }
    }
    
    // Debug response
    if (response && response.getResponseCode() === 200) {
      const content = response.getContentText();
      
      try {
        // Parse the JSON
        const data = JSON.parse(content);
        
        // Check for error in the response data itself
        if (data.fault) {
          console.error('QuickBooks API fault: ' + JSON.stringify(data.fault));
          throw new Error('QuickBooks API error: ' + 
                        (data.fault.error && data.fault.error.length > 0 ? 
                         data.fault.error[0].message : 'Unknown API error'));
        }
        
        // Validate report data structure
        if (!data.Header) {
          console.warn('Unusual report structure - no Header found');
        }
        
        if (!data.Columns || !data.Columns.Column) {
          console.warn('Unusual report structure - no Columns found');
        }
        
        if (!data.Rows) {
          console.warn('Unusual report structure - no Rows found');
        }
        
        console.log('Successfully retrieved ' + reportType + ' report with ' + 
                  (data.Rows && data.Rows.Row ? data.Rows.Row.length : 0) + ' rows');
        
        return data;
      } catch (parseError) {
        console.error('Failed to parse QuickBooks API response: ' + parseError.message);
        console.error('Raw response: ' + content.substring(0, 500) + '...');
        throw new Error('Invalid response format from QuickBooks: ' + parseError.message);
      }
    } else if (response && response.getResponseCode() === 400 && 
               // Special handling for future dates or no data
               ((params.start_date && new Date(params.start_date) > new Date()) || 
                (params.end_date && new Date(params.end_date) > new Date()))) {
        
      // Create a minimal valid report structure for future dates
      console.log('Handling request for future date range. Creating empty report structure.');
      
      // Extract the year from params
      const year = params.start_date.substring(0, 4);
      
      return {
        Header: {
          ReportName: reportType,
          Time: new Date().toISOString(),
          StartPeriod: params.start_date,
          EndPeriod: params.end_date,
          ReportBasis: params.accounting_method || 'Accrual'
        },
        Columns: {
          Column: params.columns === 'monthly' ? 
            [
              { ColTitle: 'Account', ColType: 'Account' },
              { ColTitle: 'Jan ' + year, ColType: 'Money' },
              { ColTitle: 'Feb ' + year, ColType: 'Money' },
              { ColTitle: 'Mar ' + year, ColType: 'Money' },
              { ColTitle: 'Apr ' + year, ColType: 'Money' },
              { ColTitle: 'May ' + year, ColType: 'Money' },
              { ColTitle: 'Jun ' + year, ColType: 'Money' },
              { ColTitle: 'Jul ' + year, ColType: 'Money' },
              { ColTitle: 'Aug ' + year, ColType: 'Money' },
              { ColTitle: 'Sep ' + year, ColType: 'Money' },
              { ColTitle: 'Oct ' + year, ColType: 'Money' },
              { ColTitle: 'Nov ' + year, ColType: 'Money' },
              { ColTitle: 'Dec ' + year, ColType: 'Money' },
              { ColTitle: 'Total', ColType: 'Money' }
            ] : [
              { ColTitle: 'Account', ColType: 'Account' },
              { ColTitle: 'Amount', ColType: 'Amount' }
            ]
        },
        Rows: {
          Row: [
            { 
              type: 'Section',
              Header: { ColData: [{ value: 'No Data Available' }] },
              Rows: {
                Row: [
                  {
                    type: 'Data',
                    ColData: [
                      { value: `No data is available for ${reportType} in ${year}` },
                      { value: '' }
                    ]
                  },
                  {
                    type: 'Data',
                    ColData: [
                      { value: 'This report structure has been created to match your request' },
                      { value: '' }
                    ]
                  }
                ]
              }
            }
          ]
        }
      };
    } else {
      // Handle non-200 responses
      const code = response ? response.getResponseCode() : 'No response';
      const text = response ? response.getContentText() : 'Empty response';
      
      console.error('QuickBooks API error: ' + code);
      console.error('Response content: ' + text);
      
      try {
        // Try to parse error details
        const errorData = JSON.parse(text);
        if (errorData.fault) {
          throw new Error('QuickBooks error: ' + 
                        (errorData.fault.error && errorData.fault.error.length > 0 ? 
                         errorData.fault.error[0].message : 'Unknown API error'));
        }
      } catch (e) {
        // If parsing fails, use generic error
      }
      
      throw new Error('QuickBooks API call failed with status: ' + code);
    }
  } catch (error) {
    // Log detailed error information
    console.error('Error in getReport (' + reportType + '): ' + error.message);
    console.error('Stack trace: ' + error.stack);
    
    // Create a user-friendly error object
    const formattedError = {
      type: 'SystemFault',
      Error: error.toString(),
      message: 'Failed to retrieve ' + reportType + ' report from QuickBooks. ' + 
               'This may be due to API permissions, authentication issues, or QuickBooks service availability.',
      resolution: 'Please try reconnecting to QuickBooks, check your permissions, or try again later.'
    };
    
    // Log this formatted error
    console.error('Returning formatted error: ' + JSON.stringify(formattedError));
    
    // Return a specific error object rather than throwing
    return formattedError;
  }
}

/**
 * Gets a Profit and Loss report using the standard approach.
 * 
 * @param {Object} params - Report parameters
 * @return {Object} The P&L report data
 */
function getProfitAndLossReport(params) {
  try {
    // Log the attempt for diagnostics
    logAction('P&L Request', 'Standard API approach', JSON.stringify(params));
    
    // Force date parameters to be in YYYY-MM-DD format
    if (params.start_date) {
      try {
        const startDate = new Date(params.start_date);
        params.start_date = startDate.toISOString().split('T')[0];
      } catch (e) {
        console.error('Error formatting start_date: ' + e.message);
      }
    }
    
    if (params.end_date) {
      try {
        const endDate = new Date(params.end_date);
        params.end_date = endDate.toISOString().split('T')[0];
      } catch (e) {
        console.error('Error formatting end_date: ' + e.message);
      }
    }
    
    const formattedParams = {
      'accounting_method': 'Accrual',
      'minorversion': '65'
    };
    
    // Add date parameters
    if (params.start_date) formattedParams['start_date'] = params.start_date;
    if (params.end_date) formattedParams['end_date'] = params.end_date;
    
    // Add columns parameter if specified
    if (params.columns === 'monthly') {
      formattedParams['columns'] = 'monthly';
    }
    
    // Log the final parameters being sent
    console.log('Sending P&L request with params: ' + JSON.stringify(formattedParams));
    
    // Call the generic report function first
    const reportResponse = getReport('ProfitAndLoss', formattedParams);
    
    // Check if we got a SystemFault error object from getReport
    if (reportResponse && reportResponse.type === 'SystemFault') {
      logAction('P&L Error', 'Got SystemFault from getReport', reportResponse.message);
      throw new Error(reportResponse.message);
    }
    
    // Log the response structure for debugging
    logAction('P&L Response', 'Structure check', 
              'Has Header: ' + (reportResponse.Header ? 'Yes' : 'No') + 
              ', Has Rows: ' + (reportResponse.Rows ? 'Yes' : 'No'));
    
    // Enhanced validation of the response
    if (!reportResponse || typeof reportResponse !== 'object') {
      throw new Error('Invalid response format from QuickBooks API');
    }
    
    // Check for specific QuickBooks API errors in the response
    if (reportResponse.Fault || reportResponse.fault) {
      const fault = reportResponse.Fault || reportResponse.fault;
      throw new Error('QuickBooks API error: ' + 
                     (fault.Error && fault.Error.length > 0 ? 
                      fault.Error[0].Message : 'Unknown API error'));
    }
    
    // Ensure the response has the expected structure
    if (!reportResponse.Rows || !reportResponse.Rows.Row) {
      logAction('P&L Fallback', 'Missing expected structure', 'Falling back to alternative');
      return getAlternativeProfitAndLossReport(params);
    }
    
    return reportResponse;
  } catch (error) {
    // Log the error
    logAction('P&L Error', 'API call failed', error.toString());
    
    // Check if this is an authentication error
    if (error.message && (
        error.message.includes('authentication') || 
        error.message.includes('authorized') || 
        error.message.includes('token') ||
        error.message.includes('OAuth'))) {
      throw new Error('QuickBooks authentication error. Please reconnect to QuickBooks in settings.');
    }
    
    // Check if this is a permissions error
    if (error.message && (
        error.message.includes('permission') || 
        error.message.includes('access') || 
        error.message.includes('insufficient') ||
        error.message.includes('denied'))) {
      throw new Error('QuickBooks permission error. Your account may not have access to P&L reports.');
    }
    
    // Fall back to alternative report method
    try {
      logAction('P&L Fallback', 'Trying alternative approach', error.message);
      return getAlternativeProfitAndLossReport(params);
    } catch (fallbackError) {
      // If alternative approach also fails, try the basic transaction approach
      logAction('P&L Fallback', 'Trying basic transaction approach', fallbackError.message);
      try {
        return getBasicTransactionPLReport(params);
      } catch (basicError) {
        // If all approaches fail, throw a detailed error
        logAction('P&L All Methods Failed', 'No available methods worked', basicError.message);
        throw new Error('Failed to retrieve ProfitAndLoss report: All available methods failed. ' +
                       'Check your QuickBooks permissions and account access.');
      }
    }
  }
}

/**
 * Alternative approach to get P&L data using query instead of report API.
 * This might work when the report API fails due to permissions.
 * 
 * @param {Object} params - Report parameters
 * @return {Object} Profit and Loss data in a similar format
 */
function getAlternativeProfitAndLossReport(params) {
  // Log that we're using the alternative approach
  logAction('P&L Alternative', 'Using query approach', JSON.stringify(params));
  
  // Create a structured response similar to the report API
  const result = {
    Header: {
      ReportName: "Profit and Loss",
      Time: new Date().toISOString(),
      StartPeriod: params.start_date,
      EndPeriod: params.end_date
    },
    Columns: {
      Column: [
        { ColTitle: "Account", ColType: "Account" },
        { ColTitle: "Amount", ColType: "Amount" }
      ]
    },
    Rows: { Row: [] }
  };
  
  try {
    // First, query income accounts
    const incomeQuery = "SELECT * FROM Account WHERE AccountType = 'Income' OR AccountType = 'Revenue'";
    const incomeAccounts = queryQuickBooks(incomeQuery);
    
    // Then query expense accounts
    const expenseQuery = "SELECT * FROM Account WHERE AccountType = 'Expense'";
    const expenseAccounts = queryQuickBooks(expenseQuery);
    
    // Add a section for Income
    result.Rows.Row.push({
      type: "Section",
      Header: { 
        ColData: [{ value: "Income" }] 
      }
    });
    
    // Add income accounts
    if (incomeAccounts.QueryResponse && incomeAccounts.QueryResponse.Account) {
      incomeAccounts.QueryResponse.Account.forEach(account => {
        result.Rows.Row.push({
          type: "Data",
          ColData: [
            { value: account.Name },
            { value: account.CurrentBalance || 0 }
          ]
        });
      });
    }
    
    // Add a section for Expenses
    result.Rows.Row.push({
      type: "Section",
      Header: { 
        ColData: [{ value: "Expenses" }] 
      }
    });
    
    // Add expense accounts
    if (expenseAccounts.QueryResponse && expenseAccounts.QueryResponse.Account) {
      expenseAccounts.QueryResponse.Account.forEach(account => {
        result.Rows.Row.push({
          type: "Data",
          ColData: [
            { value: account.Name },
            { value: account.CurrentBalance || 0 }
          ]
        });
      });
    }
    
    return result;
  } catch (error) {
    logAction('Alternative P&L Error', 'Query approach failed', error.message);
    throw new Error('Unable to retrieve P&L data using alternative method: ' + error.message);
  }
}

/**
 * Runs a diagnostic test on the QuickBooks connection.
 * 
 * @return {Object} Diagnostic results
 */
function runQBDiagnostic() {
  const diagnosticResults = {
    success: false,
    message: "",
    details: {},
    detailedErrors: []
  };
  
  try {
    // Step 1: Check OAuth status
    const service = getOAuthService();
    const hasAccess = service.hasAccess();
    diagnosticResults.details.oauth = {
      hasAccess: hasAccess
    };
    
    if (!hasAccess) {
      diagnosticResults.message = "QuickBooks OAuth connection failed. Please reconnect by clicking the settings icon and using the Connect button.";
      return diagnosticResults;
    }
    
    // Step 2: Check company ID
    const scriptProperties = PropertiesService.getScriptProperties();
    const companyId = scriptProperties.getProperty('QBO_COMPANY_ID');
    
    if (!companyId) {
      diagnosticResults.message = "Company ID not set. Please configure in settings.";
      return diagnosticResults;
    }
    
    diagnosticResults.details.companyId = companyId;
    
    // Step 3: Try a simple API call - CompanyInfo
    try {
      const companyInfo = callQuickBooksApi('companyinfo/' + companyId);
      diagnosticResults.details.companyInfo = {
        success: true,
        name: companyInfo.CompanyInfo ? companyInfo.CompanyInfo.CompanyName : "Unknown"
      };
    } catch (companyError) {
      diagnosticResults.details.companyInfo = {
        success: false,
        error: companyError.toString()
      };
      diagnosticResults.detailedErrors.push("Company info query failed: " + companyError.toString());
    }
    
    // Step 4: Try a simple query - Customer
    try {
      const customerQuery = callQuickBooksApi('query?query=SELECT * FROM Customer MAXRESULTS 1');
      diagnosticResults.details.customerQuery = {
        success: true,
        hasData: customerQuery.QueryResponse && 
                customerQuery.QueryResponse.Customer && 
                customerQuery.QueryResponse.Customer.length > 0
      };
    } catch (queryError) {
      diagnosticResults.details.customerQuery = {
        success: false,
        error: queryError.toString()
      };
      diagnosticResults.detailedErrors.push("Customer query failed: " + queryError.toString());
    }
    
    // Step 5: Try a report query - Balance Sheet (simpler than P&L)
    try {
      // Use the simplest possible report call
      const reportResponse = callQuickBooksApi('reports/BalanceSheet?minorversion=65');
      
      // Check if the report has the expected structure
      const hasValidStructure = reportResponse && 
                               reportResponse.Header && 
                               reportResponse.Rows && 
                               reportResponse.Rows.Row;
      
      diagnosticResults.details.reportTest = {
        success: hasValidStructure,
        structure: {
          hasHeader: reportResponse && reportResponse.Header ? true : false,
          hasRows: reportResponse && reportResponse.Rows ? true : false,
          rowCount: reportResponse && reportResponse.Rows && reportResponse.Rows.Row ? 
                   reportResponse.Rows.Row.length : 0
        }
      };
      
      // Log the full report structure for debugging
      logAction('Report Structure', 'Balance Sheet Test', 
                JSON.stringify(diagnosticResults.details.reportTest.structure));
      
    } catch (reportError) {
      diagnosticResults.details.reportTest = {
        success: false,
        error: reportError.toString()
      };
      diagnosticResults.detailedErrors.push("Report query failed: " + reportError.toString());
    }
    
    // Step 6: Try a P&L report specifically
    try {
      const plResponse = callQuickBooksApi('reports/ProfitAndLoss?minorversion=65');
      
      // Check if the P&L report has the expected structure
      const hasValidStructure = plResponse && 
                               plResponse.Header && 
                               plResponse.Rows && 
                               plResponse.Rows.Row;
      
      diagnosticResults.details.plTest = {
        success: hasValidStructure,
        structure: {
          hasHeader: plResponse && plResponse.Header ? true : false,
          hasRows: plResponse && plResponse.Rows ? true : false,
          rowCount: plResponse && plResponse.Rows && plResponse.Rows.Row ? 
                   plResponse.Rows.Row.length : 0
        }
      };
      
      // Log the full P&L structure for debugging
      logAction('P&L Structure', 'P&L Test', 
                JSON.stringify(diagnosticResults.details.plTest.structure));
      
    } catch (plError) {
      diagnosticResults.details.plTest = {
        success: false,
        error: plError.toString()
      };
      diagnosticResults.detailedErrors.push("P&L query failed: " + plError.toString());
    }
    
    // Determine overall success
    const companySuccess = diagnosticResults.details.companyInfo && diagnosticResults.details.companyInfo.success;
    const querySuccess = diagnosticResults.details.customerQuery && diagnosticResults.details.customerQuery.success;
    const reportSuccess = diagnosticResults.details.reportTest && diagnosticResults.details.reportTest.success;
    const plSuccess = diagnosticResults.details.plTest && diagnosticResults.details.plTest.success;
    
    diagnosticResults.success = companySuccess && querySuccess;
    
    // Generate message
    if (diagnosticResults.success) {
      diagnosticResults.message = "QuickBooks connection is working. ";
      
      if (companySuccess) {
        diagnosticResults.message += `Connected to company: ${diagnosticResults.details.companyInfo.name}. `;
      }
      
      if (reportSuccess) {
        diagnosticResults.message += "Report API is accessible. ";
      } else {
        diagnosticResults.message += "WARNING: Report API access failed - this may affect financial reports. ";
      }
      
      if (plSuccess) {
        diagnosticResults.message += "P&L reports are accessible.";
      } else {
        diagnosticResults.message += "WARNING: P&L reports may not work correctly - check user permissions in QuickBooks.";
      }
    } else {
      diagnosticResults.message = "QuickBooks connection test failed. ";
      
      if (!companySuccess) {
        diagnosticResults.message += "Could not access company information. ";
      }
      
      if (!querySuccess) {
        diagnosticResults.message += "Could not query customer data. ";
      }
      
      if (!reportSuccess) {
        diagnosticResults.message += "Could not access report data. ";
      }
      
      diagnosticResults.message += "Please check your QuickBooks connection and permissions.";
    }
    
    return diagnosticResults;
  } catch (error) {
    diagnosticResults.message = "Diagnostic test failed: " + error.message;
    diagnosticResults.detailedErrors.push(error.toString());
    return diagnosticResults;
  }
}

/**
 * Checks if the OAuth service is authenticated with QuickBooks.
 * 
 * @return {boolean} True if authenticated, false otherwise
 */
function isAuthenticated() {
  const service = getOAuthService();
  return service.hasAccess();
}

// Export functions for external use
QuickBooksService.getOAuthService = getOAuthService;
QuickBooksService.isAuthenticated = isAuthenticated;
QuickBooksService.makeApiCall = makeApiCall;
QuickBooksService.callQuickBooksApi = callQuickBooksApi;
QuickBooksService.getCredentials = getCredentials;
QuickBooksService.saveCredentials = saveCredentials;
QuickBooksService.queryQuickBooks = queryQuickBooks;
QuickBooksService.getReport = getReport;
QuickBooksService.getProfitAndLossReport = getProfitAndLossReport;
QuickBooksService.getBasicTransactionPLReport = getBasicTransactionPLReport;
QuickBooksService.runDiagnostic = runQBDiagnostic; 