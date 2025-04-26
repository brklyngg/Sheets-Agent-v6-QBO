/**
 * OpenAI Integration Service
 * Functions for connecting to OpenAI API to interpret user queries.
 */

// OpenAI API configuration
const OPENAI_CONFIG = {
  API_URL: 'https://api.openai.com/v1/chat/completions',
  MODEL: 'gpt-4.1',  // Using the latest model, can be adjusted based on availability
  MAX_TOKENS: 10000, // Adjust based on response length needs
  TEMPERATURE: 0.1  // Lower temperature for more consistent, predictable responses
};

/**
 * Gets the OpenAI API key from script properties.
 * 
 * @return {string} The stored API key
 */
function getOpenAIApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('OPENAI_API_KEY') || '';
}

/**
 * Saves the OpenAI API key to script properties.
 * 
 * @param {string} apiKey - The OpenAI API key
 * @return {Object} Success status
 */
function saveOpenAIApiKey(apiKey) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('OPENAI_API_KEY', apiKey);
    return { success: true };
  } catch (error) {
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}

/**
 * Processes a user query with OpenAI to determine intent.
 * 
 * @param {string} query - The user's natural language query
 * @param {Array} conversationHistory - Previous messages for context
 * @param {Object} apiCallList - Available API calls for reference
 * @return {Object} The interpreted intent
 */
function processWithOpenAI(query, conversationHistory = [], apiCallList = {}) {
  const apiKey = getOpenAIApiKey();
  
  if (!apiKey) {
    throw new Error('OpenAI API key not configured. Please set it in the settings.');
  }
  
  // Prepare system message with available API calls
  const systemMessage = {
    role: 'system',
    content: `You are an expert assistant that interprets user queries about QuickBooks Online and Google Sheets data. 
Your job is to extract the intent and parameters from the user's query and translate it into concrete API actions.
Your response should be a JSON object only, without any additional text.

Available QuickBooks API calls:
${JSON.stringify(apiCallList.qbo || {}, null, 2)}

Available Google Sheets API actions:
${JSON.stringify(apiCallList.sheets || {}, null, 2)}

Based on the user query, determine:
1. The specific intent type and action
2. Which API(s) need to be called (QuickBooks, Google Sheets, or both)
3. All necessary parameters for the API calls
4. Where the results should be placed (if applicable)

Return ONLY a JSON object with this structure:
{
  "type": "fetch|create|modify|help|diagnostic|custom|unknown",
  "action": "report|query|entity|createSheet|addRow|clearRange|etc",
  "entity": "ProfitAndLoss|Invoice|Bill|etc", // Only for QuickBooks entities
  "filters": {
    "startDate": "YYYY-MM-DD",
    "endDate": "YYYY-MM-DD",
    "otherFields": "values" // Any filters for queries
  },
  "destination": "SheetName!A1:D10", // Where to put data
  "parameters": {
    // Any additional parameters needed for the action
    "sheetName": "name",
    "range": "A1:B5",
    "data": [...], // Data to insert
    "formatOptions": {...} // Formatting options
  },
  "qboApiCall": "Specific QBO API endpoint to use",
  "qboMethod": "GET|POST", // HTTP method for QBO API
  "qboData": {}, // Data for POST requests
  "sheetsApiCall": "createSheet|updateRange|formatRange|clearRange",
  "explanation": "A brief explanation of what this intent will do"
}

For the "custom" type, include specific API calls in qboApiCall and/or sheetsApiCall.
For complex tasks that require multiple steps, break it down into a primary action.
If you're unsure about specific fields, provide your best estimate or leave them blank.
Always include the "explanation" field to explain what the intent will do.`
  };
  
  // Create messages array with system message, conversation history, and current query
  let messages = [systemMessage];
  
  // Add conversation history if available (limited to last few messages for context)
  if (conversationHistory && conversationHistory.length > 0) {
    // Only add the last 5 messages to avoid token limits
    const recentHistory = conversationHistory.slice(-5);
    messages = messages.concat(recentHistory);
  }
  
  // Add the current query
  messages.push({
    role: 'user',
    content: query
  });
  
  try {
    // Make the API request to OpenAI
    const response = UrlFetchApp.fetch(OPENAI_CONFIG.API_URL, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: OPENAI_CONFIG.MODEL,
        messages: messages,
        max_tokens: OPENAI_CONFIG.MAX_TOKENS,
        temperature: OPENAI_CONFIG.TEMPERATURE,
        response_format: { type: "json_object" } // Ensure response is JSON
      }),
      muteHttpExceptions: true
    });
    
    // Parse the response
    const responseData = JSON.parse(response.getContentText());
    
    // Check for API errors
    if (responseData.error) {
      console.error('OpenAI API error: ' + JSON.stringify(responseData.error));
      throw new Error('OpenAI API error: ' + responseData.error.message);
    }
    
    // Extract the content from the response
    const content = responseData.choices[0].message.content;
    
    // Parse the JSON response
    try {
      const parsedIntent = JSON.parse(content);
      
      // Validate and normalize the intent
      return validateAndNormalizeIntent(parsedIntent, query);
    } catch (parseError) {
      console.error('Error parsing OpenAI response as JSON: ' + parseError.toString());
      console.log('Raw response: ' + content);
      throw new Error('Failed to parse OpenAI response: ' + parseError.message);
    }
  } catch (error) {
    console.error('Error calling OpenAI API: ' + error.toString());
    throw new Error('Failed to process with OpenAI: ' + error.message);
  }
}

/**
 * Validates and normalizes the intent object returned by OpenAI.
 * 
 * @param {Object} intent - The raw intent from OpenAI
 * @param {string} originalQuery - The original user query
 * @return {Object} The validated and normalized intent
 */
function validateAndNormalizeIntent(intent, originalQuery) {
  // Ensure required fields are present
  intent.type = intent.type || 'unknown';
  intent.action = intent.action || '';
  intent.entity = intent.entity || '';
  intent.filters = intent.filters || {};
  intent.parameters = intent.parameters || {};
  
  // Store the original query
  intent.text = originalQuery;
  
  // Set default values for dates if not provided but needed
  if ((intent.type === 'fetch' && intent.action === 'report') && 
      !intent.filters.startDate && !intent.filters.endDate) {
    
    // Default to current month if no dates specified
    const today = new Date();
    const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
    const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    
    intent.filters.startDate = Utilities.formatDate(firstDay, 'GMT', 'yyyy-MM-dd');
    intent.filters.endDate = Utilities.formatDate(lastDay, 'GMT', 'yyyy-MM-dd');
    
    // Add a note that default dates were applied
    intent.parameters.usedDefaultDates = true;
  }
  
  // Ensure API call fields are properly formatted
  if (intent.qboApiCall && !intent.qboApiCall.startsWith('/')) {
    intent.qboApiCall = '/' + intent.qboApiCall;
  }
  
  // For custom intents with API calls but no type, set type to custom
  if ((intent.qboApiCall || intent.sheetsApiCall) && intent.type === 'unknown') {
    intent.type = 'custom';
  }
  
  // If we have a destination but no sheetsApiCall, infer a basic update
  if (intent.destination && !intent.sheetsApiCall) {
    intent.sheetsApiCall = 'updateRange';
    intent.parameters.range = intent.destination;
  }
  
  // Set reasonable defaults for parameters based on the action
  if (intent.action === 'createSheet' && !intent.parameters.name) {
    // Try to extract a sensible sheet name from the query or entity
    const possibleName = intent.entity || 
                          (originalQuery.match(/sheet(?:\s+named)?\s+["']?([a-zA-Z0-9 ]+)["']?/i) || [])[1];
    
    intent.parameters.name = possibleName || 'New Sheet';
  }
  
  // Fill in missing explanation if needed
  if (!intent.explanation) {
    intent.explanation = generateExplanation(intent);
  }
  
  return intent;
}

/**
 * Generates an explanation of what the intent will do.
 * 
 * @param {Object} intent - The normalized intent
 * @return {string} A plain language explanation
 */
function generateExplanation(intent) {
  try {
    switch (intent.type) {
      case 'fetch':
        if (intent.action === 'report') {
          const dateRange = intent.filters.startDate && intent.filters.endDate ? 
            ` from ${intent.filters.startDate} to ${intent.filters.endDate}` : '';
          const destination = intent.destination ? ` and place it in ${intent.destination}` : '';
          return `Fetch the ${intent.entity} report${dateRange}${destination}`;
        } else if (intent.action === 'query' || intent.action === 'entity') {
          const entityName = intent.entity || 'data';
          const destination = intent.destination ? ` and place it in ${intent.destination}` : '';
          return `Fetch ${entityName} from QuickBooks${destination}`;
        }
        break;
      
      case 'create':
        if (intent.action === 'createSheet') {
          const sheetName = intent.parameters.name || 'new sheet';
          return `Create a new sheet named "${sheetName}"`;
        }
        break;
        
      case 'modify':
        if (intent.action === 'addRow') {
          const sheetName = intent.parameters.sheetName || 'the active sheet';
          return `Add a new row to ${sheetName}`;
        } else if (intent.action === 'clearRange') {
          const rangeName = intent.parameters.range || 'the specified range';
          const sheetName = intent.parameters.sheetName || 'the active sheet';
          return `Clear ${rangeName} in ${sheetName}`;
        }
        break;
        
      case 'custom':
        return `Execute a custom operation using ${intent.qboApiCall ? 'QuickBooks API' : ''}${(intent.qboApiCall && intent.sheetsApiCall) ? ' and ' : ''}${intent.sheetsApiCall ? 'Google Sheets API' : ''}`;
        
      case 'help':
        return 'Provide help information about available commands';
        
      case 'diagnostic':
        return 'Run diagnostic tests on the QuickBooks connection';
    }
    
    // Default explanation if we couldn't generate a specific one
    return 'Process the request: ' + intent.text;
  } catch (error) {
    // If anything goes wrong, return a simple explanation
    return 'Process the user query';
  }
}

/**
 * Gets a list of available API calls for reference.
 * 
 * @return {Object} Available QuickBooks and Google Sheets API calls
 */
function getAvailableApiCalls() {
  // QuickBooks API calls
  const qboApiCalls = {
    reports: {
      profitAndLoss: {
        endpoint: '/reports/ProfitAndLoss',
        description: 'Get a Profit and Loss report for a specific time period',
        parameters: {
          start_date: 'YYYY-MM-DD',
          end_date: 'YYYY-MM-DD',
          accounting_method: 'Accrual or Cash',
          columns: 'Comma-separated columns to include'
        },
        example: '/reports/ProfitAndLoss?start_date=2023-01-01&end_date=2023-03-31'
      },
      balanceSheet: {
        endpoint: '/reports/BalanceSheet',
        description: 'Get a Balance Sheet report for a specific time period',
        parameters: {
          start_date: 'YYYY-MM-DD',
          end_date: 'YYYY-MM-DD',
          accounting_method: 'Accrual or Cash',
          columns: 'Comma-separated columns to include'
        },
        example: '/reports/BalanceSheet?start_date=2023-01-01&end_date=2023-03-31'
      },
      cashFlow: {
        endpoint: '/reports/CashFlow',
        description: 'Get a Cash Flow report for a specific time period',
        parameters: {
          start_date: 'YYYY-MM-DD',
          end_date: 'YYYY-MM-DD',
          columns: 'Comma-separated columns to include'
        },
        example: '/reports/CashFlow?start_date=2023-01-01&end_date=2023-03-31'
      },
      trialBalance: {
        endpoint: '/reports/TrialBalance',
        description: 'Get a Trial Balance report',
        parameters: {
          start_date: 'YYYY-MM-DD',
          end_date: 'YYYY-MM-DD',
          accounting_method: 'Accrual or Cash',
          columns: 'Comma-separated columns to include'
        },
        example: '/reports/TrialBalance?start_date=2023-01-01&end_date=2023-03-31'
      },
      generalLedger: {
        endpoint: '/reports/GeneralLedger',
        description: 'Get a General Ledger report for a specific time period',
        parameters: {
          start_date: 'YYYY-MM-DD',
          end_date: 'YYYY-MM-DD',
          columns: 'Comma-separated columns to include'
        },
        example: '/reports/GeneralLedger?start_date=2023-01-01&end_date=2023-03-31'
      }
    },
    entities: {
      invoice: {
        endpoint: '/query?query=select * from Invoice',
        description: 'Query invoice data',
        filters: {
          'TxnDate': 'Filter by transaction date (e.g., TxnDate > \'2023-01-01\')',
          'CustomerRef': 'Filter by customer reference',
          'DueDate': 'Filter by due date',
          'TotalAmt': 'Filter by total amount',
          'Balance': 'Filter by balance'
        },
        example: '/query?query=select * from Invoice where TxnDate > \'2023-01-01\''
      },
      customer: {
        endpoint: '/query?query=select * from Customer',
        description: 'Query customer data',
        filters: {
          'DisplayName': 'Filter by customer name',
          'Active': 'Filter by active status',
          'Balance': 'Filter by balance'
        },
        example: '/query?query=select * from Customer where Active = true'
      },
      bill: {
        endpoint: '/query?query=select * from Bill',
        description: 'Query bill data',
        filters: {
          'TxnDate': 'Filter by transaction date',
          'VendorRef': 'Filter by vendor reference',
          'DueDate': 'Filter by due date',
          'TotalAmt': 'Filter by total amount'
        },
        example: '/query?query=select * from Bill where TxnDate > \'2023-01-01\''
      },
      vendor: {
        endpoint: '/query?query=select * from Vendor',
        description: 'Query vendor data',
        filters: {
          'DisplayName': 'Filter by vendor name',
          'Active': 'Filter by active status',
          'Balance': 'Filter by balance'
        },
        example: '/query?query=select * from Vendor where Active = true'
      },
      account: {
        endpoint: '/query?query=select * from Account',
        description: 'Query account data',
        filters: {
          'Name': 'Filter by account name',
          'AccountType': 'Filter by account type',
          'Active': 'Filter by active status'
        },
        example: '/query?query=select * from Account where AccountType = \'Expense\''
      },
      item: {
        endpoint: '/query?query=select * from Item',
        description: 'Query item data',
        filters: {
          'Name': 'Filter by item name',
          'Type': 'Filter by item type',
          'Active': 'Filter by active status'
        },
        example: '/query?query=select * from Item where Active = true'
      },
      payment: {
        endpoint: '/query?query=select * from Payment',
        description: 'Query payment data',
        filters: {
          'TxnDate': 'Filter by transaction date',
          'CustomerRef': 'Filter by customer reference',
          'TotalAmt': 'Filter by total amount'
        },
        example: '/query?query=select * from Payment where TxnDate > \'2023-01-01\''
      },
      purchase: {
        endpoint: '/query?query=select * from Purchase',
        description: 'Query purchase data',
        filters: {
          'TxnDate': 'Filter by transaction date',
          'VendorRef': 'Filter by vendor reference',
          'TotalAmt': 'Filter by total amount'
        },
        example: '/query?query=select * from Purchase where TxnDate > \'2023-01-01\''
      }
    }
  };
  
  // Google Sheets API actions
  const sheetsApiActions = {
    createSheet: {
      description: 'Create a new sheet with the specified name',
      parameters: {
        name: 'Name for the new sheet',
        data: 'Optional 2D array of data to populate the sheet',
        headers: 'Optional array of header values for the first row',
        fillWith: 'Optional content to fill the sheet with (e.g., "emoji")'
      }
    },
    updateRange: {
      description: 'Update values in a specified range',
      parameters: {
        range: 'A1 notation of the range (e.g., "A1:D5" or "Sheet1!A1:D5")',
        data: 'Data to write (can be array or object)',
        sheetName: 'Optional sheet name, uses active sheet if not specified'
      }
    },
    formatRange: {
      description: 'Format a specified range',
      parameters: {
        range: 'A1 notation of the range',
        format: 'Formatting options (fontWeight, fontStyle, backgroundColor, etc.)',
        sheetName: 'Optional sheet name, uses active sheet if not specified'
      }
    },
    clearRange: {
      description: 'Clear values in a specified range',
      parameters: {
        range: 'A1 notation of the range',
        sheetName: 'Optional sheet name, uses active sheet if not specified'
      }
    },
    addRow: {
      description: 'Add a new row to a sheet',
      parameters: {
        rowData: 'Array of values for the new row',
        sheetName: 'Optional sheet name, uses active sheet if not specified'
      }
    },
    getActiveSheetName: {
      description: 'Get the name of the currently active sheet',
      parameters: {}
    },
    getSheetNames: {
      description: 'Get a list of all sheet names in the spreadsheet',
      parameters: {}
    },
    getCellValue: {
      description: 'Get the value from a specific cell',
      parameters: {
        reference: 'Cell reference (e.g., "A1" or "Sheet1!A1")'
      }
    }
  };
  
  return {
    qbo: qboApiCalls,
    sheets: sheetsApiActions
  };
} 