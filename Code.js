/**
 * QBO Agent
 * A Google Sheets add-on that enables QuickBooks data integration via natural language.
 */

/**
 * Creates the menu items when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open QBO Agent V7', 'showSidebar')
    .addToUi();
  // Automatically open the sidebar on open
  showSidebar();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall() {
  onOpen();
}

/**
 * Opens the sidebar with the QBO Agent interface.
 */
function showSidebar() {
  const ui = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('QBO Agent V7')
    .setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Includes an HTML file from the project.
 * 
 * @param {string} filename - The file name without the .html extension
 * @return {string} The contents of the file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Returns the user's OAuth token for the QuickBooks API.
 * 
 * @return {Object} The OAuth token status and value if available
 */
function getOAuthToken() {
  const service = getOAuthService();
  const hasAccess = service.hasAccess();
  
  return {
    hasAccess: hasAccess,
    authorizationUrl: hasAccess ? null : service.getAuthorizationUrl(),
    token: hasAccess ? service.getAccessToken() : null
  };
}

/**
 * Handles the OAuth callback.
 * 
 * @param {Object} request - The request parameters
 * @return {HtmlOutput} Confirmation page
 */
function authCallback(request) {
  // Log the full request object to see what's happening
  Logger.log("Auth callback request: " + JSON.stringify(request));
  
  const service = getOAuthService();
  const authorized = service.handleCallback(request);
  
  if (authorized) {
    return HtmlService.createHtmlOutput(
      '<h3>Success!</h3>' +
      '<p>You have successfully connected to QuickBooks.</p>' +
      '<p>You can close this tab and return to QBO Agent.</p>' +
      '<script>setTimeout(function() { window.close(); }, 3000);</script>'
    );
  } else {
    return HtmlService.createHtmlOutput(
      '<h3>Authorization Failed</h3>' +
      '<p>There was a problem connecting to QuickBooks.</p>' +
      '<p>Please try again or contact support if the issue persists.</p>'
    );
  }
}

/**
 * Logs the user out of QuickBooks.
 * 
 * @return {Object} Success status
 */
function logoutQuickBooks() {
  try {
    const service = getOAuthService();
    service.reset();
    
    // Also clear any stored tokens that might have outdated permissions
    PropertiesService.getUserProperties().deleteProperty('oauth2.quickbooks');
    
    return { 
      success: true,
      message: "Successfully disconnected from QuickBooks. Please reconnect to grant the necessary permissions."
    };
  } catch (error) {
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}

/**
 * Processes a user query using the AI-driven chat interface.
 * 
 * @param {string} query - The user's natural language query
 * @param {Array} conversationHistory - Previous messages in the conversation
 * @return {Object} The processed response
 */
function processQuery(query, conversationHistory = []) {
  try {
    // Log the query for future reference
    logAction('User Query', query, 'N/A');
    
    // Process the query through the NLP module
    const intent = analyzeQueryIntent(query, conversationHistory);
    
    // Execute the appropriate action based on the intent
    const response = executeIntent(intent);
    
    // Log successful action
    logAction('System Response', 'Success', JSON.stringify(response));
    
    return response;
  } catch (error) {
    // Log error
    logAction('Error', query, error.toString());
    
    return {
      type: 'error',
      message: 'Sorry, I encountered an error: ' + error.message,
      data: null
    };
  }
}

/**
 * Completely resets all OAuth tokens and settings.
 * Use this as a last resort when authorization issues persist.
 * 
 * @return {Object} Status message
 */
function resetAllAuth() {
  try {
    // Get all the property stores
    const userProps = PropertiesService.getUserProperties();
    const scriptProps = PropertiesService.getScriptProperties();
    
    // Clear all OAuth-related properties - be more thorough
    userProps.deleteProperty('oauth2.quickbooks');
    userProps.deleteProperty('oauth2.quickbooks.refresh');
    userProps.deleteProperty('oauth2.quickbooks.id_token');
    userProps.deleteProperty('oauth2.quickbooks.last_code');
    userProps.deleteProperty('oauth2.quickbooks.access_token');
    userProps.deleteProperty('oauth2.quickbooks.expires_in');
    userProps.deleteProperty('oauth2.quickbooks.issued_at');
    
    // Get the service and reset it
    const service = getOAuthService();
    service.reset();
    
    // Only keep the client credentials
    const clientId = scriptProps.getProperty('QBO_CLIENT_ID');
    const clientSecret = scriptProps.getProperty('QBO_CLIENT_SECRET');
    const companyId = scriptProps.getProperty('QBO_COMPANY_ID');
    
    // Clear all stored properties
    scriptProps.deleteAllProperties();
    
    // Restore only the client credentials
    if (clientId) scriptProps.setProperty('QBO_CLIENT_ID', clientId);
    if (clientSecret) scriptProps.setProperty('QBO_CLIENT_SECRET', clientSecret);
    if (companyId) scriptProps.setProperty('QBO_COMPANY_ID', companyId);
    
    // Log the reset
    logAction('Auth Reset', 'Complete OAuth reset', 'All tokens cleared');
    
    return {
      success: true,
      message: "Authorization completely reset. Please reconnect to QuickBooks by clicking the Connect button in settings."
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Runs a diagnostic test on the QuickBooks connection.
 * 
 * @return {Object} Formatted diagnostic results
 */
function runQBDiagnostic() {
  try {
    // Call the diagnostic function in QuickBooksService.js
    const results = runQBDiagnostic();
    
    // Format the results for display
    let message = "QuickBooks Diagnostic Results:\n\n";
    
    // Add the main status message
    message += results.message + "\n\n";
    
    // Add connection details
    message += "Connection Status:\n";
    message += "OAuth: " + (results.details.oauth.hasAccess ? "✅ Connected" : "❌ Not Connected") + "\n";
    message += "Company ID: " + (results.details.companyId ? "✅ Set" : "❌ Missing") + "\n";
    message += "Company Info: " + 
               (results.details.companyInfo && results.details.companyInfo.success ? "✅ Success" : "❌ Failed") + "\n";
    message += "Customer Query: " + 
               (results.details.customerQuery && results.details.customerQuery.success ? "✅ Success" : "❌ Failed") + "\n";
    message += "Report Access: " + 
               (results.details.reportTest && results.details.reportTest.success ? "✅ Success" : "❌ Failed") + "\n\n";
    
    // Add P&L specific info
    message += "P&L Report Status: " + 
               (results.details.plTest && results.details.plTest.success ? "✅ Success" : "❌ Failed") + "\n\n";
    
    // Add detailed errors if any
    if (results.detailedErrors && results.detailedErrors.length > 0) {
      message += "Detailed Errors:\n";
      results.detailedErrors.forEach(error => {
        message += "- " + error + "\n";
      });
    }
    
    // Add troubleshooting tips
    message += "\nTroubleshooting Tips:\n";
    message += "1. If OAuth is not connected, click the settings icon and use the Connect button.\n";
    message += "2. If reports are failing, ensure your QuickBooks user has Reports access.\n";
    message += "3. For P&L issues, try using an Admin account to connect.\n";
    
    return {
      type: 'diagnostic',
      message: message,
      success: results.success
    };
  } catch (error) {
    return {
      type: 'error',
      message: 'Error running diagnostics: ' + error.message
    };
  }
}

/**
 * Saves the OpenAI API key to script properties.
 * 
 * @param {Object} openaiSettings - Object containing the OpenAI API key
 * @return {Object} Success status
 */
function saveOpenAISettings(openaiSettings) {
  try {
    console.log('Saving OpenAI settings: ' + JSON.stringify(openaiSettings));
    
    if (!openaiSettings || !openaiSettings.apiKey) {
      return { 
        success: false, 
        error: "API key is required" 
      };
    }
    
    // Save the API key to script properties
    const result = saveOpenAIApiKey(openaiSettings.apiKey);
    console.log('Save result: ' + JSON.stringify(result));
    
    return result;
  } catch (error) {
    console.error('Error saving OpenAI settings: ' + error.toString());
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}

/**
 * Gets the current OpenAI settings.
 * 
 * @return {Object} Object containing the API key status
 */
function getOpenAISettings() {
  try {
    const apiKey = getOpenAIApiKey();
    console.log('OpenAI API key exists: ' + !!apiKey);
    
    return {
      hasApiKey: !!apiKey,
      // We don't return the actual API key for security reasons
      // Only whether it's configured or not
    };
  } catch (error) {
    console.error('Error getting OpenAI settings: ' + error.toString());
    return { 
      hasApiKey: false, 
      error: error.toString() 
    };
  }
}

/**
 * Runs a test of the OpenAI integration.
 * 
 * @return {Object} Test results
 */
function testOpenAIConnection() {
  try {
    console.log('Testing OpenAI connection');
    const apiKey = getOpenAIApiKey();
    
    if (!apiKey) {
      return {
        success: false,
        message: 'OpenAI API key not configured. Please set it in the settings.'
      };
    }
    
    // Run a simple test query
    const testIntent = processWithOpenAI(
      "Hello, this is a test query to check if OpenAI is working.",
      [],
      getAvailableApiCalls()
    );
    
    console.log('Test result: ' + JSON.stringify(testIntent));
    
    if (testIntent) {
      return {
        success: true,
        message: 'OpenAI connection test successful!',
        details: { 
          model: OPENAI_CONFIG.MODEL,
          test_response_received: true 
        }
      };
    } else {
      return {
        success: false,
        message: 'OpenAI test failed. Received empty response.'
      };
    }
  } catch (error) {
    console.error('OpenAI test error: ' + error.toString());
    return {
      success: false,
      message: 'OpenAI test failed with error: ' + error.message,
      error: error.toString()
    };
  }
} 