/**
 * Utility Functions
 * Helper functions for Sheets Agent functionality.
 */

/**
 * Logs actions to a hidden sheet for tracking.
 * Creates the log sheet if it doesn't exist.
 * 
 * @param {string} action - The action type
 * @param {string} query - The user query or action details
 * @param {string} result - The outcome or error
 */
function logAction(action, query, result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet;
  
  // Try to get the log sheet, create it if it doesn't exist
  try {
    logSheet = ss.getSheetByName('Action Log');
    if (!logSheet) {
      logSheet = ss.insertSheet('Action Log');
      
      // Set up headers
      logSheet.getRange('A1:D1').setValues([['Timestamp', 'Action', 'Query/Details', 'Result/Error']]);
      logSheet.getRange('A1:D1').setFontWeight('bold');
      
      // Hide the sheet
      logSheet.hideSheet();
    }
  } catch (error) {
    console.error('Error creating log sheet:', error);
    return; // Exit if we can't log
  }
  
  // Add the log entry
  const timestamp = new Date().toISOString();
  logSheet.appendRow([timestamp, action, query, result]);
}

/**
 * Creates or gets a sheet in the active spreadsheet.
 * 
 * @param {string} sheetName - The name of the sheet
 * @param {boolean} createIfMissing - Whether to create the sheet if it doesn't exist
 * @return {Sheet} The sheet object
 */
function getOrCreateSheet(sheetName, createIfMissing = true) {
  console.log('getOrCreateSheet: Looking for sheet named "' + sheetName + '"');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.error('Failed to get active spreadsheet');
      throw new Error('Could not access active spreadsheet');
    }
    
    // Check if sheet already exists
    let sheet = ss.getSheetByName(sheetName);
    console.log('Sheet already exists: ' + (sheet ? 'Yes' : 'No'));
    
    if (!sheet && createIfMissing) {
      console.log('Creating new sheet: ' + sheetName);
      
      try {
        // Try first with insertSheet(name)
        sheet = ss.insertSheet(sheetName);
        console.log('Sheet created successfully with insertSheet(name)');
      } catch (insertError) {
        console.error('Error creating sheet with insertSheet(name): ' + insertError);
        
        // Fall back to create sheet without name then rename
        try {
          console.log('Trying fallback: create sheet without name then rename');
          sheet = ss.insertSheet();
          sheet.setName(sheetName);
          console.log('Sheet created with fallback method and renamed to: ' + sheetName);
        } catch (fallbackError) {
          console.error('Fallback sheet creation also failed: ' + fallbackError);
          throw new Error('Failed to create sheet using both methods: ' + fallbackError.message);
        }
      }
      
      // Verify sheet was created
      const verifySheet = ss.getSheetByName(sheetName);
      if (!verifySheet) {
        console.error('Sheet verification failed - sheet does not exist after creation');
        throw new Error('Sheet creation failed - sheet does not exist after creation attempt');
      }
      
      sheet = verifySheet;
    } else if (!sheet) {
      console.error('Sheet not found and createIfMissing is false');
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    
    return sheet;
  } catch (error) {
    console.error('Error in getOrCreateSheet: ' + error);
    throw new Error(`Failed to get or create sheet "${sheetName}": ${error.message}`);
  }
}

/**
 * Formats data from QuickBooks API to a 2D array for Google Sheets.
 * 
 * @param {Object} data - The data to format
 * @param {string} entityType - The type of entity (e.g., 'Invoice', 'report')
 * @return {Array} 2D array of data with headers
 */
function formatDataForSheet(data, entityType) {
  // Handle different entity types
  switch (entityType.toLowerCase()) {
    case 'report':
      return formatReportData(data);
    case 'query':
      return formatQueryData(data);
    default:
      throw new Error(`Unsupported entity type: ${entityType}`);
  }
}

/**
 * Formats QuickBooks report data for Google Sheets.
 * 
 * @param {Object} reportData - The report data from QuickBooks
 * @return {Array} 2D array of data with headers
 */
function formatReportData(reportData) {
  if (!reportData || !reportData.Rows || !reportData.Rows.Row) {
    return [['No data found']];
  }
  
  // Check if this is a basic report with fewer columns
  const isBasicReport = reportData.Header && 
                       (reportData.Header.ReportName === 'Profit and Loss (Basic)' || 
                        reportData.Header.ReportName === 'Profit and Loss (Direct)' ||
                        reportData.Header.ReportName === 'P&L Skeleton (API Access Limited)');
  
  // Get headers based on report columns or default to basic headers
  let headers;
  if (isBasicReport) {
    headers = ['Category', 'Amount'];
  } else if (reportData.Columns && reportData.Columns.Column) {
    headers = reportData.Columns.Column.map(col => col.ColTitle || col.ColType);
  } else {
    headers = ['Category', 'Amount'];
  }
  
  const result = [headers];
  const rows = reportData.Rows.Row;
  
  // Process each row
  rows.forEach(row => {
    if (row.type === 'Section') {
      // Add section header
      if (isBasicReport) {
        // For basic reports, only add 2 columns
        // Safely access ColData with null checks
        const headerValue = row.Header && row.Header.ColData && row.Header.ColData[0] ? 
                           row.Header.ColData[0].value : 'Section';
        result.push([headerValue, '']);
      } else {
        // For standard reports, match the number of headers
        // Safely access ColData with null checks
        const headerValue = row.Header && row.Header.ColData && row.Header.ColData[0] ? 
                           row.Header.ColData[0].value : 'Section';
        const sectionRow = [headerValue];
        // Fill remaining columns with empty strings
        for (let i = 1; i < headers.length; i++) {
          sectionRow.push('');
        }
        result.push(sectionRow);
      }
    } else if (row.type === 'Data') {
      // Add data row - ensure column count matches headers
      if (isBasicReport) {
        // For basic reports, only take first 2 columns or pad to 2
        const rowData = row.ColData || [];
        if (rowData.length >= 2) {
          result.push([rowData[0].value || '', rowData[1].value || '']);
        } else if (rowData.length === 1) {
          result.push([rowData[0].value || '', '']);
        } else {
          result.push(['', '']);
        }
      } else {
        // For standard reports, match all headers
        const rowData = row.ColData || [];
        const dataRow = [];
        
        // Ensure we have exactly as many columns as headers
        for (let i = 0; i < headers.length; i++) {
          dataRow.push(i < rowData.length ? (rowData[i].value || '') : '');
        }
        
        result.push(dataRow);
      }
    } else if (row.Summary) {
      // Handle summary rows
      if (isBasicReport) {
        // For basic reports, only add 2 columns
        const summaryData = row.Summary.ColData || [];
        const summaryValue = summaryData.length > 0 ? summaryData[0].value : 'Total';
        const amountValue = summaryData.length > 1 ? summaryData[1].value : '';
        result.push([summaryValue, amountValue]);
      } else {
        // For standard reports, match all headers
        const summaryData = row.Summary.ColData || [];
        const summaryRow = [];
        
        // Ensure we have exactly as many columns as headers
        for (let i = 0; i < headers.length; i++) {
          summaryRow.push(i < summaryData.length ? (summaryData[i].value || '') : '');
        }
        
        result.push(summaryRow);
      }
    }
  });
  
  return result;
}

/**
 * Formats QuickBooks query data for Google Sheets.
 * 
 * @param {Object} queryData - The query results from QuickBooks
 * @return {Array} 2D array of data with headers
 */
function formatQueryData(queryData) {
  if (!queryData || !queryData.QueryResponse) {
    return [['No data found']];
  }
  
  // Find the entity key (like 'Invoice', 'Customer', etc.)
  const entityKey = Object.keys(queryData.QueryResponse).find(key => 
    Array.isArray(queryData.QueryResponse[key]) && key !== 'maxResults'
  );
  
  if (!entityKey || !queryData.QueryResponse[entityKey].length) {
    return [['No data found']];
  }
  
  const entities = queryData.QueryResponse[entityKey];
  
  // Create headers from the first object's keys
  // Flatten nested objects by using dot notation (e.g., Customer.Name)
  const firstEntity = entities[0];
  const headers = [];
  
  function extractHeaders(obj, prefix = '') {
    for (const key in obj) {
      if (typeof obj[key] === 'object' && obj[key] !== null && !Array.isArray(obj[key])) {
        extractHeaders(obj[key], prefix + key + '.');
      } else if (!Array.isArray(obj[key])) {
        headers.push(prefix + key);
      }
    }
  }
  
  extractHeaders(firstEntity);
  
  // Create rows with the data
  const rows = entities.map(entity => {
    return headers.map(header => {
      const parts = header.split('.');
      let value = entity;
      
      for (const part of parts) {
        if (value && value[part] !== undefined) {
          value = value[part];
        } else {
          value = '';
          break;
        }
      }
      
      return value;
    });
  });
  
  return [headers, ...rows];
}

/**
 * Writes data to a Google Sheet.
 * 
 * @param {string} sheetName - The name of the sheet to write to
 * @param {Array} data - The 2D array of data to write
 * @param {boolean} clearFirst - Whether to clear the sheet first
 * @return {Object} Status of the operation
 */
function writeToSheet(sheetName, data, clearFirst = true) {
  try {
    if (!data || !data.length) {
      console.error('No data to write to sheet: ' + sheetName);
      return { success: false, message: 'No data to write' };
    }
    
    // Log sheet creation attempt
    console.log('Attempting to write to sheet: ' + sheetName);
    console.log('Data dimensions: ' + data.length + ' rows × ' + (data[0] ? data[0].length : 0) + ' columns');
    
    // Ensure all rows have the same number of columns (use the header row as reference)
    const columnCount = data[0].length;
    for (let i = 1; i < data.length; i++) {
      // If this row has fewer columns than the header, pad with empty strings
      while (data[i].length < columnCount) {
        data[i].push('');
      }
      // If this row has more columns than the header, truncate
      if (data[i].length > columnCount) {
        data[i] = data[i].slice(0, columnCount);
      }
    }
    
    // Get or create the sheet - with multiple attempts to ensure reliability
    let sheet = null;
    let creationAttempts = 0;
    const maxCreationAttempts = 3;
    
    while (!sheet && creationAttempts < maxCreationAttempts) {
      try {
        console.log(`Getting or creating sheet: ${sheetName} (attempt ${creationAttempts + 1})`);
        sheet = getOrCreateSheet(sheetName, true);
        
        // Verify sheet exists
        if (sheet) {
          console.log('Sheet object obtained successfully');
          
          // Double check the sheet name is correct
          if (sheet.getName() !== sheetName) {
            console.warn(`Sheet name mismatch. Expected: ${sheetName}, Actual: ${sheet.getName()}`);
            // Try to rename it
            sheet.setName(sheetName);
            console.log(`Renamed sheet to ${sheetName}`);
          }
        } else {
          console.error('Sheet object is null after getOrCreateSheet');
        }
      } catch (sheetError) {
        creationAttempts++;
        console.error(`Failed to create or get sheet (attempt ${creationAttempts}): ${sheetError.toString()}`);
        
        if (creationAttempts >= maxCreationAttempts) {
          return { 
            success: false, 
            message: `Failed to create sheet "${sheetName}" after ${maxCreationAttempts} attempts: ${sheetError.message}`
          };
        }
        
        // Wait a bit before retrying
        Utilities.sleep(500);
      }
    }
    
    // Verify that we have a valid sheet before continuing
    if (!sheet) {
      console.error('Sheet object is null or undefined after multiple creation attempts');
      return { 
        success: false, 
        message: `Failed to create or get sheet "${sheetName}" - sheet object is null after multiple attempts`
      };
    }
    
    // Verify sheet exists in spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetExists = ss.getSheetByName(sheetName) !== null;
    if (!sheetExists) {
      console.error('Sheet appears to be missing after creation attempt: ' + sheetName);
      
      // Try one more direct creation attempt
      try {
        console.log('Attempting final direct creation of sheet: ' + sheetName);
        sheet = ss.insertSheet(sheetName);
        console.log('Final direct creation successful');
      } catch (finalError) {
        return { 
          success: false, 
          message: `Sheet "${sheetName}" couldn't be created in the spreadsheet after multiple attempts. Try a different name.`
        };
      }
    }
    
    // Clear the sheet if requested
    if (clearFirst) {
      try {
        console.log('Clearing sheet: ' + sheetName);
        sheet.clear();
      } catch (clearError) {
        console.error('Error clearing sheet: ' + clearError.toString());
        // Continue anyway, since writing the data is the primary goal
      }
    }
    
    // Write the data
    try {
      console.log('Writing ' + data.length + ' rows to sheet');
      
      // Use batch update for better performance
      if (data.length > 0) {
        const range = sheet.getRange(1, 1, data.length, data[0].length);
        range.setValues(data);
        console.log('Data written successfully');
        
        // Auto-resize columns for better readability
        try {
          console.log('Auto-resizing columns for better display');
          sheet.autoResizeColumns(1, data[0].length);
        } catch (resizeError) {
          console.warn('Non-critical error auto-resizing columns: ' + resizeError.toString());
          // Continue anyway, this is just for aesthetics
        }
      } else {
        console.warn('No rows to write');
      }
      
      // Final verification
      const finalRowCount = sheet.getLastRow();
      console.log(`Sheet ${sheetName} now has ${finalRowCount} rows`);
      
      return {
        success: true,
        message: `Data written to sheet "${sheetName}"`,
        rowCount: data.length,
        columnCount: data[0].length
      };
    } catch (writeError) {
      console.error('Error writing data to sheet: ' + writeError.toString());
      return {
        success: false,
        message: `Failed to write data to sheet "${sheetName}": ${writeError.message}`
      };
    }
  } catch (error) {
    console.error('Unexpected error in writeToSheet: ' + error.toString());
    return {
      success: false,
      message: `Unexpected error: ${error.message}`
    };
  }
}

/**
 * Creates a new sheet with the given name.
 * 
 * @param {string} sheetName - The name for the new sheet
 * @param {Object} options - Additional options (e.g., position, template)
 * @return {Object} Result with success status and the new sheet
 */
function createSheet(sheetName, options = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if a sheet with this name already exists
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      return { 
        success: false, 
        message: `A sheet named "${sheetName}" already exists.`,
        sheet: sheet
      };
    }
    
    // Create the new sheet
    sheet = ss.insertSheet(sheetName);
    
    // Apply any template formatting if specified
    if (options.template) {
      applySheetTemplate(sheet, options.template);
    }
    
    // Fill with data if provided
    if (options.data) {
      if (Array.isArray(options.data) && options.data.length > 0) {
        const numRows = options.data.length;
        const numCols = Array.isArray(options.data[0]) ? options.data[0].length : 1;
        
        if (numRows > 0 && numCols > 0) {
          sheet.getRange(1, 1, numRows, numCols).setValues(options.data);
        }
      }
    }
    
    // Format headers if provided
    if (options.headers && Array.isArray(options.headers) && options.headers.length > 0) {
      sheet.getRange(1, 1, 1, options.headers.length).setValues([options.headers]).setFontWeight('bold');
    }
    
    // Fill with special content if specified (e.g., emojis)
    if (options.fillWith) {
      fillSheetWithContent(sheet, options.fillWith);
    }
    
    return { 
      success: true, 
      message: `Created new sheet named "${sheetName}".`,
      sheet: sheet
    };
  } catch (error) {
    console.error('Error creating sheet:', error);
    return { 
      success: false, 
      message: `Failed to create sheet: ${error.message}`
    };
  }
}

/**
 * Updates a range in a sheet with the provided data.
 * 
 * @param {string} range - The A1 notation of the range
 * @param {Array|Object} data - The data to write (2D array or object)
 * @param {string} sheetName - Optional sheet name, uses active sheet if not provided
 * @return {Object} Result with success status
 */
function updateSheetRange(range, data, sheetName = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
    
    // Convert data to 2D array if it's not already
    let values;
    if (Array.isArray(data)) {
      // Ensure we have a 2D array
      if (!Array.isArray(data[0])) {
        values = [data]; // Convert 1D array to 2D
      } else {
        values = data;
      }
    } else if (typeof data === 'object') {
      // Convert object to 2D array for display
      values = objectToSheetData(data);
    } else {
      // Handle primitive values
      values = [[data]];
    }
    
    // Get the dimensions of the data
    const numRows = values.length;
    const numCols = values[0].length;
    
    // If range is just a cell reference (e.g., "A1"), expand it to cover the data
    let targetRange;
    if (/^[A-Z]+\d+$/.test(range)) {
      // It's a single cell reference
      const rangeA1 = sheet.getRange(range);
      const startRow = rangeA1.getRow();
      const startCol = rangeA1.getColumn();
      targetRange = sheet.getRange(startRow, startCol, numRows, numCols);
    } else {
      // It's a range specification
      targetRange = sheet.getRange(range);
    }
    
    // Write the data
    targetRange.setValues(values);
    
    return {
      success: true,
      message: `Updated range ${range} in sheet "${sheet.getName()}".`,
      updatedRange: targetRange.getA1Notation()
    };
  } catch (error) {
    console.error('Error updating sheet range:', error);
    return {
      success: false,
      message: `Failed to update range: ${error.message}`
    };
  }
}

/**
 * Clears a range in a sheet.
 * 
 * @param {string} range - The A1 notation of the range
 * @param {string} sheetName - Optional sheet name, uses active sheet if not provided
 * @return {Object} Result with success status
 */
function clearSheetRange(range, sheetName = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
    
    // If range is specified, clear that range, otherwise clear the entire sheet
    if (range) {
      sheet.getRange(range).clear();
    } else {
      sheet.clear();
    }
    
    return {
      success: true,
      message: range ? 
        `Cleared range ${range} in sheet "${sheet.getName()}".` : 
        `Cleared all data in sheet "${sheet.getName()}".`
    };
  } catch (error) {
    console.error('Error clearing sheet range:', error);
    return {
      success: false,
      message: `Failed to clear range: ${error.message}`
    };
  }
}

/**
 * Formats a range in a sheet with the provided formatting options.
 * 
 * @param {string} range - The A1 notation of the range
 * @param {Object} format - The formatting options
 * @param {string} sheetName - Optional sheet name, uses active sheet if not provided
 * @return {Object} Result with success status
 */
function formatSheetRange(range, format, sheetName = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
    
    const targetRange = sheet.getRange(range);
    
    // Apply formatting options
    if (format) {
      if (format.fontWeight) targetRange.setFontWeight(format.fontWeight);
      if (format.fontStyle) targetRange.setFontStyle(format.fontStyle);
      if (format.fontSize) targetRange.setFontSize(format.fontSize);
      if (format.fontColor) targetRange.setFontColor(format.fontColor);
      if (format.backgroundColor) targetRange.setBackground(format.backgroundColor);
      if (format.horizontalAlignment) targetRange.setHorizontalAlignment(format.horizontalAlignment);
      if (format.verticalAlignment) targetRange.setVerticalAlignment(format.verticalAlignment);
      if (format.border) {
        const {top, bottom, left, right} = format.border;
        if (top) targetRange.setBorder(true, null, null, null, null, null, top.color, top.style);
        if (bottom) targetRange.setBorder(null, null, true, null, null, null, bottom.color, bottom.style);
        if (left) targetRange.setBorder(null, null, null, true, null, null, left.color, left.style);
        if (right) targetRange.setBorder(null, null, null, null, true, null, right.color, right.style);
      }
      if (format.numberFormat) targetRange.setNumberFormat(format.numberFormat);
    }
    
    return {
      success: true,
      message: `Formatted range ${range} in sheet "${sheet.getName()}".`
    };
  } catch (error) {
    console.error('Error formatting sheet range:', error);
    return {
      success: false,
      message: `Failed to format range: ${error.message}`
    };
  }
}

/**
 * Writes data to a specific range in a sheet.
 * 
 * @param {string} rangeA1 - The A1 notation of the range
 * @param {Array|Object} data - The data to write
 * @return {Object} Result with success status
 */
function writeDataToRange(rangeA1, data) {
  try {
    // Parse the range to extract sheet name and cell range
    const match = rangeA1.match(/^(?:'([^']+)'|([^!]+))!(.+)$/);
    
    let sheetName, range;
    if (match) {
      // Range includes sheet name (e.g. 'Sheet1'!A1:B10)
      sheetName = match[1] || match[2];
      range = match[3];
    } else {
      // Range is just cells, use active sheet (e.g. A1:B10)
      range = rangeA1;
    }
    
    return updateSheetRange(range, data, sheetName);
  } catch (error) {
    console.error('Error writing data to range:', error);
    return {
      success: false,
      message: `Failed to write data to range: ${error.message}`
    };
  }
}

/**
 * Converts an object to a 2D array suitable for writing to a sheet.
 * 
 * @param {Object} data - The object to convert
 * @return {Array} 2D array representation
 */
function objectToSheetData(data) {
  if (!data) return [['No data']];
  
  if (Array.isArray(data)) {
    if (data.length === 0) return [['No data']];
    
    if (typeof data[0] === 'object' && !Array.isArray(data[0])) {
      // Array of objects
      const headers = Object.keys(data[0]);
      const rows = data.map(item => headers.map(header => item[header] || ''));
      return [headers, ...rows];
    } else {
      // Array of primitives or arrays
      return Array.isArray(data[0]) ? data : [data];
    }
  } else if (typeof data === 'object') {
    // Single object
    const headers = Object.keys(data);
    const values = headers.map(key => data[key]);
    return [headers, values];
  }
  
  // Fall back to simple representation
  return [[JSON.stringify(data)]];
}

/**
 * Fills a sheet with specified content (e.g., emoji patterns).
 * 
 * @param {Sheet} sheet - The sheet to fill
 * @param {string} contentType - Type of content to fill with
 */
function fillSheetWithContent(sheet, contentType) {
  // Default dimensions
  const rows = 20;
  const cols = 10;
  
  // Create different patterns based on content type
  let content;
  
  switch (contentType.toLowerCase()) {
    case 'poop emoji':
    case 'poop':
      content = Array(rows).fill().map(() => Array(cols).fill('💩'));
      break;
    case 'heart emoji':
    case 'heart':
      content = Array(rows).fill().map(() => Array(cols).fill('❤️'));
      break;
    case 'star emoji':
    case 'star':
      content = Array(rows).fill().map(() => Array(cols).fill('⭐'));
      break;
    case 'smile emoji':
    case 'smile':
      content = Array(rows).fill().map(() => Array(cols).fill('😊'));
      break;
    case 'cat emoji':
    case 'cat':
      content = Array(rows).fill().map(() => Array(cols).fill('🐱'));
      break;
    case 'pattern':
      // Create a checkerboard pattern
      content = Array(rows).fill().map((_, rowIdx) => 
        Array(cols).fill().map((_, colIdx) => 
          (rowIdx + colIdx) % 2 === 0 ? '⬛' : '⬜'
        )
      );
      break;
    default:
      // Default to a single emoji
      content = Array(rows).fill().map(() => Array(cols).fill('👍'));
  }
  
  // Write the content to the sheet
  sheet.getRange(1, 1, rows, cols).setValues(content);
}

/**
 * Formats QuickBooks data specifically for Google Sheets.
 * 
 * @param {Object} data - The QBO API response data
 * @param {Object} options - Formatting options
 * @return {Array} 2D array ready for sheet insertion
 */
function formatDataForSheet(data, options = {}) {
  if (!data) return [['No data received']];
  
  // Add debug logging
  console.log('Formatting data of type: ' + (typeof options === 'string' ? options : 'object'));
  
  // Handle different data structures from QBO API
  if (data.QueryResponse) {
    // It's a query response
    const entityName = Object.keys(data.QueryResponse).find(key => Array.isArray(data.QueryResponse[key]));
    
    if (entityName && data.QueryResponse[entityName]) {
      const entities = data.QueryResponse[entityName];
      if (entities.length === 0) return [['No data found for this query']];
      
      // Get headers from first entity
      const headers = Object.keys(entities[0]);
      
      // Create rows
      const rows = entities.map(entity => 
        headers.map(header => {
          const value = entity[header];
          return typeof value === 'object' ? JSON.stringify(value) : value;
        })
      );
      
      return [headers, ...rows];
    }
  } else if (data.Rows && data.Columns) {
    // It's a report response
    console.log('Processing report with columns: ' + (data.Columns.Column ? data.Columns.Column.length : 'none'));
    
    // Determine if it's a P&L report with monthly columns
    const isProfitAndLoss = data.Header && data.Header.ReportName && 
                          data.Header.ReportName.includes('Profit and Loss');
    const hasMonthlyColumns = data.Columns && data.Columns.Column && 
                            data.Columns.Column.length > 2;
    
    console.log('Report type: ' + (isProfitAndLoss ? 'P&L' : 'Other') + 
               ', Has monthly columns: ' + (hasMonthlyColumns ? 'Yes' : 'No'));
    
    // Get headers from columns
    let headers = [];
    if (data.Columns && data.Columns.Column) {
      headers = data.Columns.Column.map(col => col.ColTitle || col.ColType || 'Column');
      console.log('Found ' + headers.length + ' columns: ' + headers.join(', '));
    } else {
      headers = ['Account', 'Amount'];
      console.log('No columns found, using default headers');
    }
    
    // Process rows recursively to handle nested structures
    const processRow = (row) => {
      if (!row) {
        console.log('Empty row encountered');
        return [];
      }
      
      // Check for section rows with children
      if (row.Rows && row.Rows.Row) {
        console.log('Processing section with ' + row.Rows.Row.length + ' child rows');
        
        // First add the section header
        const sectionRows = [];
        
        // Add section header if available
        if (row.Header) {
          const headerData = row.Header.ColData || [];
          const headerValues = headerData.map(col => col.value || '');
          
          // Pad with empty strings to match header length
          while (headerValues.length < headers.length) {
            headerValues.push('');
          }
          
          // Add the section header row
          sectionRows.push(headerValues);
        }
        
        // Process child rows and add them
        const childRows = row.Rows.Row.flatMap(processRow);
        
        // Return combined section header and child rows
        return [...sectionRows, ...childRows];
      }
      
      // This is a data row - process it directly
      if (row.ColData) {
        const values = row.ColData.map(col => col.value || '');
        
        // Pad with empty strings to match header length
        while (values.length < headers.length) {
          values.push('');
        }
        
        return [values];
      }
      
      // Handle summary rows
      if (row.Summary) {
        const summaryData = row.Summary.ColData || [];
        const summaryValues = summaryData.map(col => col.value || '');
        
        // Pad with empty strings to match header length
        while (summaryValues.length < headers.length) {
          summaryValues.push('');
        }
        
        return [summaryValues];
      }
      
      console.log('Unhandled row type encountered');
      return [];
    };
    
    // Process all rows
    let rows = [];
    if (data.Rows && data.Rows.Row && data.Rows.Row.length > 0) {
      console.log('Processing ' + data.Rows.Row.length + ' top-level rows');
      rows = data.Rows.Row.flatMap(processRow);
      console.log('Generated ' + rows.length + ' final data rows');
    } else {
      console.log('No rows found in report data');
    }
    
    // If we have headers but no rows, add a note row
    if (rows.length === 0) {
      const emptyRow = Array(headers.length).fill('');
      emptyRow[0] = 'No data available for this period';
      rows.push(emptyRow);
    }
    
    return [headers, ...rows];
  } else if (Array.isArray(data)) {
    // It's already an array
    if (data.length === 0) return [['No data']];
    
    if (typeof data[0] === 'object' && !Array.isArray(data[0])) {
      // Array of objects
      const headers = Object.keys(data[0]);
      const rows = data.map(item => 
        headers.map(header => {
          const value = item[header];
          return typeof value === 'object' ? JSON.stringify(value) : value;
        })
      );
      
      return [headers, ...rows];
    } else {
      // Simple array
      return [data];
    }
  }
  
  // Fall back to simple JSON representation
  return [['Data'], [JSON.stringify(data)]];
} 