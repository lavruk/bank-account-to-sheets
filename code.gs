/*
  =============================================================================
  Project Page: https://github.com/cmenon12/bank-account-to-sheets
  Copyright:    (c) 2021 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Make a request to the URL using the params.
 *
 * @param {string} url the URL to make the request to.
 * @param {Object} params the params to use with the request.
 * @return {string} the text of the response if successful.
 * @throws {Error} response status code was not 200.
 */
function makeRequest(url, params) {

  // Make the POST request
  const response = UrlFetchApp.fetch(url, params);
  const status = response.getResponseCode();
  const responseText = response.getContentText();

  // If successful then return the response text
  if (status === 200) {
    return responseText;

    // Otherwise log and throw an error
  } else {
    Logger.log(`There was a ${status} error fetching ${url}.`);
    Logger.log(responseText);
    throw Error(`There was a ${status} error fetching ${url}.`);
  }

}


/**
 * Exchanges a public token for an access token (internal logic).
 * @param {string} publicToken The public token to exchange.
 */
function _exchangePublicTokenInternal(publicToken) {
    // Get or create the Plaid sheet
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plaid");
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Plaid");
    }

    // Ensure headers are present
    const headers = ["link_token_response", "link_get_response", "last_cursor", "item_id", "access_token"];
    for (let i = 0; i < headers.length; i++) {
        if (sheet.getRange(1, i + 1).getValue() !== headers[i]) {
            sheet.getRange(1, i + 1).setValue(headers[i]);
        }
    }

    const request = {
      client_id: getSecrets().CLIENT_ID,
      secret: getSecrets().SECRET,
      public_token: publicToken,
    };

    const params = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(request),
      muteHttpExceptions: true,
    };

    const responseText = makeRequest(`${getSecrets().URL}/item/public_token/exchange`, params);
    const data = JSON.parse(responseText);

    if (data.access_token && data.item_id) {
      // Store the new credentials in the last row
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 4).setValue(data.item_id);
      sheet.getRange(lastRow, 5).setValue(data.access_token);

      SpreadsheetApp.getActiveSpreadsheet().toast(`Successfully exchanged public token. Access token stored.`);
      Logger.log(`Successfully exchanged public token for item_id: ${data.item_id}`);
    } else {
      Logger.log("Failed to exchange public token. Response did not contain access_token and item_id.");
      Logger.log(JSON.stringify(data, null, 2));
      SpreadsheetApp.getActiveSpreadsheet().toast("Failed to exchange public token. Please check the logs.");
      throw new Error("Failed to exchange public token.");
    }
}

/**
 * Gets information about a link token.
 */
function getLinkTokenInfo() {
  try {
    // Get the Plaid sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plaid");
    if (!sheet) {
      SpreadsheetApp.getActiveSpreadsheet().toast("The 'Plaid' sheet does not exist. Please create a link token first.");
      return;
    }

    // Get the last link_token from the sheet
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast("The 'Plaid' sheet is empty. Please create a link token first.");
      return;
    }
    const linkTokenData = JSON.parse(sheet.getRange(lastRow, 1).getValue());
    const linkToken = linkTokenData.link_token;

    if (!linkToken) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Could not find a link_token in the last row of the 'Plaid' sheet.");
      return;
    }

    // Prepare the request body
    const body = {
      "client_id": getSecrets().CLIENT_ID,
      "secret": getSecrets().SECRET,
      "link_token": linkToken,
    };

    // Condense the above into a single object
    const params = {
      "contentType": "application/json",
      "method": "post",
      "payload": JSON.stringify(body),
      "muteHttpExceptions": true
    };

    // Make the POST request
    const responseText = makeRequest(`${getSecrets().URL}/link/token/get`, params);
    const result = JSON.parse(responseText);

    Logger.log('Full response from /link/token/get:');
    Logger.log(JSON.stringify(result, null, 2));

    // Automatically exchange public token if one is available
    if (result && result.link_sessions && result.link_sessions.length > 0) {
      for (const session of result.link_sessions) {
        if (session.results && session.results.item_add_results && session.results.item_add_results.length > 0 && session.results.item_add_results[0].public_token) {
          const publicToken = session.results.item_add_results[0].public_token;
          Logger.log("Public token found in /link/token/get response. Exchanging automatically.");
          _exchangePublicTokenInternal(publicToken);
          // We'll just exchange the first one we find and then break.
          break;
        }
      }
    }

    // Check for the results object from a completed Link flow and store it if present
    if (result && result.link_sessions && result.link_sessions.length > 0 && result.link_sessions[0].results) {
      const resultsObject = result.link_sessions[0].results;
      sheet.getRange(lastRow, 2).setValue(JSON.stringify(resultsObject, null, 2));
      SpreadsheetApp.getActiveSpreadsheet().toast(`Successfully retrieved and stored the 'results' object in column B.`);
    } else {
      Logger.log("The '/link/token/get' response did not contain a 'results' object. Nothing was written to the sheet.");
      SpreadsheetApp.getActiveSpreadsheet().toast("Link token info retrieved, but no 'results' object was found to store.");
    }

  } catch (e) {
    Logger.log('An error occurred in getLinkTokenInfo:');
    Logger.log(e);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Failed to get link token info. Check logs for details. Error: ${e.message}`);
  }
}


/**
 * Downloads and returns all transactions from Plaid.
 * 
 * @return {Object} the result of transactions.get, with all transactions.
 */
function syncTransactionsFromPlaid() {
  // Get the Plaid sheet for storing the cursor and access token
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plaid");
  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("The 'Plaid' sheet does not exist. Please create a link token and exchange it first.");
    throw new Error("The 'Plaid' sheet does not exist.");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { // Headers are in row 1
      SpreadsheetApp.getActiveSpreadsheet().toast("No data found in the 'Plaid' sheet.");
      throw new Error("No data in 'Plaid' sheet.");
  }

  let cursor = sheet.getRange(lastRow, 3).getValue() || null;
  let accessToken = sheet.getRange(lastRow, 5).getValue();

  if (!accessToken) {
      SpreadsheetApp.getActiveSpreadsheet().toast("No Plaid access token found. Please link an account first or set it in the script properties.");
      throw new Error("Plaid access token not found.");
  }

  let added = [];
  let modified = [];
  let removed = [];
  let accounts = [];
  let hasMore = true;

  try {
    // Iterate through each page of new transaction updates for item
    while (hasMore) {
      const request = {
        client_id: getSecrets().CLIENT_ID,
        secret: getSecrets().SECRET,
        access_token: accessToken,
        cursor: cursor,
      };

      const params = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(request),
        muteHttpExceptions: true,
      };

      const responseText = makeRequest(`${getSecrets().URL}/transactions/sync`, params);
      const data = JSON.parse(responseText);

      if (data.error_code) {
          throw new Error(`Plaid API Error: ${data.error_code} - ${data.error_message}`);
      }

      // On the first page of the response, get the accounts.
      if (accounts.length === 0) {
        accounts = data.accounts || [];
      }

      added = added.concat(data.added || []);
      modified = modified.concat(data.modified || []);
      removed = removed.concat(data.removed || []);
      hasMore = data.has_more;
      cursor = data.next_cursor;
    }

    // Persist cursor and updated data
    sheet.getRange(lastRow, 3).setValue(cursor);
    Logger.log(`Sync complete. Next cursor stored: ${cursor}`);

    return { added, modified, removed, accounts };

  } catch (e) {
    Logger.log(`An error occurred during transaction sync: ${e.message}`);
    Logger.log(e);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error during transaction sync: ${e.message}`);
    throw e;
  }
}


/**
 * Updates the transactions in the Transactions sheet.
 */
function updateTransactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  const existingTransactions = getTransactionsFromSheet(sheet);
  const plaid = syncTransactionsFromPlaid();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Handle removed transactions
  const rowsToDelete = plaid.removed.map(removed => {
    const existing = existingTransactions[removed.transaction_id];
    return existing ? existing.rowNumber : null;
  }).filter(rowNum => rowNum !== null);

  // Sort rows in descending order to avoid shifting issues when deleting
  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  // Handle modified transactions
  plaid.modified.forEach(transaction => {
    const existing = existingTransactions[transaction.transaction_id];
    if (existing) {
      const rowData = transactionToRow(transaction);
      sheet.getRange(existing.rowNumber, 1, 1, headers.length).setValues([rowData]);
    }
  });

  // Handle added transactions
  if (plaid.added.length > 0) {
    const rowsToAdd = plaid.added.map(transaction => transactionToRow(transaction));
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Transactions updated successfully!');
}

/**
 * Gets all the transactions from the "Transactions" sheet.
 *
 * @param {Sheet} sheet the sheet to get the transactions from.
 * @return {Object} a map of transaction IDs to row data.
 */
function getTransactionsFromSheet(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Transactions");
  }

  const headers = [
    "Transaction ID", "Account ID", "Account Owner", "Amount", "Authorized Date", "Authorized Datetime",
    "Category", "Category ID", "Check Number", "Counterparties", "Date", "Datetime",
    "ISO Currency Code", "Location", "Logo URL", "Merchant Entity ID", "Merchant Name", "Name",
    "Payment Channel", "Payment Meta", "Pending", "Pending Transaction ID", "Personal Finance Category",
    "Personal Finance Category Icon URL", "Transaction Code", "Transaction Type", "Unofficial Currency Code", "Website"
  ];

  // Set headers if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    // Freeze the header row
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  const transactions = {};
  
  // Start from 1 to skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const transactionId = row[0]; // Transaction ID is in the first column
    if (transactionId) {
      transactions[transactionId] = {
        rowNumber: i + 1,
        data: row
      };
    }
  }

  return transactions;
}

/**
 * Converts a Plaid transaction object to a row for the Google Sheet.
 *
 * @param {Object} transaction the Plaid transaction object.
 * @return {Array} an array of values for the row.
 */
function transactionToRow(transaction) {
  return [
    transaction.transaction_id,
    transaction.account_id,
    transaction.account_owner,
    transaction.amount,
    transaction.authorized_date,
    transaction.authorized_datetime,
    transaction.category ? JSON.stringify(transaction.category) : null,
    transaction.category_id,
    transaction.check_number,
    transaction.counterparties ? JSON.stringify(transaction.counterparties) : null,
    transaction.date,
    transaction.datetime,
    transaction.iso_currency_code,
    transaction.location ? JSON.stringify(transaction.location) : null,
    transaction.logo_url,
    transaction.merchant_entity_id,
    transaction.merchant_name,
    transaction.name,
    transaction.payment_channel,
    transaction.payment_meta ? JSON.stringify(transaction.payment_meta) : null,
    transaction.pending,
    transaction.pending_transaction_id,
    transaction.personal_finance_category ? JSON.stringify(transaction.personal_finance_category) : null,
    transaction.personal_finance_category_icon_url,
    transaction.transaction_code,
    transaction.transaction_type,
    transaction.unofficial_currency_code,
    transaction.website,
  ];
}


/**
 * Creates a link token to be used to initialize Plaid Link.
 */
function createLinkToken() {

  try {
    // Prepare the request body
    const body = {
      "client_id": getSecrets().CLIENT_ID,
      "secret": getSecrets().SECRET,
      "client_name": "Bank Account to Sheets",
      "language": "en",
      "country_codes": ["US"],
      "user": {
        "client_user_id": "1",
      },
      "products": ["transactions"],
      "hosted_link": {}
    };

    // Condense the above into a single object
    const params = {
      "contentType": "application/json",
      "method": "post",
      "payload": JSON.stringify(body),
      "muteHttpExceptions": true
    };

    // Make the POST request
    const responseText = makeRequest(`${getSecrets().URL}/link/token/create`, params);
    const result = JSON.parse(responseText);

    Logger.log('Full response from /link/token/create:');
    Logger.log(JSON.stringify(result, null, 2));

    // Get or create the 'Plaid' sheet
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plaid");
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Plaid");
    }

    // Store the result in the sheet
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1).setValue(JSON.stringify(result, null, 2));

    // Tell the user that it was successful
    if (result.link_token) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Successfully created and stored link token in 'Plaid' sheet.`);
    } else {
        SpreadsheetApp.getActiveSpreadsheet().toast('Call to Plaid was successful, but no link token was returned.');
        Logger.log('Call to Plaid was successful, but no link token was returned. Full response logged above.');
    }
  } catch (e) {
    Logger.log('An error occurred in createLinkToken:');
    Logger.log(e);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Failed to create link token. Check logs for details. Error: ${e.message}`);
  }
}


/**
 * Adds the Scripts menu to the menu bar at the top.
 */
function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Scripts");
  menu.addItem("Create Link Token", "createLinkToken");
  menu.addItem("Get Link Token Info", "getLinkTokenInfo");  
  menu.addSeparator();  
  menu.addItem("Update Transactions", "updateTransactions");
  menu.addToUi();
}
