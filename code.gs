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
        if (session.on_success && session.on_success.public_token) {
          Logger.log("Public token found in /link/token/get response. Exchanging automatically.");
          _exchangePublicTokenInternal(session.on_success.public_token);
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
    Logger.log("No access token found in 'Plaid' sheet. Falling back to getSecrets().");
    accessToken = getSecrets().ACCESS_TOKEN;
  }

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
 * Fetch the transactions that are currently on the sheet.
 *
 * @param {SpreadsheetApp.Sheet} sheet the sheet to fetch the transactions from.
 * @return {Object} the transactions.
 */
function getTransactionsFromSheet(sheet) {

  const result = {};
  result.transactions = [];
  result.available = 0.0;
  result.current = 0.0;

  // Get the headers
  result.headers = sheet.getRange(getHeaderRowNumber(sheet), 1, 1, sheet.getLastColumn()).getValues().flat();
  result.headers = result.headers.map(item => item.replace("?", ""));
  result.headers = result.headers.map(item => item.toLowerCase());

  // Don't bother if it's empty
  if (sheet.getLastRow() === getHeaderRowNumber(sheet)) {
    Logger.log(`We fetched ${result.transactions.length} transactions from the sheet named ${sheet.getName()}.`);
    return result;
  }

  // Get the transactions, starting with most recent
  const values = sheet.getRange(getHeaderRowNumber(sheet) + 1, 1, sheet.getLastRow() - getHeaderRowNumber(sheet), sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const newSheetTxn = {};
    for (let j = 0; j < result.headers.length; j++) {
      newSheetTxn[result.headers[j].toLowerCase()] = values[i][j];
    }
    if (typeof newSheetTxn.date === "number") {
      newSheetTxn.date = new Date(newSheetTxn.date)
    }
    result.transactions.push(newSheetTxn);

    // Increment the balance(s)
    result.current += Number(values[i][6]);
    if (values[i][7] === false) {
      result.available += Number(values[i][6]);
    }

  }

  Logger.log(`We fetched ${result.transactions.length} transactions from the sheet named ${sheet.getName()}.`);

  return result;

}


/**
 * Convert a Plaid transaction to a transaction for the sheet.
 *
 * @param {Object} plaidTxn the transaction to convert.
 * @param {Object} sheetTxn the existing sheet transaction to update.
 * @return {Object} the converted transaction.
 */
function plaidToSheet(plaidTxn, sheetTxn = undefined) {

  // Use existing values if we have them
  let internal;
  let notes;
  let category;
  let subcategory;
  let channel;
  if (sheetTxn === undefined) {
    internal = false;
    notes = "";
    if (plaidTxn.category === null) {
      category = "UNKNOWN";
      subcategory = "UNKNOWN";
    } else {
      category = plaidTxn.category[0];
      subcategory = "";
      for (const subcat of plaidTxn.category.slice(1)) subcategory = subcategory + subcat + " ";
      subcategory = subcategory.slice(0, -1);
    }
    channel = plaidTxn.payment_channel;

  } else {
    internal = sheetTxn.internal;
    notes = sheetTxn.notes;
    category = sheetTxn.category;
    subcategory = sheetTxn.subcategory;
    channel = sheetTxn.channel;
  }

  // Return the transaction for the sheet
  return {
    "id": plaidTxn.transaction_id,
    "date": plaidTxn.date,
    "name": plaidTxn.name,
    "category": category,
    "subcategory": subcategory,
    "channel": channel,
    "account": plaidTxn.account_name,
    "amount": -plaidTxn.amount,
    "pending": plaidTxn.pending,
    "internal": internal,
    "notes": notes
  };

}


/**
 * Searches the transactions from the sheet to see if a given Plaid transaction already exists.
 * Painfully inefficient.
 *
 * @param {Object[]} sheetTxns the sheet transactions to search.
 * @param {Object} plaidTxn the Plaid transaction to search for.
 * @return {Number} the index of the plaidTxn, or -1 if it doesn't exist.
 */
function getIndexOfPlaidFromSheet(sheetTxns, plaidTxn) {

  const sameDateAndAmount = [];

  for (let i = 0; i < sheetTxns.length; i++) {

    // Check the IDs
    if (sheetTxns[i].id === plaidTxn.pending_transaction_id) {
      return i;
    } else if (sheetTxns[i].id === plaidTxn.transaction_id) {
      return i;
    }


    /* Only enable when the ACCESS_TOKEN has been changed
    // Check the date, name, and amount
    let date = sheetTxns[i].date
    if (typeof date === "number") {
      date = new Date(date)
    }
    if (date.getTime() === plaidTxn.date &&
      sheetTxns[i].name === plaidTxn.name &&
      sheetTxns[i].amount === -plaidTxn.amount) {
      return i;
    }

    // For if the name has changed
    if (date.getTime() === plaidTxn.date &&
      sheetTxns[i].amount === -plaidTxn.amount) {
      sameDateAndAmount.push(i)
    }
    */
  }

  // If there was only one with that date and amount
  if (sameDateAndAmount.length === 1) {
    return sameDateAndAmount[0];
  }

  return -1;
}


/**
 * Searches the transactions from plaid for the transaction with the ID, and returns its index.
 * Painfully inefficient.
 *
 * @param {Object[]} plaidTxns the Plaid transactions to search.
 * @param {string} id ID to search for.
 * @return {Number} the index of the transaction, or -1 if it doesn't exist.
 */
function getIndexOfIdFromPlaid(plaidTxns, id) {

  for (let i = 0; i < plaidTxns.length; i++) {
    if (plaidTxns[i].transaction_id === id) {
      return i;
    } else if (plaidTxns[i].pending_transaction_id === id) {
      return i;
    }
  }
  return -1;
}


/**
 * Inserts the sheet transaction into the sheet transactions in the correct place.
 *
 * @param {Object[]} sheetTxns the list of transactions from the sheet.
 * @param {Object} sheetTxn the sheet transaction to insert.
 * @return {Object[]} the updated sheet transactions.
 */
function saveNewSheetTransaction(sheetTxns, sheetTxn) {

  // Insert it when we first encounter an existing one with a smaller date
  for (let i = 0; i < sheetTxns.length; i++) {
    if (sheetTxn.date >= sheetTxns[i].date) {
      sheetTxns.splice(i, 0, sheetTxn);
      return sheetTxns;
    }
  }

  // If the new transaction is the oldest then add it at the end
  sheetTxns.push(sheetTxn);
  return sheetTxns;

}


/**
 * Writes the sheet transactions to the sheet.
 *
 * @param {SpreadsheetApp.Sheet} sheet the sheet to write the transactions to.
 * @param {Object[]} sheetTxns the sheet transactions to write.
 * @param {string[]} headers the headers of the sheet.
 */
function writeTransactionsToSheet(sheet, sheetTxns, headers) {

  const result = [];
  for (let i = 0; i < sheetTxns.length; i++) {

    const row = headers.slice();
    for (const [key, value] of Object.entries(sheetTxns[i])) {
      if (key === "date") {
        let date = new Date();
        date.setTime(value);
        row[row.indexOf(key)] = date;
      } else {
        row[row.indexOf(key)] = value;
      }
    }
    result.push(row);

  }

  sheet.deleteRows(getHeaderRowNumber(sheet) + 2, sheet.getLastRow() - (getHeaderRowNumber(sheet) + 1));
  sheet.insertRowsAfter(getHeaderRowNumber(sheet) + 1, result.length - 1);
  sheet.getRange(getHeaderRowNumber(sheet) + 1, 1, result.length, sheet.getLastColumn()).setValues(result);

}


/**
 * Formats the date as a nice string.
 *
 * @param {Date} date the date to parse.
 * @return {string} the nicely formatted date.
 */
function formatDate(date) {

  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return `${days[date.getDay()]} ${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;

}


/**
 * Updates the transactions in the Transactions sheet.
 */
function updateTransactions() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");

  let existing = getTransactionsFromSheet(sheet);
  const plaid = syncTransactionsFromPlaid();

  // Handle removed transactions
  const removed_ids = plaid.removed.map(t => t.transaction_id);
  let kept_transactions = existing.transactions.filter(t => !removed_ids.includes(t.id));

  // Handle modified transactions
  const modified_map = new Map(plaid.modified.map(t => [t.transaction_id, t]));
  for (let i = 0; i < kept_transactions.length; i++) {
    const existing_txn = kept_transactions[i];
    if (modified_map.has(existing_txn.id)) {
      const plaid_txn = modified_map.get(existing_txn.id);
      // Add account_name which is needed by plaidToSheet
      plaid_txn.account_name = existing_txn.account;
      plaid_txn.date = Date.parse(plaid_txn.date); // Ensure date is in correct format
      kept_transactions[i] = plaidToSheet(plaid_txn, existing_txn);
    }
  }

  // Handle added transactions
  for (const plaidTxn of plaid.added) {
    let account_name = "?unknown?";
    for (let j = 0; j < plaid.accounts.length; j++) {
      if (plaid.accounts[j].account_id === plaidTxn.account_id) {
        account_name = plaid.accounts[j].name;
        break;
      }
    }
    plaidTxn.account_name = account_name;
    plaidTxn.date = Date.parse(plaidTxn.date); // Ensure date is in correct format
    const newSheetTxn = plaidToSheet(plaidTxn);
    kept_transactions = saveNewSheetTransaction(kept_transactions, newSheetTxn);
  }

  const num_added = plaid.added.length;
  const num_modified = plaid.modified.length;
  const num_removed = plaid.removed.length;

  if (num_added === 0 && num_modified === 0 && num_removed === 0) {
    Logger.log("No new transaction changes.");
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast("No new changes to the transactions were found.");
    } catch (error) {}
  } else {
    // Write the transactions to the sheet
    Logger.log(`There are ${kept_transactions.length} transactions to write.`);
    writeTransactionsToSheet(sheet, kept_transactions, existing.headers);
    Logger.log(`Finished writing transactions to the sheet named ${sheet.getName()}.`);

    // Format the sheet neatly
    formatNeatlyTransactions(plaid);
    Logger.log(`Finished formatting the sheet named ${sheet.getName()} neatly.`);

    // Produce a message to tell the user of the changes
    try {
      const ui = SpreadsheetApp.getUi();
      let message = `Added: ${num_added}, Modified: ${num_modified}, Removed: ${num_removed}`;
      ui.alert("Transaction Sync Complete", message, ui.ButtonSet.OK);
    } catch (error) {}
  }

  // Update when this script was last run
  const range = sheet.getRange(getHeaderRowNumber(sheet) - 1, sheet.getLastColumn());
  if (range !== undefined) {
    const date = new Date();
    let minutes = date.getMinutes().toString();
    if (parseInt(minutes) < 10) minutes = "0" + minutes;
    const dateString = `Last updated on ${formatDate(date)} at ${date.getHours()}:${minutes}.`;
    range.setValue(dateString);
  }

}


/**
 * Extract and return the totals for the given account.
 * 
 * @param {Object} account the account from Plaid.
 * @return {Object} the totals.
 */
function getPlaidAccountTotals(account) {

  const result = {};

  // For a credit card account
  if (account.type === "credit") {
    result.available = -(account.balances.limit - account.balances.available);
    result.current = -account.balances.current;
    result.pending = result.available - result.current;

    // For a depository (normal current) account, or anything else
  } else {
    if (account.balances.available === null) {
      result.available = account.balances.current;
      result.current = account.balances.current;
      result.pending = 0;
    } else {
      result.available = account.balances.available;
      result.current = account.balances.current;
      result.pending = result.available - result.current;
    }
  }

  return result;

}


/**
 * Formats the 'Transactions' sheet neatly.
 * 
 * @param {Object} plaidResult the result of transactions.get from Plaid.
 */
function formatNeatlyTransactions(plaidResult = undefined) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");

  // Get the headers
  let headers = sheet.getRange(getHeaderRowNumber(sheet), 1, 1, sheet.getLastColumn()).getValues().flat();
  headers = headers.map(item => item.replace("?", ""));
  headers = headers.map(item => item.toLowerCase());

  // Get column letters (for A1 notation)
  const amountColNum = headers.indexOf("amount") + 1;

  // Create named ranges
  for (let i = 0; i < headers.length; i++) {
    const range = sheet.getRange(getHeaderRowNumber(sheet) + 1, i + 1, sheet.getLastRow() - getHeaderRowNumber(sheet), 1);
    SpreadsheetApp.getActiveSpreadsheet().setNamedRange(`${headers[i]}s`, range);
  }

  if (plaidResult !== undefined) {
    sheet.deleteRows(1, getHeaderRowNumber(sheet) - 2);
    Logger.log(`There are ${plaidResult.accounts.length} account(s).`)
    if (plaidResult.accounts.length === 1) {
      sheet.insertRows(1, 2)

      // Add the total titles and merge them
      sheet.getRange(1, 2, 1, amountColNum - 2).setValue("CURRENT BALANCE");
      sheet.getRange(2, 2, 1, amountColNum - 2).setValue("AMOUNT PENDING (UNACCOUNTED FOR)");
      sheet.getRange(3, 2, 1, amountColNum - 2).setValue("AMOUNT PENDING (ACCOUNTED FOR)");
      sheet.getRange(4, 2, 1, amountColNum - 2).setValue("AVAILABLE BALANCE");
      sheet.getRange(1, 2, 4, amountColNum - 2).mergeAcross();

      // Extract the totals
      const totals = getPlaidAccountTotals(plaidResult.accounts[0]);

      // Add the totals themselves
      sheet.getRange(1, amountColNum).setValue(`${totals.current}`);
      sheet.getRange(2, amountColNum).setValue(`=${totals.pending}-SUMIF(pendings, "=TRUE", amounts)`);
      sheet.getRange(3, amountColNum).setValue(`=SUMIF(pendings, "=TRUE", amounts)`);
      sheet.getRange(4, amountColNum).setValue(`=${totals.current}-${totals.pending}`);

    } else {
      sheet.insertRows(1, (plaidResult.accounts.length * 3) + 4);

      // Prepare to track the grand totals
      const grandTotals = {};
      grandTotals.available = 0;
      grandTotals.current = 0;
      grandTotals.pending = 0;

      // For each account
      for (let i = 1; i <= plaidResult.accounts.length; i++) {

        // Add the total titles and merge them
        sheet.getRange((i * 3) - 2, 2, 1, amountColNum - 2).setValue(`${plaidResult.accounts[i - 1].name} CURRENT BALANCE`);
        sheet.getRange((i * 3) - 1, 2, 1, amountColNum - 2).setValue(`${plaidResult.accounts[i - 1].name} AMOUNT PENDING`);
        sheet.getRange(i * 3, 2, 1, amountColNum - 2).setValue(`${plaidResult.accounts[i - 1].name} AVAILABLE BALANCE`);
        sheet.getRange((i * 3) - 2, 2, 3, amountColNum - 2).mergeAcross();

        // Extract the totals, and accumulate the grand totals
        const totals = getPlaidAccountTotals(plaidResult.accounts[i - 1]);
        grandTotals.available = totals.available + grandTotals.available;
        grandTotals.current = totals.current + grandTotals.current;
        grandTotals.pending = totals.pending + grandTotals.pending;

        // Add the totals themselves
        sheet.getRange((i * 3) - 2, amountColNum).setValue(`=ROUND(${totals.current}, 2)`);
        sheet.getRange((i * 3) - 1, amountColNum).setValue(`=ROUND(${totals.pending}, 2)`);
        sheet.getRange(i * 3, amountColNum).setValue(`=ROUND(${totals.available}, 2)`);

      }

      // Hide the account breakdown, because it takes up too much space
      const startingRow = (plaidResult.accounts.length * 3) + 2;
      sheet.hideRows(1, startingRow - 1);

      // Add the total titles and merge them
      sheet.getRange(startingRow, 2, 1, amountColNum - 2).setValue("TOTAL CURRENT BALANCE");
      sheet.getRange(startingRow + 1, 2, 1, amountColNum - 2).setValue("TOTAL AMOUNT PENDING (UNACCOUNTED FOR)");
      sheet.getRange(startingRow + 2, 2, 1, amountColNum - 2).setValue("TOTAL AMOUNT PENDING (ACCOUNTED FOR)");
      sheet.getRange(startingRow + 3, 2, 1, amountColNum - 2).setValue("TOTAL AVAILABLE BALANCE");
      sheet.getRange(startingRow, 2, 4, amountColNum - 2).mergeAcross();

      // Add the totals themselves
      sheet.getRange(startingRow, amountColNum).setValue(`=ROUND(${grandTotals.current}, 2)`);
      sheet.getRange(startingRow + 1, amountColNum).setValue(`=ROUND(${grandTotals.pending}, 2)-SUMIF(pendings, "=TRUE", amounts)`);
      sheet.getRange(startingRow + 2, amountColNum).setValue(`=SUMIF(pendings, "=TRUE", amounts)`);
      sheet.getRange(startingRow + 3, amountColNum).setValue(`=ROUND(${grandTotals.available}, 2)`);

    }
  }

  // Convert the TRUE/FALSE columns to checkboxes
  sheet.getRange(`pendings`).insertCheckboxes();
  sheet.getRange(`internals`).insertCheckboxes();

  // Add conditional formatting to the amount column
  const amountRange = sheet.getRange(`amounts`);
  const positiveRule = SpreadsheetApp.newConditionalFormatRule().setFontColor("#1B5E20").whenNumberGreaterThan(0).setRanges([amountRange]).build();
  const negativeRule = SpreadsheetApp.newConditionalFormatRule().setFontColor("#B71C1C").whenNumberLessThan(0).setRanges([amountRange]).build();
  sheet.setConditionalFormatRules([positiveRule, negativeRule]);

  // Add data validation for the categories, subcategories, and channels
  let range = sheet.getRange("categorys");
  let values = sheet.getRange("Categories")
  let rule = SpreadsheetApp.newDataValidation().requireValueInRange(values, true).setAllowInvalid(false).build();
  range.setDataValidation(rule);

  range = sheet.getRange("subcategorys");
  values = sheet.getRange("Subcategories")
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(values, true).setAllowInvalid(false).build();
  range.setDataValidation(rule);

  range = sheet.getRange("channels");
  values = sheet.getRange("ChannelsValues")
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(values, true).setAllowInvalid(false).build();
  range.setDataValidation(rule);

  // Freeze the top rows and hide two columns
  sheet.setFrozenRows(getHeaderRowNumber(sheet));
  sheet.hideColumn(sheet.getRange("ids"));
  sheet.hideColumn(sheet.getRange("accounts"));

  // Add protection for ranges that shouldn't be edited
  for (const protection of sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)) protection.remove();
  for (const name of ["ids", "dates", "names", "accounts", "amounts", "pendings"]) {
    sheet.getRange(name).protect().setWarningOnly(true);
  }

  // Recreate the filter
  amountRange.getFilter().remove();
  sheet.getRange(getHeaderRowNumber(sheet), 1, sheet.getLastRow() - (getHeaderRowNumber(sheet) - 1), sheet.getLastColumn()).createFilter();
}


/**
 * Formats the 'Weekly Summary' sheet neatly.
 */
function formatNeatlyWeeklySummary() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly Summary");

  // Hide rows in the future
  sheet.showRows(1, sheet.getLastRow());
  const now = new Date();
  for (let i = 3; i < sheet.getLastRow() - 1; i++) {
    if (sheet.getRange(i, 2).getValue().getTime() <= now.getTime()) {
      sheet.hideRows(3, i - 3)
      break;
    }
  }

}


/** 
 * Searches for and returns the row number of the header row.
 * 
 * @param {SpreadsheetApp.Sheet} sheet the sheet to search.
 * @return {number} the row number, or -1 if it can't be found.
*/
function getHeaderRowNumber(sheet) {

  const range = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  for (let i = 0; i < range.length; i++) {
    if (range[i][0] === "ID") {
      return i + 1;
    }
  }

  return -1;

}


/**
 * Runs all the formatNeatly functions.
 */
function formatAll() {
  formatNeatlyTransactions()
  formatNeatlyWeeklySummary()
}


/**
 * Updates transactions and then formats everything neatly.
 */
function doEverything() {
  updateTransactions()
  formatNeatlyWeeklySummary()
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
  menu.addItem("Update Transactions", "updateTransactions");
  menu.addItem("Format the Transactions sheet neatly", "formatNeatlyTransactions");
  menu.addItem("Format the Weekly Summary sheet neatly", "formatNeatlyWeeklySummary");
  menu.addSeparator();
  menu.addItem("Format all sheets neatly", "formatAll");
  menu.addSeparator();
  menu.addItem("Do everything", "doEverything");
  menu.addSeparator();
  menu.addItem("Create Link Token", "createLinkToken");
  menu.addItem("Get Link Token Info", "getLinkTokenInfo");
  menu.addToUi();
}
