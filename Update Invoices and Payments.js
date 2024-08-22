function updateInvoicesandPayments() {
  updateStripePayments();
  updateStripeInvoices();
}

function updateStripeInvoices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Invoices');
  var paymentsSheet = ss.getSheetByName('Payments');
  var paymentsData = paymentsSheet.getDataRange().getValues();
  var paymentInfo = {}; // Object to hold payment info with invoice.id as key
  
  // Start from row 2 to skip headers
  for (var p = 1; p < paymentsData.length; p++) {
    var paymentRow = paymentsData[p];
    var invoiceId = paymentRow[0]; // Assuming invoice ID is in column A
    var paidAt = paymentRow[8]; // Assuming paid_at is in column I
    paymentInfo[invoiceId] = paidAt;
  }
  
  var customerDetails = getCustomerDetails(); // Fetch customer details
  
  var API_KEY = PropertiesService.getScriptProperties().getProperty('STRIPE_API_KEY');
  if (!API_KEY) {
    Logger.log("Error: Stripe API Key not found.");
    return;
  }

  var apiOptions = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + API_KEY
    },
    'muteHttpExceptions': true
  };

  var today = new Date();
  var startOfToday = new Date(today);
  startOfToday.setDate(today.getDate() - 7); // Set the date to 7 days ago
  var startTimestamp = Math.floor(startOfToday.getTime() / 1000); // Convert to Unix timestamp
  
  var apiUrl = `https://api.stripe.com/v1/invoices?created[gte]=${startTimestamp}`;

  var allRows = [];
  var existingInvoiceData = {}; 

  // Get existing data from the sheet
  var existingData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  existingData.forEach(function(row) {
    existingInvoiceData[row[0]] = row[8]; // Assuming invoice ID is in column A and status in column I
  });

  do {
    var response = UrlFetchApp.fetch(apiUrl, apiOptions);
    var httpResponseCode = response.getResponseCode();  // Get HTTP response code
    
    if (httpResponseCode !== 200) {
      Logger.log("Failed API request. HTTP Response Code: " + httpResponseCode);
      return;
    }
    
    var json = JSON.parse(response.getContentText());
    var invoices = json.data;
    
    if (invoices.length === 0) {
      Logger.log("No new invoices found from API.");
      return;
    }

    for (var i = 0; i < invoices.length; i++) {
      var invoice = invoices[i];
      if (invoice && invoice.id) {
        var row = processInvoice(invoice, customerDetails, paymentInfo);

        if (!existingInvoiceData[invoice.id]) {
          allRows.push(row);
          Logger.log("Added new invoice: " + invoice.id);
        } else if (existingInvoiceData[invoice.id] !== invoice.status) {
          allRows.push(row);
          Logger.log("Updated invoice: " + invoice.id + ". Status changed from " + existingInvoiceData[invoice.id] + " to " + invoice.status);
          
          var rowIndex = existingData.findIndex(function(existingRow) {
            return existingRow[0] === invoice.id;
          }) + 1;
          sheet.deleteRow(rowIndex);
        }
      }
    }

    apiUrl = json.has_more ? `https://api.stripe.com/v1/invoices?limit=100&starting_after=${invoices[invoices.length - 1].id}` : null;

  } while (apiUrl);

  if (allRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
  }
  deleteDuplicateInvoices();
}

function processInvoice(invoice, customerDetails, paymentInfo) {
  var customerID = invoice.customer;
  var customerInfo = customerDetails[customerID] || {};
  var customerName = customerInfo.name || 'N/A';
  var customerEmail = customerInfo.email || 'N/A';

  var createdDate = invoice.created ? addHoursToDate(new Date(invoice.created * 1000), 4).toUTCString() : '';
  var dueDate = invoice.due_date ? addHoursToDate(new Date(invoice.due_date * 1000), 4).toUTCString() : '';
  var paidAt = paymentInfo[invoice.id] ? addHoursToDate(new Date(paymentInfo[invoice.id]), 4).toUTCString() : '';

  var row = [
    invoice.id || '',
    customerID || '',
    customerName,
    customerEmail,
    invoice.amount_due / 100 || '',
    invoice.amount_paid / 100 || '',
    invoice.amount_remaining / 100 || '',
    invoice.currency || '',
    invoice.status || '',
    createdDate,
    dueDate,
    paidAt,
    invoice.subscription || '',
    invoice.collection_method || '',
    invoice.hosted_invoice_url || ''
  ];

  return row;
}

function getCustomerDetails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Customers');

  var data = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues(); // assuming headers are in the first row
  var customerDetails = {};

  data.forEach(function(row) {
    var customerId = row[0];
    var customerName = row[3];
    var customerEmail = row[2];

    customerDetails[customerId] = {
      name: customerName,
      email: customerEmail
    };
  });

  return customerDetails;
}

function deleteDuplicateInvoices() {
  var sheetName = 'Invoices'; // Name of the sheet where duplicates need to be removed
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues(); // Get all data from the sheet
  var invoiceIds = {};
  var rowsToDelete = [];

  // Loop through all rows in reverse order to keep the row with the greatest number
  for (var i = data.length - 1; i >= 0; i--) {
    var invoiceId = data[i][0]; // Assuming invoice ID is in the first column
    if (invoiceIds[invoiceId]) {
      // If we have seen this ID, mark the current row for deletion
      rowsToDelete.push(i + 1); // +1 because rows are 1-indexed in Apps Script
    } else {
      // If we haven't seen this ID, add it to the object
      invoiceIds[invoiceId] = true;
    }
  }

  // Delete marked rows starting from the bottom of the sheet
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }
}

function addHoursToDate(date, hours) {
  return new Date(date.getTime() + hours * 60 * 60 * 1000);
}

function updateStripePayments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var paymentsSheet = ss.getSheetByName('Payments');
  var customersSheet = ss.getSheetByName('Customers');

  var apiKey = PropertiesService.getScriptProperties().getProperty('STRIPE_API_KEY');
  var apiOptions = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'muteHttpExceptions': true
  };

  // Map customer details from the Customers sheet
  var customerDataRange = customersSheet.getRange(1, 1, customersSheet.getLastRow(), 4).getValues();
  var customerMap = {}; 
  for (var i = 1; i < customerDataRange.length; i++) { // Start from 1 to skip header
    var customerId = customerDataRange[i][0];
    var customerEmail = customerDataRange[i][2];
    var customerName = customerDataRange[i][3];
    customerMap[customerId] = {
      email: customerEmail,
      name: customerName
    };
  }

  var existingPaymentIds = paymentsSheet.getRange(1, 2, paymentsSheet.getLastRow()).getValues().flat();

  // Fetch all payments created in the last 10 days
  var today = new Date();
  var startOfToday = new Date(today.setDate(today.getDate() - 10));
  var startTimestamp = Math.floor(startOfToday.getTime() / 1000);
  var apiUrl = `https://api.stripe.com/v1/charges?created[gte]=${startTimestamp}`;

  var allRows = [];
  
  do {
    var response = UrlFetchApp.fetch(apiUrl, apiOptions);
    var json = JSON.parse(response.getContentText());
    var charges = json.data;

    for (var i = 0; i < charges.length; i++) {
      var charge = charges[i];
      if (charge && charge.id && !existingPaymentIds.includes(charge.id)) {
        var row = processPayment(charge, customerMap);
        allRows.push(row);
      }
    }

    apiUrl = json.has_more ? `https://api.stripe.com/v1/charges?created[gte]=${startTimestamp}&starting_after=${charges[charges.length - 1].id}` : null;

  } while (apiUrl);

  if (allRows.length > 0) {
    paymentsSheet.getRange(paymentsSheet.getLastRow() + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
  }
}

function processPayment(charge, customerMap) {
  var customerID = charge.customer;
  var customerName = customerMap[customerID] ? customerMap[customerID].name : 'N/A';
  var customerEmail = customerMap[customerID] ? customerMap[customerID].email : 'N/A';
  var createdDateCharge = charge.created ? addHoursToDate(new Date(charge.created * 1000), 4).toUTCString() : '';
  var paidAtDateCharge = charge.paid_at ? addHoursToDate(new Date(charge.paid_at * 1000), 4).toUTCString() : '';

  var row = [
    charge.invoice || '',
    charge.id || '',
    customerID || '',
    customerName,
    customerEmail,
    charge.amount / 100 || '',
    charge.currency || '',
    charge.status || '',
    createdDateCharge,
    paidAtDateCharge,
    charge.description || '',
    charge.outcome ? charge.outcome.seller_message : ''
  ];

  return row;
}

function addHoursToDate(date, hours) {
  return new Date(date.getTime() + hours * 60 * 60 * 1000);
}
