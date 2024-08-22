function loadStripePayments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var paymentsSheet = ss.getSheetByName('Payments');
  var customersSheet = ss.getSheetByName('Customers');

  paymentsSheet.clear();

  var API_KEY = PropertiesService.getScriptProperties().getProperty('STRIPE_API_KEY');
  var apiOptions = {
    'method' : 'get',
    'headers' : {'Authorization' : 'Bearer ' + API_KEY},
    'muteHttpExceptions': true
  };

  // Fetch all customer data from the Customers sheet
  var customerData = customersSheet.getRange(2, 1, customersSheet.getLastRow() - 1, 4).getValues();
  var customerLookup = {};
  customerData.forEach(function(row) {
    customerLookup[row[0]] = {name: row[3], email: row[2]};
  });

  var allRows = [];
  var apiUrl = 'https://api.stripe.com/v1/charges?limit=100';
  var hasMore = true;

  while (hasMore) {
    var response = UrlFetchApp.fetch(apiUrl, apiOptions);
    var json = JSON.parse(response.getContentText());
    var charges = json.data;

    if (charges.length === 0) {
      hasMore = false;
      break;
    }

    charges.forEach(function(charge) {
      var customerID = charge.customer;
      var customerName = customerLookup[customerID] ? customerLookup[customerID].name : 'N/A';
      var customerEmail = customerLookup[customerID] ? customerLookup[customerID].email : 'N/A';
      
      var row = [
        charge.invoice || '',
        charge.id || '',
        customerID || '',
        customerName,
        customerEmail,
        charge.amount / 100 || '',
        charge.currency || '',
        charge.status || '',
        charge.created ? new Date(charge.created * 1000).toUTCString() : '',
        charge.paid_at ? new Date(charge.paid_at * 1000).toUTCString() : '',
        charge.description || '',
        charge.outcome ? charge.outcome.seller_message : "", 
      ];

      allRows.push(row);
    });

    apiUrl = 'https://api.stripe.com/v1/charges?limit=100&starting_after=' + charges[charges.length - 1].id;
  }

  // Append rows to the spreadsheet
  if (allRows.length > 0) {
    paymentsSheet.getRange(paymentsSheet.getLastRow() + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
  }
}
