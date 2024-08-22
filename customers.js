function loadStripeCustomers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Customers'); // The sheet where customer data will be stored

  var apiKey = PropertiesService.getScriptProperties().getProperty('STRIPE_API_KEY'); // Retrieve the Stripe API key
  var apiOptions = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'muteHttpExceptions': true
  };

  var lastCustomerId = null;
  var allRows = []; // Array to hold all customer data rows
  
  var paymentData = fetchPaymentsData(); // Get payment data for customers

  do {
    var apiUrl = 'https://api.stripe.com/v1/customers?limit=100';
    if (lastCustomerId) {
      apiUrl += '&starting_after=' + lastCustomerId;
    }

    var response = UrlFetchApp.fetch(apiUrl, apiOptions);
    var jsonResponse = JSON.parse(response.getContentText());
    var customers = jsonResponse.data;

    if (customers.length === 0) {
      break; // Stop fetching if no more customers are found
    }

    for (var i = 0; i < customers.length; i++) {
      var customer = customers[i];
      lastCustomerId = customer.id;
      Logger.log(customer); // Log each customer for debugging
      
      var totalSpend = 0;
      var paymentCount = 0;

      if (paymentData[customer.id]) {
        totalSpend = paymentData[customer.id].totalSpend;
        paymentCount = paymentData[customer.id].paymentCount;
      }

      var row = [
        customer.id || '',
        customer.description || '',
        customer.email || '',
        customer.name || '',
        customer.created ? new Date(customer.created * 1000).toUTCString() : '',
        totalSpend,
        paymentCount,
        customer.default_source || '',
        customer.invoice_settings.default_payment_method || '',
        customer.invoice_prefix || ''
      ];
      allRows.push(row); // Add the customer data to the allRows array
    }
  } while (customers.length === 100); // Continue fetching until fewer than 100 customers are returned

  sheet.clear();
  var headers = ["ID", "Description", "Email", "Name", "Created (UTC)", "Total Spend", "Payment Count", "Card ID", "Default Payment Method", "Invoice Prefix"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
}

function fetchPaymentsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var paymentSheet = ss.getSheetByName('Payments'); // The sheet containing payment data
  var data = paymentSheet.getDataRange().getValues();

  var paymentData = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var customerId = row[2];
    var amount = parseFloat(row[5]);

    if (!paymentData[customerId]) {
      paymentData[customerId] = {
        totalSpend: 0,
        paymentCount: 0
      };
    }

    paymentData[customerId].totalSpend += amount;
    paymentData[customerId].paymentCount++;
  }

  return paymentData; // Return the aggregated payment data
}
