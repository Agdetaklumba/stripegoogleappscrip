function loadStripeInvoices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invoicesSheet = ss.getSheetByName('Invoices');
  var paymentsSheet = ss.getSheetByName('Payments');
  var customersSheet = ss.getSheetByName('Customers');
  
  invoicesSheet.clear();
  var headers = ["Invoice ID", "Customer ID", "Customer Name", "Customer email", "Amount Due", "Amount Paid", "Amount Remaining", "Currency", "Status", "Created (UTC)", "Due Date (UTC)", "Paid At (UTC)", "Subscription", "Billing", "Payment Page"];
  invoicesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var apiKey = PropertiesService.getScriptProperties().getProperty('STRIPE_API_KEY');
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'muteHttpExceptions': true
  };

  var customerData = customersSheet.getDataRange().getValues();
  var customerMap = {};
  
  for (var i = 1; i < customerData.length; i++) {
    var row = customerData[i];
    customerMap[row[0]] = {name: row[3], email: row[2]};
  }

  var paymentData = paymentsSheet.getDataRange().getValues();
  var paymentMap = {};

  for (var j = 1; j < paymentData.length; j++) {
    var paymentRow = paymentData[j];
    var invoiceId = paymentRow[0];
    var paidAt = paymentRow[8];
    paymentMap[invoiceId] = paidAt;
  }

  var lastInvoiceId = null;
  var allRows = [];

  while (true) {
    var apiUrl = 'https://api.stripe.com/v1/invoices?limit=100';
    if (lastInvoiceId !== null) {
      apiUrl += '&starting_after=' + lastInvoiceId;
    }
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseData = JSON.parse(response.getContentText());
    var invoices = responseData.data;
    
    if (invoices.length === 0) {
      break;
    }
    
    for (var k = 0; k < invoices.length; k++) {
      var invoice = invoices[k];
      if (!invoice || !invoice.id) {
        continue;
      }

      var customerId = invoice.customer;
      var customerInfo = customerMap[customerId] || {name: 'N/A', email: 'N/A'};
      var subscription = invoice.subscription || '';

      var createdDate = '';
      if (invoice.created) {
        var createdTimestamp = new Date(invoice.created * 1000);
        createdDate = addHoursToDate(createdTimestamp, 4).toUTCString();
      }

      var dueDate = '';
      if (invoice.due_date) {
        var dueTimestamp = new Date(invoice.due_date * 1000);
        dueDate = addHoursToDate(dueTimestamp, 4).toUTCString();
      }

      var paidDate = '';
      if (paymentMap[invoice.id]) {
        var paidTimestamp = new Date(paymentMap[invoice.id]);
        paidDate = addHoursToDate(paidTimestamp, 4).toUTCString();
      }

      var row = [
        invoice.id || '',
        customerId || '',
        customerInfo.name,
        customerInfo.email,
        (invoice.amount_due / 100) || '',
        (invoice.amount_paid / 100) || '',
        (invoice.amount_remaining / 100) || '',
        invoice.currency || '',
        invoice.status || '',
        createdDate,
        dueDate,
        paidDate,
        subscription,
        invoice.collection_method,
        invoice.hosted_invoice_url || ''
      ];

      lastInvoiceId = invoice.id;
      allRows.push(row);
    }
  }

  invoicesSheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
}

function addHoursToDate(date, hours) {
  return new Date(date.getTime() + hours * 60 * 60 * 1000);
}
