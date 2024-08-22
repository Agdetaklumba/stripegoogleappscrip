function loadSubs() {
  loadStripeSubscriptions();
  loadStripeScheduledSubscriptions();
}

function loadStripeSubscriptions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Subscriptions'); // Subscriptions Sheet
  var customerSheet = ss.getSheetByName('Customers'); // Customer Sheet

  // Reading Customer Data from 'Customer' Sheet
  var customerDataRange = customerSheet.getDataRange();
  var customerValues = customerDataRange.getValues();
  var customerMap = {};
  
  for (var i = 1; i < customerValues.length; i++) {
    var row = customerValues[i];
    customerMap[row[0]] = { description: row[1], email: row[2] }; // Map customer ID to description and email
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('STRIPE_API_KEY');
  var apiOptions = {
    'method': 'get',
    'headers': { 'Authorization': 'Bearer ' + apiKey },
    'muteHttpExceptions': true
  };
  
  var lastSubscriptionId = null;
  var productCache = {};
  var allRows = [];

  while (true) {
    var apiUrl = 'https://api.stripe.com/v1/subscriptions?limit=100';
    if (lastSubscriptionId) {
      apiUrl += '&starting_after=' + lastSubscriptionId;
    }
    
    var response = UrlFetchApp.fetch(apiUrl, apiOptions);
    var json = JSON.parse(response.getContentText());
    var subscriptions = json.data;
    
    if (subscriptions.length === 0) {
      break;
    }

    for (var i = 0; i < subscriptions.length; i++) {
      var subscription = subscriptions[i];
      if (!(subscription && subscription.id)) {
        continue;
      }

      lastSubscriptionId = subscription.id;
      var customerID = subscription.customer;

      // Use cached customer data from the Customer Sheet
      var customerInfo = customerMap[customerID] || { description: 'N/A', email: 'N/A' };

      var productID = subscription.items.data[0].plan.product;
      var productData = productCache[productID];
      if (!productData) {
        var productApiUrl = 'https://api.stripe.com/v1/products/' + productID;
        var productResponse = UrlFetchApp.fetch(productApiUrl, apiOptions);
        productData = JSON.parse(productResponse.getContentText());
        productCache[productID] = productData;
      }

      var totalAmount = 0;
      for (var j = 0; j < subscription.items.data.length; j++) {
        var item = subscription.items.data[j];
        if (item.plan && item.plan.amount && item.quantity) {
          totalAmount += item.plan.amount * item.quantity;
        }
      }
      totalAmount = totalAmount / 100;

      var row = [
        subscription.id || '',
        customerID || '',
        customerInfo.description,
        customerInfo.email,
        subscription.plan && subscription.plan.id || '',
        subscription.quantity || '',
        subscription.plan && subscription.plan.interval || '',
        totalAmount,
        subscription.status || '',
        formatUtcDateToUtcPlus4(subscription.created),
        formatUtcDateToUtcPlus4(subscription.start_date),
        formatUtcDateToUtcPlus4(subscription.current_period_start),
        formatUtcDateToUtcPlus4(subscription.current_period_end),
        subscription.canceled_at ? formatUtcDateToUtcPlus4(subscription.canceled_at) : 'N/A',
        subscription.cancel_at_period_end ? 'Yes' : 'No',
        subscription.ended_at ? formatUtcDateToUtcPlus4(subscription.ended_at) : 'N/A',
        customerInfo.name || 'N/A',
        (subscription.cancel_at ? formatUtcDateToUtcPlus4(subscription.cancel_at) : 'N/A'),
        productData && productData.name || 'N/A'
      ];

      allRows.push(row);
    }
  }

  sheet.clear();
  var headers = ["ID", "Customer ID", "Customer Description", "Customer Email", "Plan", "Quantity", "Interval", "Amount", "Status", "Created (UTC+4)", "Start Date (UTC+4)", "Current Period Start (UTC+4)", "Current Period End (UTC+4)", "Canceled At (UTC+4)", "Cancel At Period End", "Ended At (UTC+4)", "Customer name", "Scheduled to Cancel", "Product"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
}

function formatUtcDateToUtcPlus4(timestamp) {
  if (!timestamp) return 'N/A';
  var date = new Date(timestamp * 1000);
  date.setHours(date.getHours() + 4); // Convert UTC to UTC+4
  return date.toUTCString();
}

function loadStripeScheduledSubscriptions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scheduled Subscriptions');
  var customersSheet = ss.getSheetByName('Customers');

  var customerDataRange = customersSheet.getRange(2, 1, customersSheet.getLastRow() - 1, 4).getValues();
  var customerCache = {};
  customerDataRange.forEach(function (row) {
    customerCache[row[0]] = { name: row[3], email: row[2] };
  });

  var headers = [
    "Schedule ID",
    "Customer ID",
    "Customer Name",
    "Customer Email",
    "Price ID",
    "Quantity",
    "Interval",
    "Amount",
    "Status",
    "Created (UTC)",
    "Start Date (UTC)",
    "End Date (UTC)",
    "End Behavior",
    "Currency",
    "Product"
  ];

  var apiKey = PropertiesService.getScriptProperties().getProperty('STRIPE_API_KEY');
  var apiOptions = {
    'method': 'get',
    'headers': { 'Authorization': 'Bearer ' + apiKey },
    'muteHttpExceptions': true
  };

  var lastSubscriptionScheduleId = null;
  var priceCache = {};
  var allRows = [];
  var productCache = {};

  while (true) {
    var apiUrl = 'https://api.stripe.com/v1/subscription_schedules?limit=100';
    if (lastSubscriptionScheduleId) {
      apiUrl += '&starting_after=' + lastSubscriptionScheduleId;
    }

    var response = UrlFetchApp.fetch(apiUrl, apiOptions);
    var json = JSON.parse(response.getContentText());
    var subscriptionSchedules = json.data;

    if (!subscriptionSchedules || subscriptionSchedules.length === 0) {
      break;
    }

    for (var i = 0; i < subscriptionSchedules.length; i++) {
      var schedule = subscriptionSchedules[i];
      lastSubscriptionScheduleId = schedule.id;

      var created = formatUtcDateToUtcPlus4(schedule.created);
      var startDate = 'N/A';
      var endDate = 'N/A';

      if (schedule.phases && schedule.phases[0]) {
        var phase = schedule.phases[0];
        startDate = formatUtcDateToUtcPlus4(phase.start_date);
        endDate = formatUtcDateToUtcPlus4(phase.end_date);
      }

      var endBehavior = schedule.end_behavior;
      var status = schedule.status;

      var customerID = schedule.customer;
      var customerInfo = customerCache[customerID] || { name: 'N/A', email: 'N/A' };
      var customerName = customerInfo.name;
      var customerEmail = customerInfo.email;

      var item = phase && phase.items && phase.items[0];
      var quantity = item ? item.quantity : '';
      var priceId = item ? item.price : '';

      var amount, interval, productId, productName;
      if (priceCache[priceId]) {
        amount = priceCache[priceId].amount;
        interval = priceCache[priceId].interval;
        productName = priceCache[priceId].productName;
      } else {
        var priceApiUrl = 'https://api.stripe.com/v1/prices/' + priceId;
        var priceResponse = UrlFetchApp.fetch(priceApiUrl, apiOptions);
        var priceData = JSON.parse(priceResponse.getContentText());

        amount = priceData.unit_amount / 100;
        interval = priceData.recurring ? priceData.recurring.interval : '';
        productId = priceData.product;

        if (productCache[productId]) {
          productName = productCache[productId];
        } else {
          var productApiUrl = 'https://api.stripe.com/v1/products/' + productId;
          var productResponse = UrlFetchApp.fetch(productApiUrl, apiOptions);
          var productData = JSON.parse(productResponse.getContentText());
          productName = productData.name || 'N/A';
          productCache[productId] = productName;
        }

        priceCache[priceId] = {
          amount: amount,
          interval: interval,
          productName: productName
        };
      }

      var row = [
        schedule.id || '',
        customerID || '',
        customerName,
        customerEmail,
        priceId || '',
        quantity || '',
        interval || '',
        amount || '',
        status || '',
        created,
        startDate,
        endDate,
        endBehavior || '',
        currency || '',
        productName
      ];

      allRows.push(row); // Add row data to allRows array
    }
  }

  sheet.clear();
  allRows.unshift(headers);
  sheet.getRange(1, 1, allRows.length, headers.length).setValues
