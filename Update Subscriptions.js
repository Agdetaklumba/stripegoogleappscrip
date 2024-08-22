function updateSubscriptionsboth(){
  updateStripeSubscriptions();
  updateStripeScheduledSubscriptions();
}

function updateStripeSubscriptions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Subscriptions');
  
  // Fetch customers for name and email mapping
  var customerSheet = ss.getSheetByName('Customers');
  var customersData = customerSheet.getRange(1, 1, customerSheet.getLastRow(), 4).getValues();
  var customerMap = {};
  customersData.forEach(function(row) {
    customerMap[row[0]] = { name: row[3], email: row[2] }; 
  });

  var API_KEY = PropertiesService.getScriptProperties().getProperty('STRIPE_API_KEY');
  var apiOptions = {
    'method' : 'get',
    'headers' : {'Authorization' : 'Bearer ' + API_KEY},
    'muteHttpExceptions': true
  };
  
  var existingSubscriptionIds = sheet.getRange(1, 1, sheet.getLastRow()).getValues().flat();
  
  // Fetch subscriptions created in the last 10 days
  var today = new Date();
  var startOfToday = new Date(today.setDate(today.getDate() - 10));
  var startTimestamp = Math.floor(startOfToday / 1000);
  var apiUrl = 'https://api.stripe.com/v1/subscriptions?created[gte]=' + startTimestamp;

  var allRows = [];
  
  do {
    var response = UrlFetchApp.fetch(apiUrl, apiOptions);
    var json = JSON.parse(response.getContentText());
    var subscriptions = json.data;

    for (var i = 0; i < subscriptions.length; i++) {
      var subscription = subscriptions[i];
      if (subscription && subscription.id && existingSubscriptionIds.indexOf(subscription.id) === -1) {
        var row = processSubscription(subscription, customerMap);
        allRows.push(row);
      }
    }

    apiUrl = json.has_more ? 'https://api.stripe.com/v1/subscriptions?created[gte]=' + startTimestamp + '&starting_after=' + subscriptions[subscriptions.length - 1].id : null;

  } while (apiUrl);

  if (allRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
  }
}

function processSubscription(subscription, customerMap) {
    var customerID = subscription.customer;
    var customerName = customerMap[customerID] ? customerMap[customerID].name : 'N/A';
    var customerEmail = customerMap[customerID] ? customerMap[customerID].email : 'N/A';

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
        customerName || '',
        customerEmail || '',
        subscription.plan && subscription.plan.id || '',
        subscription.quantity || '',
        subscription.plan && subscription.plan.interval || '',
        totalAmount,
        subscription.status || '',
        subscription.created ? new Date(subscription.created * 1000).toUTCString() : '',
        subscription.start_date ? new Date(subscription.start_date * 1000).toUTCString() : '',
        subscription.current_period_start ? new Date(subscription.current_period_start * 1000).toUTCString() : '',
        subscription.current_period_end ? new Date(subscription.current_period_end * 1000).toUTCString() : '',
        subscription.canceled_at ? new Date(subscription.canceled_at * 1000).toUTCString() : 'N/A',
        subscription.cancel_at_period_end ? 'Yes' : 'No',
        subscription.ended_at ? new Date(subscription.ended_at * 1000).toUTCString() : 'N/A',
        customerName || 'N/A',
        subscription.cancel_at ? new Date(subscription.cancel_at * 1000).toUTCString() : 'N/A',
        subscription.items.data[0].plan.product || 'N/A'
    ];

    return row;
}

function updateStripeScheduledSubscriptions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scheduled Subscriptions');
  var existingData = sheet.getDataRange().getValues();

  var existingSchedulesMap = {};
  for (var i = 1; i < existingData.length; i++) {
    var scheduleId = existingData[i][0];
    existingSchedulesMap[scheduleId] = existingData[i];
  }

  var API_KEY = PropertiesService.getScriptProperties().getProperty('STRIPE_API_KEY');
  var apiOptions = {
    'method': 'get',
    'headers': { 'Authorization': 'Bearer ' + API_KEY },
    'muteHttpExceptions': true
  };

  var lastSubscriptionScheduleId = null;
  var customerCache = {};
  var priceCache = {};
  var productCache = {};
  var updatedRows = [existingData[0]]; // Keep the headers

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

      var created = new Date(schedule.created * 1000).toUTCString();
      var endBehavior = schedule.end_behavior;
      var status = schedule.status;

      var customerID = schedule.customer;
      var customerName, customerEmail;

      if (customerCache[customerID]) {
        customerName = customerCache[customerID].name;
        customerEmail = customerCache[customerID].email;
      } else {
        var customerApiUrl = 'https://api.stripe.com/v1/customers/' + customerID;
        var customerResponse = UrlFetchApp.fetch(customerApiUrl, apiOptions);
        var customerData = JSON.parse(customerResponse.getContentText());
        customerName = customerData.name || 'N/A';
        customerEmail = customerData.email || 'N/A';

        customerCache[customerID] = {
          name: customerName,
          email: customerEmail
        };
      }

      var phase = schedule.phases && schedule.phases[0];
      var currency = phase ? phase.currency : '';
      var item = phase && phase.items && phase.items[0];
      var quantity = item ? item.quantity : '';
      var startDate = phase ? new Date(phase.start_date * 1000).toUTCString() : '';
      var endDate = phase ? new Date(phase.end_date * 1000).toUTCString() : '';

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

      if (existingSchedulesMap[schedule.id]) {
        var index = existingData.findIndex(function(dataRow) {
          return dataRow[0] === schedule.id;
        });
        if (index !== -1) existingData[index] = row;
        delete existingSchedulesMap[schedule.id];
      } else {
        updatedRows.push(row);
      }
    }
  }

  updatedRows = updatedRows.concat(existingData.slice(1));
  sheet.clear();
  sheet.getRange(1, 1, updatedRows.length, updatedRows[0].length).setValues(updatedRows);
}
