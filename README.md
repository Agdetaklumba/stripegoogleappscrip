# Stripe Integration with Google Sheets

This repository contains Google Apps Script code for integrating Stripe with Google Sheets. The scripts fetch and process data related to customers, payments, invoices, subscriptions, and scheduled subscriptions, and store it in Google Sheets.

## Features
- Fetch and store Stripe customer data.
- Retrieve and manage payments, invoices, and subscriptions.
- Update Google Sheets with the latest Stripe data.
- Handle regular and scheduled subscriptions.
- Utility functions for processing and formatting data.

## Setup

1. Clone the repository:
   ```sh
   git clone https://github.com/Agdetaklumba/stripegoogleappscrip.git
   cd stripegoogleappscrtip
2.	Set up your Google Apps Script project:
	•	Open Google Sheets and create a new Google Apps Script project.
	•	Copy the contents of the src/ files into your script editor.
	3.	Configure the script properties:
	•	Set your Stripe API key in the script properties
	4.	Run the script:
	•	Use the Google Apps Script interface to run the scripts as needed.
	5.	Setup triggers:
    	.	Setup hourly/minutes triggers for updating function

Usage

Customers

	•	loadStripeCustomers: Fetches and stores Stripe customer data in the Customers sheet.
	•	fetchPaymentsData: Retrieves payment data and aggregates it by customer ID.

Invoices

	•	loadStripeInvoices: Fetches and stores invoice data in the Invoices sheet.
	•	updateStripeInvoices: Updates existing invoice data.

Subscriptions

	•	loadStripeSubscriptions: Fetches and stores subscription data in the Subscriptions sheet.
	•	updateStripeSubscriptions: Updates existing subscription data.

Payments

	•	loadStripePayments: Fetches and stores payment data in the Payments sheet.
	•	updateStripePayments: Updates existing payment data.

Scheduled Subscriptions

	•	loadStripeScheduledSubscriptions: Fetches and stores scheduled subscription data.
	•	updateStripeScheduledSubscriptions: Updates existing scheduled subscription data.
