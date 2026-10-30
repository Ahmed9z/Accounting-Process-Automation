# Payoff Closer – Multi-Currency Payment & Invoice Settlement Automation

## Overview

Payoff Closer is a Chrome extension developed to automate and streamline customer payment reconciliation and invoice settlement.

It retrieves customer billing information from a business management platform, calculates outstanding balances across multiple currencies, validates received payments, handles bank-specific invoice information, and automates the invoice closure process.

The solution also maintains a reference list of bills of lading (B/Ls) with exceptional charges, allowing the user to identify unusual billing conditions before proceeding with payment settlement.

## The Problem

Customer payment settlement originally involved manually searching for bills of lading (B/Ls), reviewing associated invoices, calculating outstanding balances, reconciling partial payments, and determining whether the customer's deposit was sufficient to cover the outstanding amount.

Additional manual steps were required to identify exceptional charges and enter the appropriate bank information on invoices.

This process became particularly time-consuming when payments involved multiple currencies, partially settled invoices, or unusual billing conditions.

## Automated Workflow

### 1. Payment Information Input

The extension accepts the following information:

- Bill of Lading (B/L) serial number.
- Amount received in USD.
- Amount received in EGP.
- Bank name where the customer deposited the payment.

The bank information is used to populate the corresponding bank details on the invoices.

### 2. Customer Billing Retrieval

- Searches for the specified B/L through the business management platform.
- Retrieves the associated billing information.
- Identifies the invoices and their corresponding outstanding balances.

### 3. Multi-Currency Payment Reconciliation

The extension calculates the outstanding balance for each currency independently.

- Identifies previously settled and partially paid invoices.
- Calculates remaining amounts due in USD and EGP.
- Accounts for partial payments made in either currency.
- Determines the total outstanding balance.
- Evaluates whether the customer's payment is sufficient to settle the invoices.

### 4. Exceptional Charges Detection

Some B/Ls occasionally contain unusual charges, such as War Risk Charges.

To account for these exceptions, the extension uses an external Excel reference sheet containing the B/L numbers associated with unusual charges.

- Reads and maintains a reference list of exceptional B/Ls.
- Checks the entered B/L number against the reference sheet.
- Displays a warning whenever a B/L with exceptional charges is identified.
- Alerts the user before proceeding with payment settlement.

This provides an additional verification step for infrequent billing exceptions without requiring the user to memorize or manually check every exceptional B/L.

### 5. Payment Validation & Status Indicators

Before proceeding with settlement, the extension evaluates the received payment against the outstanding balance.

It provides visual indicators for different payment scenarios:

- **Fully Settled:** Indicates that the associated billing has already been paid.
- **Exact Payment:** Confirms that the received amount covers the outstanding balance.
- **Underpayment:** Displays a warning indicating the remaining amount due.
- **Overpayment:** Highlights any excess amount received.

### 6. Automated Invoice Settlement

When the received payment is sufficient, the extension provides an option to proceed with automated settlement.

The process includes:

- Finalizing draft invoices.
- Populating the relevant bank information.
- Registering payments against the associated invoices.
- Closing the customer's outstanding balance.
- Generating and downloading the payment receipt for printing.

## Key Features

- Automated B/L-based billing retrieval.
- Multi-currency reconciliation (USD & EGP).
- Partial payment calculations.
- Outstanding balance verification.
- Bank-specific invoice information.
- Exceptional B/L detection through Excel reference data.
- War Risk and other unusual charge warnings.
- Underpayment and overpayment detection.
- Automated invoice finalization and settlement.
- Payment receipt generation and downloading.

## Technologies

- JavaScript
- Chrome Extension APIs
- Browser Automation
- DOM Manipulation
- Excel / Spreadsheet Processing
- Multi-Currency Financial Calculations

## Business Impact

Payoff Closer combines billing retrieval, multi-currency reconciliation, exceptional charge detection, payment validation, and invoice settlement into a streamlined workflow.

It reduces repetitive manual operations, minimizes calculation errors, improves awareness of unusual billing conditions, and accelerates customer payment processing.

## Note

This repository contains a sanitized version of the project. Company-specific URLs, credentials, customer information, internal configurations, and confidential business data are excluded.
