# Payoff Closer – Multi-Currency Payment Automation

## Overview

Payoff Closer is a Chrome extension developed to automate and streamline customer payment reconciliation and invoice settlement.

It retrieves customer billing information from a business management platform, calculates outstanding balances across multiple currencies, validates received payments, and automates the invoice closure process.

The solution supports both USD and EGP transactions, reducing manual calculations and accelerating customer payment processing.

## The Problem

Customer payment settlement originally involved manually searching for bills of lading (B/Ls), reviewing associated invoices, calculating outstanding balances, reconciling partial payments, and determining whether the customer's deposit was sufficient to cover the outstanding amount.

This process became particularly time-consuming when payments involved multiple currencies or partially settled invoices.

## Automated Workflow

### 1. Payment Information Input

The extension accepts:

- Bill of Lading (B/L) serial number.
- Amount received in USD.
- Amount received in EGP.

These values are entered based on the customer's bank deposit receipts.

### 2. Customer Billing Retrieval

- Searches for the specified B/L through the business management platform.
- Retrieves the associated billing information.
- Identifies the invoices and their corresponding outstanding balances.

### 3. Multi-Currency Payment Reconciliation

The extension calculates the outstanding balance for each currency independently.

- Identifies previously settled and partially settled invoices.
- Calculates remaining amounts due in USD and EGP.
- Accounts for partial payments made in either currency.
- Determines the total outstanding balance.
- Identifies whether the customer's payment is sufficient to settle the outstanding invoices.

### 4. Payment Validation & Status Indicators

Before proceeding with settlement, the extension evaluates the received payment against the outstanding balance.

It provides visual indicators for different payment scenarios:

- **Fully Settled:** Indicates that the associated billing has already been paid.
- **Exact Payment:** Confirms that the received amount covers the outstanding balance.
- **Underpayment:** Displays a warning indicating the remaining amount due.
- **Overpayment:** Highlights any excess amount received.

### 5. Automated Invoice Settlement

When the received payment is sufficient, the extension provides an option to proceed with automated settlement.

The process includes:

- Finalizing draft invoices.
- Registering payments against the associated invoices.
- Closing the customer's outstanding balance.
- Generating and downloading the payment receipt for printing.

## Key Features

- Automated B/L-based billing retrieval.
- Multi-currency reconciliation (USD & EGP).
- Partial payment calculations.
- Outstanding balance verification.
- Underpayment and overpayment detection.
- Automated invoice finalization and settlement.
- Payment receipt generation and downloading.
- Reduced manual calculations and repetitive administrative work.

## Technologies

- JavaScript
- Chrome Extension APIs
- Browser Automation
- DOM Manipulation
- Multi-Currency Financial Calculations

## Business Impact

The extension streamlines the customer payoff process by combining billing retrieval, multi-currency reconciliation, payment validation, and invoice settlement into a single workflow.

It reduces repetitive manual operations, minimizes calculation errors, and accelerates the processing of customer payments.

## Note

This repository contains a sanitized version of the project. Company-specific URLs, credentials, customer information, and confidential business configurations are excluded.
