# Bank Reconciliation & Exception Detection Tool

A Python-based reconciliation application that automates the comparison of bank statement deposits against Odoo payment records, identifies discrepancies, and generates a detailed Excel reconciliation report.

## Overview

Bank reconciliation can become time-consuming when dealing with multiple bank accounts, currencies, and large volumes of payment records.

This tool streamlines the process by automatically matching bank deposits with Odoo records using reference numbers, transaction amounts, and dates. It also identifies potential errors and exceptions that may require manual investigation.

The application runs locally through a browser-based interface and generates an annotated Excel workbook containing reconciliation results.

## Key Features

- **Automated Reconciliation:** Matches bank deposits against Odoo payment records.
- **Multi-Sheet Support:** Processes workbooks containing multiple bank and Odoo sheets.
- **Reference-Based Matching:** Prioritizes exact reference matches.
- **Amount & Date Fallback:** Identifies potential matches when exact reference matching is unavailable.
- **Discrepancy Detection:** Flags transactions where references match but amounts differ.
- **Typo Detection:** Identifies potentially mistyped reference numbers using similarity checks.
- **Cross-Bank Detection:** Highlights possible payment records allocated to a different bank account.
- **Orphan Detection:** Identifies Odoo entries without a corresponding bank deposit.
- **Excel Reporting:** Generates a structured reconciliation report with highlighted statuses, transaction details, and summary statistics.

## How It Works

### 1. Upload Source Files

The application accepts two Excel workbooks:

- **Bank Statement Workbook:** Contains bank deposit transactions, with currency-specific sheets where applicable.
- **Odoo Export Workbook:** Contains payment records exported from Odoo, organized by bank account.

An optional period label can be provided for the report header.

### 2. Automated Matching

The reconciliation engine applies multiple matching stages:

1. **Exact Reference Matching:** Compares normalized reference numbers between bank and Odoo records.
2. **Amount Discrepancy Detection:** Flags reference matches where transaction amounts differ.
3. **Amount & Date Matching:** Identifies tentative matches based on transaction amount and date.
4. **Reference Typo Detection:** Searches for similar reference numbers that may contain typing errors.
5. **Cross-Bank Detection:** Identifies possible references recorded under another bank's Odoo sheet.
6. **Orphan Identification:** Lists Odoo payment entries without a corresponding bank-side match.

Potential typo and cross-bank matches are presented as investigation hints rather than automatically confirmed matches.

### 3. Generate Reconciliation Report

The application creates an annotated Excel workbook containing:

- Bank-side reconciliation totals.
- Odoo-side reconciliation totals.
- Match-confidence breakdown.
- Unmatched bank deposits.
- Reference matches with amount discrepancies.
- Possible reference typos.
- Possible wrong-bank entries.
- Odoo entries without matching bank deposits.

The original bank workbook is preserved as the basis for the output, with reconciliation annotations and a dedicated report worksheet added.

## Match Classification

| Status | Description |
|---|---|
| MATCHED | Exact reference match with an acceptable amount difference. |
| FLAG | Reference matched, but the transaction amounts differ. |
| TENTATIVE | Potential match based on amount and date. |
| UNMATCHED | No corresponding Odoo record identified. |

**Note:** Summary statistics group exact, flagged, and tentative records under the matched count. Flagged and tentative records should still be reviewed before being treated as confirmed reconciliations.

## Technology Stack

- Python
- Flask
- OpenPyXL
- HTML
- CSS
- JavaScript
- Excel automation and data processing
- Levenshtein distance for reference similarity checks

## Running the Application

The application launches a local Flask server and opens its interface in the default browser.

1. Place `BankRecon.py` and the launcher batch file in the application directory.
2. Ensure the required embedded Python runtime (`python-embed`) is available in the expected directory.
3. Run the batch file.
4. Upload the bank statement and Odoo export workbooks.
5. Enter an optional reporting period.
6. Start reconciliation.
7. Review the results and download the generated Excel workbook.

The application uses `localhost` and is configured to bind to `127.0.0.1`.

## Data Handling & Privacy

- Reconciliation is performed locally using Python.
- Uploaded workbooks are temporarily stored on the local machine during processing.
- Temporary source files are deleted after the reconciliation job finishes.
- Generated output workbooks are stored locally for download.
- No external cloud processing or third-party reconciliation service is required by the application code.

Users should handle source workbooks and generated reports according to their organization's data protection policies.

## Project Purpose

This project demonstrates the application of Python automation to practical accounting workflows, combining financial data processing, transaction matching, exception detection, and automated Excel reporting.

It aims to reduce repetitive manual reconciliation work and make discrepancies easier to identify and investigate.
