# Automated Invoice Processing

## Overview

A collection of automation scripts designed to streamline invoice creation by extracting invoice information from Logistaas (SaaS), mapping its charges to their corresponding Odoo products, and preparing structured data for bulk invoice creation.

The solution minimizes repetitive manual data entry, improves consistency, and significantly reduces the time required to process large volumes of invoices.

## The Problem

Invoice creation originally required manually retrieving invoice PDFs from Logistaas, reviewing their contents, identifying the corresponding charges in Odoo, and entering each line individually.

Since the two systems use different charge descriptions and product naming conventions, this process also required manual charge mapping, making it time-consuming and prone to human error.

## Automated Workflow

The solution consists of three main stages:

### 1. Invoice Extraction & Charge Mapping (Logistaas Reader)

- Reads invoice numbers from an Excel workbook named `Reader`.
- Processes invoice numbers sequentially, cell by cell.
- Searches for each invoice through the Logistaas interface.
- Retrieves and extracts the corresponding invoice PDF content.
- Identifies invoice charges and maps them to their corresponding Odoo products and charge descriptions.
- Exports the processed information into a separate Excel workbook named `Logistaas Results`.

### 2. Data Structuring & Sorting

Once invoice extraction is complete, a separate script (merger) processes the `Logistaas Results` workbook.

- Sorts and restructures the extracted invoice data.
- Arranges invoice lines and associated information according to Odoo's import requirements.
- Produces a structured output file ready for bulk uploading.

### 3. Odoo Bulk Invoice Import

The final structured Excel file is uploaded through Odoo's native import functionality.

This eliminates the need to manually create invoice lines or simulate individual keyboard entries.

## Development Evolution

An earlier implementation relied on browser automation to simulate human typing, entering invoice information directly into Odoo fields.

Although this approach was considerably faster than manual data entry, it occasionally introduced inconsistencies, such as unintended spaces that resulted in empty invoice lines.

A validation process was implemented to identify these discrepancies by exporting draft invoices from Odoo and comparing them against the expected invoice data in Excel before posting.

The workflow was subsequently improved by moving towards structured Excel generation and Odoo's native bulk-import functionality, providing a cleaner and more consistent approach.

## Key Benefits

- Automated retrieval of invoice information.
- Automated mapping between different charge naming conventions.
- Reduced repetitive manual data entry.
- Structured and standardized invoice data.
- Bulk invoice creation through Odoo.
- Improved accuracy and reduced opportunities for manual errors.
- Support for processing multiple invoices sequentially.

## Technologies & Tools

- Python
- Excel / Spreadsheet Processing
- Browser Automation
- PDF Data Extraction
- Logistaas
- Odoo ERP

## Note

This project is presented as a sanitized demonstration of business process automation. Company-specific configurations, internal URLs, credentials, and confidential business information are excluded.
