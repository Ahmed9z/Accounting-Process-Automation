# B/L Processing & Invoice Generation Automation

A Chrome extension that automates shipment preparation, cost entry, and customer invoice generation in Logistaas.

## Overview

This project was developed to streamline repetitive shipment-processing tasks by automating several stages of the workflow, from searching for a Bill of Lading (B/L) to entering predefined charges and generating customer invoices.

Instead of manually navigating through multiple shipment tabs, adding individual cost entries, configuring invoice details, and downloading PDFs, the extension executes these steps through an automated, sequential workflow.

It also supports batch processing, allowing multiple B/Ls to be processed consecutively.

## The Problem

Shipment processing involves several repetitive operations:

- Searching for individual B/L records.
- Reviewing and completing shipment information.
- Adding multiple predefined charges with different quantities, units, and currencies.
- Creating separate invoices based on shipment type and currency.
- Downloading and organizing generated invoice PDFs.

Performing these operations manually increases processing time and creates opportunities for data-entry errors and inconsistencies.

## How It Works

### 1. Shipment Search

The extension searches Logistaas using the provided B/L number and navigates to the corresponding shipment record.

### 2. Shipment Preparation

Before entering charges, the automation checks and updates selected shipment details where necessary:

- Sets ETA and arrival dates when the respective fields are empty.
- Sets the destination country when it has not been specified.
- Checks the shipping line and selects a predefined option if missing.
- Adds a container when no container record exists.

Existing information is generally preserved when the relevant fields are already populated.

### 3. Automated Cost & Revenue Entry

The extension automatically generates predefined charge entries according to the selected shipment type.

For each charge, it handles:

- Charge description selection.
- Quantity.
- Unit of measurement.
- Currency selection (USD or EGP).
- Selling rate.
- Cost field initialization.

Charge descriptions are matched against Logistaas service-charge records to select the corresponding system entries.

Different predefined charge configurations are used depending on the shipment category.

### 4. Automated Invoice Generation

After saving the shipment charges, the extension proceeds to invoice generation.

It automatically:

- Applies the configured tax scheme.
- Selects the appropriate invoice currency.
- Selects the relevant charge items for each invoice.
- Populates the external invoice reference when provided.
- Sets the invoice status to Draft.
- Enables the invoice logo option.

Separate invoices are generated according to the predefined invoice plans for each shipment category.

### 5. PDF Download & Organization

Generated invoice PDFs are downloaded automatically and organized into a structured folder hierarchy:

```text
Batch/
├── BL_Number_1/
│   ├── invoice_1_USD.pdf
│   ├── invoice_2_USD.pdf
│   └── invoice_3_EGP.pdf
│
├── BL_Number_2/
│   ├── invoice_1_USD.pdf
│   └── invoice_2_EGP.pdf
```

Chrome's download handling is used to assign the intended filenames and folder paths.

### 6. Batch Processing & Progress Tracking

The extension supports sequential processing of multiple B/Ls.

It maintains its processing state using Chrome's local storage, allowing the workflow to continue across page navigations and service-worker restarts.

Progress notifications provide feedback about the current operation, while error handling identifies the B/L associated with a failed step.

## Key Features

- Automated B/L searching and shipment navigation.
- Conditional shipment information updates.
- Automatic container creation when required.
- Predefined multi-charge entry configurations.
- Automatic charge matching with Logistaas records.
- Support for multiple shipment categories.
- Multi-currency invoice generation (USD and EGP).
- Automated invoice item selection.
- Invoice PDF downloading and folder organization.
- Sequential batch processing.
- Persistent workflow state and error notifications.

## Technologies Used

- **JavaScript** – Core automation logic and workflow management.
- **Chrome Extension APIs** – Messaging, local storage, and download management.
- **DOM Manipulation** – Interaction with Logistaas interface elements.
- **MutationObserver** – Detection of dynamically loaded interface elements.
- **Fetch API** – Retrieval of available service-charge records.
- **HTML & CSS** – Extension interface and user notifications.

## Important Notes

- The extension relies on predefined charge configurations and invoice rules. These must be reviewed and adapted to the relevant business requirements before use.
- It is designed around the structure and behavior of the target Logistaas interface. Changes to that interface may require adjustments to the automation selectors and workflow.
- Shipment updates, charge entries, and invoice creation should be verified before being treated as final business records.
- The project is intended for authorized use with the relevant Logistaas environment.

## Project Purpose

This project demonstrates the practical application of browser automation to accounting and shipping operations, combining repetitive data-entry automation, predefined business rules, multi-currency processing, and document management into a single workflow.

The main objective is to reduce manual processing time, improve consistency, and streamline high-volume shipment and invoice operations.
