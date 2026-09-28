
# PBTT Utility Management System

A Google Apps Script-based utility billing automation and validation system built for managing Electricity, Water, and LPG submission workflows in a structured spreadsheet environment.

This project is designed to:
- fetch raw utility data from external source spreadsheets
- map source columns to target utility tabs
- calculate and validate billing formulas
- detect anomalies and data inconsistencies
- enforce billing period controls and submission rules
- store submission records in a master database
- rotate and archive database files when usage approaches Google Sheets capacity limits

---

## Overview

The PBTT Utility Management System helps automate utility data processing and ensures that submissions are consistent, validated, and aligned with the master database requirements.

It is built around a Google Sheets workbook where users:
1. open the template file,
2. install or initialize the system,
3. fetch source data,
4. run utility formula calculations,
5. scan for anomalies,
6. submit the active PBTT dataset to the master database.

The system is specifically designed for business utility billing operations and supports structured validation rules for:
- Electricity
- Water
- LPG

---

## Key Features

### 1. Multi-Utility Support
The system manages three utility tabs:
- `Elec`
- `Water`
- `LPG`

Each tab uses a dedicated formula engine and data validation flow.

### 2. Dynamic Data Fetching
The system reads external spreadsheet data and copies it into the active utility tab using custom mapping rules.

Mapping is defined in:
```javascript
const FETCH_MAPS = {
  "Elec": { "J": "K", "AF": "L", "AI": "P" },
  "Water": { "J": "K", "AF": "L", "AI": "W" },
  "LPG": { "J": "K", "AF": "N", "AI": "P" }
};
```

This allows the script to align external source columns with the internal utility sheet structure.

### 3. Formula Automation
Each utility tab has a formula engine that computes:
- consumption
- VAT values
- variance values
- subtotal rows
- total rows
- rate comparisons

The script uses formula maps such as:
- `formulaMapElec`
- `formulaMapWater`
- `formulaMapLPG`

### 4. Validation & Anomaly Detection
The script performs comprehensive scanning before allowing submission.

It checks:
- missing required fields
- invalid values
- zero or negative consumption issues
- rate mismatches
- KA/SR/REG classification logic
- variance thresholds
- blind or placeholder remarks
- tenant/property mismatch against the master validation sheet

Logs are written to:
- `Basic Anomalies`
- `Client Rate Anomalies`

### 5. Master Database Submission Control
Before submission, the app verifies:
- the selected billing period exists and is active
- the period is not locked
- the property name matches the master validation file
- tenant names exist in the master property sheet
- tenant codes match the corresponding records
- all required fields are populated

Submission is recorded in the master database:
- `PBTT Submission`

### 6. Reference Number Generation
A unique reference number is generated for each active PBTT submission:
```javascript
Ref# + 6-character alphanumeric value
```

Example:
```text
Ref#A1B2C3
```

This is stored in the `Instructions` sheet and used for tracking and duplicate prevention.

### 7. Database Rotation & Capacity Protection
The system monitors the active database usage and rotates to a new backup database when the workbook approaches the Google Sheets cell limit.

Relevant configuration:
```javascript
const CELL_LIMIT_MAX = 10000000;
const CELL_ROTATION_LIMIT = 8000000;
```

This helps prevent the database from exceeding safe Google Sheets capacity thresholds.

### 8. Locking and Concurrency Safety
The script uses:
- `LockService`
- `PropertiesService`
- `SpreadsheetApp.flush()`

to protect against overlapping transactions and ensure secure sheet updates.

---

## Tech Stack

- Google Apps Script (JavaScript V8)
- Google Sheets
- Google Drive
- Google Apps Script UI
- Spreadsheet formulas and validation logic
- Google Drive file handling
- JSON-like PropertiesService configuration

---

## Project Structure

This project is a single-file Apps Script project using:
- `code.gs` – full backend logic
- spreadsheet tabs used by the workflow:
  - `Elec`
  - `Water`
  - `LPG`
  - `Instructions`
  - `Basic Anomalies`
  - `Client Rate Anomalies`
  - `dvPeriod`
  - `dvGen`
  - `KA_DATA`
  - `PBTT Submission`

---

## Important Configuration

The script contains several hardcoded IDs and environment-specific references. These should be reviewed and updated for your own workspace.

### Main IDs
```javascript
const PBTT_DB_ID = "1hMMUd4ho50HP63dc2fRAo--iK-m7YotamkKtsDGT_Us";
const BACKUP_REGISTRY_ID = "10-ywOh509BNRMd0C-Mb8b5gibbu62D_K8U8cWYcV59U";
const BACKUP_FOLDER_ID = "1aokNFrCuVdLWs4AylG7LNekCtfQ5B1-p";
```

### Validation/master reference spreadsheet
```javascript
const extId = "12OOOzMVeWPb6SKJyNu3tewSPKrbu3s93jJA3SmPNSY4";
```

### KA Data source
```javascript
const sourceId = "1jY-9FMha3x972o4Gz1d6DVD36d3ppjHW_WM1DHJz6ag";
```

These IDs are essential for:
- master data syncing
- validation lookup
- database rotation
- KA dataset integration

---

## Menu and Workflow

The app adds a custom Google Sheets menu:
```javascript
ui.createMenu("Utility Manager")
  .addItem("🛠️ Setup", "INSTALL_SYSTEM")
  .addSubMenu(ui.createMenu("⚡ Electricity")
    .addItem("1. Fetch Data", "masterFetchElec")
    .addItem("2. Run Formulas", "runFormulaElec"))
  .addSubMenu(ui.createMenu("💧 Water")
    .addItem("1. Fetch Data", "masterFetchWater")
    .addItem("2. Run Formulas", "runFormulaWater"))
  .addSubMenu(ui.createMenu("🔥 LPG")
    .addItem("1. Fetch Data", "masterFetchLPG")
    .addItem("2. Run Formulas", "runFormulaLPG"))
  .addSeparator()
  .addItem("📤 Submit Active PBTT", "recordActivePBTT")
```

### Typical process
1. Run `INSTALL_SYSTEM`
2. Sync lookup and reference data
3. Fetch source utility data
4. Run formulas for the chosen utility tab
5. Run global scan
6. Resolve anomalies
7. Submit active PBTT

---

## Setup Instructions

### 1. Prepare the Spreadsheet
Open the target Google Sheet and ensure the following tabs exist:
- `Instructions`
- `Elec`
- `Water`
- `LPG`

### 2. Add the Apps Script
Open:
- Extensions > Apps Script

Paste the script into the project.

### 3. Run the Initial Setup
From the menu, run:
- `Utility Manager > Setup`

This executes:
- `syncDataAcrossFiles()`
- `syncKAData()`
- `generateUniqueAlphanumericRef()`

### 4. Validate the Document
Check that:
- `Instructions!C7` contains the generated Ref#
- the active utility tabs are properly structured
- the source file URL is valid
- date fields in `Instructions!C26` and `Instructions!C27` are configured correctly

---

## Validation Rules

The script enforces several business rules before submission, including:

### Required Data Checks
- missing data in required columns
- blank `Property Name`
- blank billing date range
- blank `Tenant Code`
- invalid tenant/property mapping

### Formula and Rate Rules
- `L5` must be greater than `L6`
- `O` cannot be zero in certain conditions
- `P` must be valid positive value for non-exception rows
- calculated variance must not exceed thresholds without explanation

### Category Checks
Rows in `F` must be valid:
- `KA`
- `SR`
- `REG`
- `KA&SR`

### Property Matching
The user property in the sheet must match a valid tab in the master external validation file.

---

## Submission Process

The submission workflow is handled by:
```javascript
function recordActivePBTT()
```

This function:
- runs the full system scan
- blocks submission if anomalies remain
- validates billing period lock status
- checks master file relationships
- verifies tenant names and codes
- extracts values into the submission payload
- updates or appends to `PBTT Submission`

---

## Database Rotation Mechanism

The system includes:
```javascript
function rotateToNewDatabase(oldDb, oldSheet)
```

This creates a new database spreadsheet, moves it into a backup folder, and updates the active pointer using:
```javascript
PropertiesService.getScriptProperties().setProperty("ACTIVE_DB_ID", newFileId);
```

This protects the system from hitting sheet cell capacity limits.

---

## Important Notes

### Master File Protection
The script prevents execution on the master template file:
```javascript
function isMasterFileBlocked()
```

This helps ensure users create a copy before making changes.

### Internal Use
This system is intended for internal business operations and contains proprietary billing logic, database mapping logic, and sensitive utility data rules.

---

## Recommended Production Enhancements

For a stronger production setup, consider adding:
- user authentication for admin-level actions
- audit trail for all changes
- versioning for formula logic and mapping
- encrypted or restricted storage for sensitive customer data
- backup verification checks
- admin dashboard for database health and submission tracking
- clearer error messaging in the frontend UI

---

## License

This project is intended for internal business use. Please verify your organization’s licensing, data handling, and distribution policies before sharing or publishing the code externally.

---

## Summary

The PBTT Utility Management System is a robust Google Apps Script solution for:
- processing utility billing data
- validating submission quality
- enforcing rate and variance controls
- preventing duplicate or invalid submissions
- maintaining structured utility reporting and master database integrity
