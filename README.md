# PBTT Utility ManagementS System

## Introduction
The **PBTT Utility ManagementS System** is an advanced automation framework designed to manage, calculate, and audit utility billing data (Electricity, Water, and LPG). It streamlines the transition from raw consumption data to final **Projected Billing to Tenants (PBTT)**.

The system acts as a bridge between various data sources and a master database, ensuring that all calculations follow strictly defined formulas and that data anomalies are flagged before final submission.

### Core Features:
*   **Dynamic Data Fetching:** Multi-tab support for Electricity, Water, and LPG with custom source-to-target column mapping.
*   **Automated Formula Engine:** Injects complex billing logic, VAT calculations, and consumption variances directly into sheets.
*   **Global Anomaly Scanner:** A two-tier auditing system that detects "Basic Anomalies" (missing data, negative values) and "Client Rate Anomalies" (Key Account identification vs. Database records).
*   **Smart Database Rotation:** Automatically detects when the master database is nearing Google’s 10-million cell limit and self-rotates to a new backup file.
*   **Persistence & Integrity:** Generates unique Alphanumeric Reference Numbers for every submission to prevent duplicate processing.

---

## Installation Instructions

### 1. Spreadsheet Prerequisites
This system requires a specific file structure to function:
*   **Active PBTT Template:** The file where this script resides.
*   **Master Database (`PBTT_DB_ID`):** A centralized sheet to record all utility submissions.
*   **KA Reference File:** A master lookup table for Key Account (KA) tenants and categories.
*   **Backup Registry (`BACKUP_REGISTRY_ID`):** To track rotated database files.

### 2. Configuration & Mapping
Modify the global constants at the top of the script to match your organization’s environment:
*   `PBTT_DB_ID`: Your Master Database ID.
*   `BACKUP_FOLDER_ID`: The Drive folder where new database rotations will be saved.
*   `FETCH_MAPS`: Define which columns from the source link should map to specific columns in the Utility tabs.

### 3. Deployment
1.  Open your Google Sheet and navigate to **Extensions > Apps Script**.
2.  Paste the provided code into the editor.
3.  Run the `INSTALL_SYSTEM` function once. This will:
    *   Sync your lookup tables (`dvPeriod`, `dvGen`, `KA_DATA`).
    *   Generate a unique Reference Number in the `Instructions` tab.
    *   Cleanup startup triggers to optimize performance.

---

## Usage Examples

### Executing the Utility Workflow
The system is controlled via a custom **"Utility Manager"** menu that appears upon opening the sheet.

#### Step 1: Fetching Raw Data
```javascript
/**
 * Example: Master Fetch Pattern
 * This function handles the initialization, 
 * database validation, and external data pulling.
 */

function masterFetchElec() {
  INITIALIZE_SYSTEM_BUTTON(); // Syncs master lookup tables
  fetchElec();                // Opens source URL and maps consumption data
}
```

#### Step 2: Running Formula Calculations
After data is fetched, use the **Run Formulas** menu. This logic protects "Theoretical" or "Fix Rate" entries while applying standard billing math to others.

#### Step 3: Scanning for Errors
Run the **Global Scan** to identify issues across all three utility types.
*   **Standard Errors:** "Consumption > 0 but Amount is 0."
*   **Variance Alerts:** Flagging any consumption change +/- 30% compared to the previous period.

#### Step 4: Final Submission
Use **Submit Active PBTT**. The system will:
1.  Check if the billing period is currently "Locked" by Admin.
2.  Validate that no mandatory fields are blank.
3.  Check the Master Database cell limit.
4.  Record the data or overwrite existing entries if the Reference Number matches.

---

## Tech Stack
*   **Language:** Google Apps Script (JavaScript V8)
*   **Optimization:** LockService (Concurrency management), CacheService (Performance), PropertiesService (State management).
*   **Database:** Distributed Google Sheets Architecture.

---

> **Warning**  
> **- For MCD Internal Use Only**  
> This system contains proprietary billing logic and sensitive utility consumption data. Unauthorized distribution of this code or its database IDs is strictly prohibited.
