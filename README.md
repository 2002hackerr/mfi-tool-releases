# Row IB Investigation Tool v6.2.5 — Official Documentation

The **Row IB Investigation Tool** is a high-performance framework designed for MFI (Marketplace Facilitator Invoice) investigation. It automates the reconciliation of shortages, cross-PO overages, and REBNI adjustments against claim datasets.

---

## 🕹️ Main Control Panel: Buttons & Functions

### 1. File Selection & Setup
*   **LOAD CLAIMS**: Imports the primary dataset containing ASINs, POs, and PQVs that require investigation. Supports `.xlsx` and `.csv`.
*   **LOAD REBNI**: Imports the Received But Not Invoiced (REBNI) data. The tool uses this to identify units physically received but missing from invoices.
*   **LOAD INVOICES**: Imports the master invoice dataset.
*   **SHIPMENT MASTER**: Imports shipment-level data used to identify total quantities shipped for a specific SID.

### 2. Core Execution
*   **PROCESS**: Starts the investigation. It scans the Claims file and begins the recursive investigation for each ASIN.
*   **STOP INVESTIGATION**: A safety-critical button that immediately terminates the recursive engine and background threads without crashing the application.
*   **RESET**: Completely wipes the tool's memory. It clears all file paths, deletes the engine cache (SIDs/Barcodes), and resets the UI for a fresh ticket.

### 3. Utilities & Portals
*   **PORTAL**: Launches the **MFI Unique Summary Portal**. This is a built-in HTML dashboard that provides:
    *   AI-powered investigation assistance.
    *   Visual data summaries.
    *   Export utilities for final reporting.
*   **FAST FETCH**: A high-speed utility for bulk data. Input a **Vendor Code**, and the tool will instantly filter the master Cloud/Local datasets to create a smaller, workable file for the current investigation.
*   **HELP**: Displays a quick-start guide and version information.

---

## 🛠️ Advanced Investigation Features

### 1. Recursive Investigation Engine
*   **Multi-Level Depth**: The tool doesn't just check the first match; it follows the "shortage chain" across multiple invoices (Level 0, Level 1, etc.) until the PQV is fully accounted for.
*   **Branch Budgeting**: Every sub-investigation is limited by the `matched_quantity` from its parent, preventing "phantom shortages."
*   **Sequential Logic**: Investigates all matches for a given SID/PO/ASIN one by one to ensure no unit is left uncounted.

### 2. Cross PO Automation (Case 1, 2, 3)
*   The tool automatically detects units received under a different PO for the same SID/ASIN:
    *   **Case 1**: Units received in a PO where the ASIN wasn't originally claimed.
    *   **Case 2**: Units received in a PO that has not been invoiced yet.
    *   **Case 3**: General overages in a different PO that offset the current shortage.

### 3. Smart Global Caching
*   **SID/Barcode Memory**: Once you enter a SID or Barcode for an invoice, the tool remembers it. If that same invoice appears again for a different ASIN, the tool automatically reuses the data, saving hours of manual entry.

### 4. Dynamic Preview & Manual Mode
*   **Live Preview Panel**: A non-modal window that shows the investigation rows as they are generated. You can edit the "Remarks" column in real-time.
*   **Manual Override Dialogs**: If the tool finds multiple matches or needs a SID, it opens a non-modal dialog. You can minimize these or move them to the side while you check other data.

---

## 📐 Accounting & Formulas
The tool strictly follows the **ROW IB Accounting Standards**:
*   **Total Accounted** = `Found Shortage` + `REBNI Available` + `EX Adjustments`.
*   **Shortage Remark**: Always formatted as:  
    *"Found {X} units short as loop started from {Y} matched qty..."*
*   **REBNI Rule**: Uses the first row (`rebni_rows[0]`) only to ensure data integrity.

---

## 🎨 Personalization (Theme Engine)
Accessible via the UI, users can switch between premium, high-contrast themes:
*   **Dark Mode (Default)**: Sleek, low-strain interface.
*   **Ocean Blue**: Professional deep-sea palette.
*   **Forest Green**: High-readability nature theme.
*   **Sunset Orange / Purple Midnight**: Vibrant, high-contrast designs.

---

## 📋 Build Specifications
*   **Version**: 6.2.5.0
*   **Type**: Standalone Single Executable (.exe)
*   **OS**: Windows 10/11 Compatible
*   **Security**: Secured logic with no external data leakage.
