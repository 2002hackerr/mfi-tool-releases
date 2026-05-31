# MFI Investigation Tool v7.7 Release Notes

### ✨ New Features & Changes
* **Shipment Level Analysis Module**: Integrated dynamic Shipment Level grouping and cell merging logic (similar to ASIN Level Analysis) directly into the Excel writer. Shipments are now clearly grouped by PO and Invoice.
* **Genuine Overages Tracking**: Added comprehensive tracking for 'Genuine Overages' (items fully received but completely missing from the invoice dataset) within the Shipment Level Analysis.
* **Summary Metrics Embedded**: Automatically calculates and appends total invoiced quantities, total received quantities, shortages, and overages in the right-side columns of both the ASIN Level and Shipment Level sheets.
* **Auto Mode Cross PO Fix**: Resolved an issue in fully automated mode where the engine would skip deep-diving into sibling invoices if a Cross PO's budget was prematurely met. The tool now correctly iterates through all sibling matches for visibility.
* **Vault Interactive Fix**: Fixed a bug where opening a pending Cross PO from the Vault bypassed the manual interactive popups. Vault investigations now seamlessly drop you back into the interactive step-by-step GUI flow.

### 🐛 Bug Fixes
* **Missing Alignment Imports**: Resolved a bug where saving the main panel threw an error due to missing Alignment imports in the openpyxl dependency chain.
* **Cross PO Vault Disruption**: Fixed an internal behavior mismatch where _start_vault_investigation lacked context variables, previously causing single-row output dumps without interactive continuity.

