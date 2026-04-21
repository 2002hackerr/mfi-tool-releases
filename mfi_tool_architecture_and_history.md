# MFI Investigation Tool (ROW IB) - Architecture & History Documentation
**Version:** v5.4.0
**Date Recorded:** April 2026

## Core Objective
The tool investigates short-received quantities (PQV/claim quantity) on Amazon inbound shipments. It automatically or manually traces missing units through complex supply chain hierarchies by matching ASINs, SIDs, and POs across multiple invoices, shipments, and cross-POs.

## Core Rules & Logic
1. **Branch Budgeting (The Golden Rule)**
   - When a parent shipment claims a shortage, the tool looks for matched invoices.
   - For every matched invoice $M_i$, the "branch budget" is STRICTLY the matched quantity `mtc_qty` from that specific invoice.
   - The recursive chain can NEVER contribute more units back to the parent than the branch budget. `contribution = min(branch_budget, shortage_found_in_child)`.

2. **REBNI Evaluation (The Single Row Source of Truth)**
   - Only the FIRST matching row `rebni_rows[0]` is used. Summing units across multiple REBNI rows is strictly forbidden to prevent artificial inflation.
   - **Accounted Units Formula:** `Total Accounted = Shortage (SR) + REBNI Available + EX Adjustments`.

3. **Termination & Early Gateways (Direct Shortage)**
   - If the REBNI data indicates the shipment is directly "Short Received" (i.e. `accounted >= claiming_pqv` or similar), the branch immediately terminates with "SR".
   - Same for "Root Cause" or "Phase 1 - Fully Accounted" scenarios. 
   - No invoice matching is performed if the shortage is already explained at the gateway.

4. **Cross PO Anomalies (Cases 1, 2, and 3)**
   - The tool actively checks for units that were received against the exact same ASIN and SID, but billed to a *different* PO.
   - It captures confirmed overages from these Cross POs and applies them to offset the missing PQV.
   - Deep nested investigations require user validation via a popup (`CrossPODialog`) when in Manual Mode.

5. **Loop Catching (Cache)**
   - A `loop_cache` uniquely keyed by `(SID, PO, ASIN, Invoice)` tracks investigated states. It prevents infinite recursion and stores tuples: `(rows, total_accounted)`.

6. **UI Conventions (Tkinter)**
   - The UI is strictly non-modal (Parallel interaction allowed).
   - "Preview Panel" allows dynamic editing of values which globally feed into the `user_overrides` dict.
   - "Lookup Inv Qty" exists to spot-check the raw invoice data quickly.

## File Dependencies
The script relies heavily on localized matching from 3 input files:
1. **Claims Sheet:** Initiator data (Barcode, SID, PO, ASIN, PQV).
2. **REBNI Results:** Shipment receipt statuses (Qty unpacked, adjusted, overages).
3. **Invoice Search:** Relational mapping indicating which invoices generated which shipments for an ASIN.

## Historical Fixes (v4.0 - v5.4.0)
- Addressed infinite depth issues by strictly enforcing the `max_depth` override.
- Introduced `ManualLevelDialog` for interactive stepwise mapping (`_man_step`), forcing manual alignment of the recursive depth-first tree.
- Cost Price (CP) Validation checks for price drift >10%.
- Headers are automatically corrected using Levenshtein distance/fuzzy matching.

## Future Translation Notes (for Web/JS Migration)
- Data parsing in Web should heavily favor **streaming CSV parsing (PapaParse)** or dumping into **IndexedDB** instead of holding arrays in heap memory, given the 1M+ row Invoice Search constraints.
- Investigation matching should use **Web Workers** to prevent blocking the UI thread.
- Tree traversal naturally fits into recursive Promises or React localized state branches.
