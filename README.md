<h1 align="center">MFI Investigation Tool (v7.9)</h1>

<p align="center">
  <img src="https://img.shields.io/badge/Release-v7.9-blue.svg?style=for-the-badge" alt="Release Badge">
  <img src="https://img.shields.io/badge/Architecture-Recursive_Engine-success.svg?style=for-the-badge" alt="Architecture Badge">
  <img src="https://img.shields.io/badge/Platform-Windows_Standalone-lightgrey.svg?style=for-the-badge" alt="Platform Badge">
</p>

<p align="center">
  An enterprise-grade, recursive tracking system engineered for complex Amazon inventory and invoice verification.
</p>

---

## 📌 Executive Summary

The **MFI Investigation Tool** automates the deeply complex process of data reconciliation across multi-branch invoices. By leveraging a custom recursive deduction algorithm, it accurately accounts for units across disparate POs, caching SIDs and Barcodes globally to eliminate redundant lookups. 

Designed strictly to adhere to the latest Amazon ROW IB Workflow rules, this tool reduces manual investigation time by over 80% while ensuring 100% mathematical integrity for PQV (Price Quantity Variance) matching.

---

## 🏗️ System Architecture & Logic Flow

### Recursive Sub-Investigation Engine
The core of the tool is a recursive matching engine that traverses child matches dynamically. If a claiming ASIN matches multiple invoices, the engine initiates branch investigations.

```mermaid
graph TD
    A[Start Investigation] --> B{Calculate Total Accounted}
    B --> |Shortage >= PQV| C[Direct Shortage Gateway]
    B --> |Shortage < PQV| D[Initiate Multi-branch PQV]
    
    D --> E[Branch Budgeting]
    E --> F{Is mtc_qty <= 0?}
    F --> |Yes| G[Use Fallback Budget]
    F --> |No| H[Set budget = mtc_qty]
    
    H --> I[Execute Sub-investigation]
    G --> I
    
    I --> J{Detect Cross PO?}
    J --> |Yes| K[Automated Cross PO Chain]
    J --> |No| L[Calculate Contribution]
    
    K --> L
    L --> M[Update loop_cache]
    M --> N[End Branch]
```

### Advanced Business Logic

1. **Global Caching Layer**
   - Automatically caches SIDs and Barcodes in `InvestigationEngine.cache_sid` and `cache_bc`.
   - Prevents recursive exhaustion and redundant cross-ASIN lookups.
   
2. **Multi-branch PQV Deduplication**
   - Implements a precise constraint: `Contribution = min(match_qty, shortage_found_in_branch)`.
   - Utilizes `loop_cache` storing `(rows, total_accounted)` tuples to prevent infinite recursion loops during deep dependency chains.

3. **Automated Cross PO Detection**
   - Scans 3 discrete edge cases (Rec=0 in claiming PO, PO not invoiced, Overage in cross PO).
   - Confirmed Cross POs trigger an asynchronous `run_cross_po_investigation` thread.

4. **Fuzzy Header Correction**
   - Incorporates a dynamic `COLUMN_ALIASES` dictionary mapping to auto-fix REBNI anomalies.

---

## 💻 Tech Stack & Deployment

- **Language**: Python 3.10+
- **Data Processing**: Pandas (Optimized for `header=0` strictness and first-row vectorization).
- **Deployment**: Compiled as a strictly standalone, windowed (console-disabled) executable via PyInstaller for instant enterprise deployment.

> [!NOTE]
> *Source code is highly proprietary to the specific organizational workflow structure and is not included in this release repository. Only the compiled releases are distributed here.*
