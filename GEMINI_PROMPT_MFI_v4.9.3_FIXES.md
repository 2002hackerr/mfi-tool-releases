# GEMINI INSTRUCTIONS — MFI Investigation Tool v4.9.3
## READ THIS ENTIRE DOCUMENT BEFORE WRITING A SINGLE LINE OF CODE

---

## 0. ABSOLUTE AUTHORITY: CODE VERSION

You have been given a file called **`row_ib_investigation_tool_cloud_code.py`**.
This file is **MFI Investigation Tool v4.9.3** and is the **ONLY authoritative version**.

- **DO NOT** reference, use, or mix code from any other version (v4.9.2, v4.5, v4.4, etc.)
- **DO NOT** rewrite functions that are not listed in the bug list below
- **DO NOT** change the overall architecture, class structure, UI, or data loading logic
- **ONLY** make the specific targeted changes described in this document
- After every change, the code must still be runnable as a complete single Python file

---

## 1. INVESTIGATION LOGIC — CORE RULES (MASTER REFERENCE)

These are the **absolute rules** governing how each ASIN at every level (claiming or matching) must be investigated.

### RULE 1 — REBNI Lookup (Always First)
For any ASIN at any depth, always look up REBNI data first:
- `rec_qty` = `quantity_unpacked` from REBNI
- `shortage` = `max(0, inv_qty - rec_qty)` (only if REBNI row found)
- If REBNI available > 0 → write REBNI remark and STOP (no matching, no sub-rows)
- If no REBNI row at depth=0 → remark = "SID not found in REBNI — validate in DICES"
- If no REBNI row at depth>0 → remark = "SR"

### RULE 2 — DIRECT SHORT (No Matching Needed)
**Condition:** `shortage >= rem_pqv` AND `rem_pqv > 0`
(The physical shortage alone covers ALL remaining PQV budget)

**What to do:**
- Write ONE row only
- `Mtc Inv` = `"Short Received"`
- `Mtc Qty` = the shortage amount
- `Remarks` = `"X units short received directly"` where X = shortage
- **DO NOT** look up Invoice Search at all
- **DO NOT** generate any sub-rows
- **DO** still run the Cross PO check (shortage exists → Cross PO possible)
- Return `actionable = []` (no recursion into matching chains)

### RULE 3 — MATCHING INVESTIGATION (When Needed)
**Condition:** `shortage < rem_pqv`
(Physical shortage alone does NOT cover all PQV — matching chain must be traced)

This includes:
- `shortage == 0` and `rem_pqv > 0` → everything received but PQV still exists → trace matching
- `shortage > 0` but `shortage < rem_pqv` → partial physical short + remaining needs matching

**What to do:**
- Look up Invoice Search for matched invoices
- Build `sorted_m` list sorted by `mtc_qty` descending
- Determine `main_mtc_inv` and `main_mtc_qty` (see Sub-rules below)
- Generate sub-rows only for legitimate non-self branches
- Run Cross PO check (shortage > 0 → Cross PO possible)

### RULE 4 — SELF MATCHING
**Condition:** `sorted_m[0]['mtc_inv'] == inv_no` (top match IS the same invoice)

**What to do:**
- `main_mtc_inv` = `"Self Matching"`
- `main_mtc_qty` = `rec_qty` (NOT `top['mtc_qty']` — that is invoice-level total, wrong)
- `sorted_m` = `[]` — **ALWAYS CLEAR IT**
  - The remaining entries in sorted_m are how the master invoice distributed its matchings system-wide. They are NOT sub-investigation branches for THIS ASIN.
  - They must NEVER be written as sub-rows
- `actionable` = `[]` — no further recursion from self-matching
- If `shortage > 0` → add to the shortage remark (handled by `run_auto`)
- Run Cross PO check (shortage > 0 → Cross PO possible)

### RULE 5 — NON-SELF MATCHING (Matching to a Different Invoice)
**Condition:** `sorted_m[0]['mtc_inv'] != inv_no` (top match is a DIFFERENT invoice)

**What to do:**
- `main_mtc_inv` = `sorted_m[0]['mtc_inv']`
- `main_mtc_qty` = `sorted_m[0]['mtc_qty']`
- Write sub-rows for `sorted_m[1:]` (the secondary matches, each on its own line)
- `actionable` = all entries in `sorted_m` where `mtc_inv != inv_no`
- Recurse into each actionable match in `run_auto`

### RULE 6 — NO MATCHES IN INVOICE SEARCH
**Condition:** `sorted_m` is empty after deduplication

**What to do:**
- If `shortage > 0`:
  - `main_mtc_inv` = `"Short Received"`
  - `Remarks` = `"X units short received directly"`
- If `shortage == 0`:
  - Leave `main_mtc_inv` = `""`
  - No remark about shortage
- `actionable` = `[]`

### RULE 7 — CROSS PO CHECK
- Always run Cross PO detection when `shortage > 0` (any physical shortage exists)
- Cross PO rows are APPENDED after the main rows for that level
- Cross PO is detected via `detect_cross_po(sid, po, asin)`

### RULE 8 — REC QTY = MTC QTY (Invariant)
For any ASIN at any level, `Mtc Qty` shown in the output must reflect what was actually matched for THAT specific ASIN, which should always equal `rec_qty` when self-matching. The invoice-level total matched qty is NEVER the correct value for Mtc Qty in the output row.

---

## 2. BUGS TO FIX — EXACT LOCATION AND EXACT FIX

### BUG 1 (CRITICAL — Causes 80+ garbage sub-rows per ASIN)
**File:** `row_ib_investigation_tool_cloud_code.py`
**Function:** `_build_level_logic` (class `InvestigationEngine`)
**Lines around:** the `if sorted_m:` block

**Current broken code:**
```python
if sorted_m:
    top = sorted_m[0]
    if top['mtc_inv'] == clean(inv_no):
        main_mtc_inv = "Self Matching"; main_mtc_qty = fmt_qty(top['mtc_qty'])
        sorted_m = sorted_m[1:] if len(sorted_m) > 1 else []   # ← BUG: keeps all other entries
    else:
        main_mtc_inv = top['mtc_inv']; main_mtc_qty = fmt_qty(top['mtc_qty'])
else:
    if not remarks and shortage > 0:
        main_mtc_inv = "Short Received"
        remarks = f"Found {int(shortage)} units short as loop started from {int(rem_pqv)} matched qty, no remaining pqv"
```

**What is wrong:**
When `top['mtc_inv'] == inv_no` (Self Matching), the code sets `sorted_m = sorted_m[1:]`.
This keeps ALL remaining invoice matches (80+ entries in real data) which then all get written as sub-rows.
These remaining entries are how invoice 2503640029 was matched across the system — they are NOT investigation branches for this specific ASIN.

**Fixed code (replace the entire block above with this):**
```python
if sorted_m:
    top = sorted_m[0]
    if top['mtc_inv'] == clean(inv_no):
        # Self Matching: invoice matched itself.
        # Mtc Qty must be rec_qty (what THIS ASIN received), NOT top['mtc_qty'] (invoice-level total).
        # sorted_m MUST be cleared — remaining entries are invoice-system matches, NOT ASIN branches.
        main_mtc_inv = "Self Matching"
        main_mtc_qty = fmt_qty(rec_qty)   # ← FIXED: use rec_qty, not top['mtc_qty']
        sorted_m = []                       # ← FIXED: always clear, never keep sorted_m[1:]
    else:
        # Non-self match: top match is a different invoice — legitimate branch
        main_mtc_inv = top['mtc_inv']
        main_mtc_qty = fmt_qty(top['mtc_qty'])
        # sorted_m keeps remaining items for sub-rows (sorted_m[1:] written as sub-rows below)
else:
    if not remarks and shortage > 0:
        main_mtc_inv = "Short Received"
        main_mtc_qty = fmt_qty(shortage)
        remarks = f"{int(shortage)} units short received directly"
```

---

### BUG 2 (CRITICAL — Direct shortage ASINs incorrectly trigger matching)
**Function:** `_build_level_logic`
**Problem:** When `shortage >= rem_pqv` (direct short covers all PQV), the code still performs the Invoice Search lookup and generates matching rows. This is wrong — when it's a direct shortage, no matching investigation is needed.

**Fix:** Add a direct-short check BEFORE the Invoice Search lookup. Insert this block immediately AFTER the `shortage = ...` calculation and BEFORE the `raw = self._inv_lookup(...)` line:

```python
shortage = max(0.0, safe_num(inv_qty) - rec_qty) if rebni_rows else 0.0

# ─── DIRECT SHORT CHECK ───────────────────────────────────────────────────
# If physical shortage >= remaining PQV budget, this ASIN is directly short.
# No matching investigation needed. Just record the shortage and check Cross PO.
if shortage >= rem_pqv > 0 and not remarks:
    main_row = self._make_row(barcode, inv_no, sid, po, asin,
                               inv_qty, rec_qty, fmt_qty(shortage),
                               "Short Received",
                               f"{int(shortage)} units short received directly",
                               rec_date, depth)
    cross_rows = self._build_cross_po_rows(sid_frag, clean(po), clean(asin), depth)
    return [main_row] + cross_rows, [], rec_qty, shortage, 0.0, ex_adj
# ─────────────────────────────────────────────────────────────────────────
```

---

### BUG 3 (Sub-row loop writes all sorted_m items when Self Matching)
**Function:** `_build_level_logic`
**Current code:**
```python
rows = [self._make_row(...)]
for i, m in enumerate(sorted_m):
    if i == 0 and main_mtc_inv not in ("Self Matching","Short Received"): continue
    rows.append(self._make_row("","","","","","","",fmt_qty(m['mtc_qty']), m['mtc_inv'],"","", depth, 'subrow'))
```

**Problem:** When `main_mtc_inv == "Self Matching"`, the condition `i == 0 and main_mtc_inv not in (...)` evaluates to `False` for i=0, so it does NOT skip i=0. It writes ALL items from sorted_m (which after Bug 1 fix will be empty, so this bug is resolved by Bug 1 fix). However, for non-self-matching cases, the loop also incorrectly skips index 0 when it shouldn't.

**Fixed loop (replace with this):**
```python
rows = [self._make_row(barcode, inv_no, sid, po, asin,
                        inv_qty, rec_qty, main_mtc_qty, main_mtc_inv,
                        remarks, rec_date, depth)]

# Sub-rows: for non-self matching cases, sorted_m[0] is already shown as main row.
# Write sorted_m[1:] as sub-rows. For Self Matching, sorted_m is empty (no sub-rows).
sub_start = 1 if (sorted_m and main_mtc_inv not in ("Self Matching", "Short Received")) else 0
for m in sorted_m[sub_start:]:
    rows.append(self._make_row("", "", "", "", "", "", "",
                                fmt_qty(m['mtc_qty']), m['mtc_inv'],
                                "", "", depth, 'subrow'))
```

---

### BUG 4 (Remark wording — minor)
**Function:** `_build_level_logic`, in the `else` block (no sorted_m, no REBNI)
**Current:** `"Found {shortage} units short as loop started from {rem_pqv} matched qty, no remaining pqv"`
**Fix:** Change to `"{shortage} units short received directly"`

This is already handled in Bug 1 and Bug 2 fixes above.

---

## 3. FUNCTION `_build_level_logic` — COMPLETE REWRITTEN LOGIC SECTION

Replace the entire body of `_build_level_logic` from the line `shortage = max(...)` to the line `return rows, actionable, rec_qty, shortage, new_rem, ex_adj` with this:

```python
        shortage = max(0.0, safe_num(inv_qty) - rec_qty) if rebni_rows else 0.0

        # ── DIRECT SHORT: shortage covers all remaining PQV — no matching needed ──
        if shortage >= rem_pqv > 0 and not remarks:
            main_row = self._make_row(barcode, inv_no, sid, po, asin,
                                       inv_qty, rec_qty, fmt_qty(shortage),
                                       "Short Received",
                                       f"{int(shortage)} units short received directly",
                                       rec_date, depth)
            cross_rows = self._build_cross_po_rows(sid_frag, clean(po), clean(asin), depth)
            return [main_row] + cross_rows, [], rec_qty, shortage, 0.0, ex_adj

        # ── MATCHING INVESTIGATION ─────────────────────────────────────────────
        raw  = self._inv_lookup(sid_frag, clean(po), clean(asin), clean(inv_no))
        seen = set()
        unique = []
        for m in raw:
            combo = (m['mtc_inv'], m['mtc_po'], m['mtc_asin'])
            if combo not in seen:
                seen.add(combo)
                unique.append(m)
        sorted_m = sorted(unique, key=lambda x: safe_num(x['mtc_qty']), reverse=True)

        main_mtc_inv = ""
        main_mtc_qty = ""

        if sorted_m:
            top = sorted_m[0]
            if top['mtc_inv'] == clean(inv_no):
                # Self Matching: invoice matched itself.
                # Mtc Qty = rec_qty for this ASIN (NOT invoice-level top['mtc_qty']).
                # sorted_m cleared: remaining entries are invoice-system distributions, NOT ASIN branches.
                main_mtc_inv = "Self Matching"
                main_mtc_qty = fmt_qty(rec_qty)
                sorted_m     = []
            else:
                # Non-self match: first entry shown as main row, rest as sub-rows.
                main_mtc_inv = top['mtc_inv']
                main_mtc_qty = fmt_qty(top['mtc_qty'])
                # sorted_m stays as-is; sub-rows written from [1:] below
        else:
            # No Invoice Search matches at all
            if not remarks and shortage > 0:
                main_mtc_inv = "Short Received"
                main_mtc_qty = fmt_qty(shortage)
                remarks = f"{int(shortage)} units short received directly"

        # Build main row
        rows = [self._make_row(barcode, inv_no, sid, po, asin,
                                inv_qty, rec_qty, main_mtc_qty, main_mtc_inv,
                                remarks, rec_date, depth)]

        # Sub-rows: only for non-self-matching cases (sorted_m[1:])
        # For Self Matching, sorted_m is [] so no sub-rows written.
        sub_start = 1 if (sorted_m and main_mtc_inv not in ("Self Matching", "Short Received")) else 0
        for m in sorted_m[sub_start:]:
            rows.append(self._make_row("", "", "", "", "", "", "",
                                        fmt_qty(m['mtc_qty']), m['mtc_inv'],
                                        "", "", depth, 'subrow'))

        # Actionable = non-self matches to recurse into
        actionable = [m for m in sorted_m if m['mtc_inv'] != clean(inv_no)]

        new_rem = max(0.0, rem_pqv - min(rem_pqv, shortage))

        # Cross PO check (runs when shortage > 0 or always at claiming level)
        rows.extend(self._build_cross_po_rows(sid_frag, clean(po), clean(asin), depth))

        return rows, actionable, rec_qty, shortage, new_rem, ex_adj
```

---

## 4. INVESTIGATION DECISION FLOW (Visual Reference)

```
For each ASIN (claiming or matching):
│
├─► Look up REBNI
│   ├─ Not found (depth=0) → remark="SID not found", STOP
│   ├─ Not found (depth>0) → remark="SR", STOP
│   └─ Found → get rec_qty, shortage
│
├─► REBNI Available > 0? → remark="REBNI Available...", STOP
│
├─► shortage >= rem_pqv (direct short)?
│   └─ YES → ONE ROW: "X units short received directly", Mtc Inv="Short Received"
│            Check Cross PO if shortage > 0. STOP.
│
└─► shortage < rem_pqv (need matching):
    │
    ├─► Look up Invoice Search
    │
    ├─► Top match = self invoice?
    │   └─ YES → Self Matching, Mtc Qty = rec_qty, sorted_m=[], no sub-rows, no recursion
    │            Check Cross PO if shortage > 0
    │
    ├─► Top match = different invoice?
    │   └─ YES → Mtc Inv = that invoice, sub-rows for rest, recurse into actionable
    │            Check Cross PO if shortage > 0
    │
    └─► No matches at all?
        └─ shortage > 0 → "X units short received directly"
           shortage = 0 → leave blank
```

---

## 5. WHAT YOU MUST NOT CHANGE

- Do NOT change `run_cross_po_investigation`
- Do NOT change `detect_cross_po`
- Do NOT change `build_one_level`
- Do NOT change `run_auto` (except if fixing rem_budget check for direct-short compatibility)
- Do NOT change any dialog classes (ManualLevelDialog, CrossPODialog, etc.)
- Do NOT change `write_excel`
- Do NOT change data loading functions (`load_rebni`, `load_invoice_search`, etc.)
- Do NOT change the UI code

---

## 6. VALIDATION — HOW TO CHECK YOUR FIX IS CORRECT

After applying the fix, the output Excel for the provided test data should show:

1. **ASINs where `Inv - Rec = PQV`**: ONE row per ASIN, Mtc Inv = "Short Received", Remarks = "X units short received directly". ZERO sub-rows.

2. **ASINs where Self Matching (Rec = Inv or partial shortage, top match = self)**: ONE row per ASIN, Mtc Inv = "Self Matching", **Mtc Qty = rec_qty** (NOT the 1085 invoice total). ZERO sub-rows.

3. **Total output rows**: Should be approximately 2-3 rows per ASIN (header + main row + blank), NOT 80-100 rows per ASIN.

4. **No rows** with only Mtc Qty and Mtc Inv filled and everything else None.

---

## 7. VERSION INSTRUCTION

After completing all changes, update the version string in the file header from:
```
MFI Investigation Tool  v4.9.3
```
to:
```
MFI Investigation Tool  v4.9.4
```
And add to the CHANGES section:
```
CHANGES IN v4.9.4 (over v4.9.3):
  ✔ Direct shortage ASINs (Inv-Rec >= PQV) now produce ONE row with "Short Received" — no matching performed
  ✔ Self Matching Mtc Qty now correctly shows rec_qty (not invoice-level total)
  ✔ Self Matching no longer generates garbage sub-rows from sorted_m[1:]
  ✔ Sub-row loop corrected: sorted_m[1:] only for non-self-matching branches
  ✔ Cross PO check retained for all cases where shortage > 0
```

---

## 8. SUMMARY OF CODE CHANGES (Quick Reference for Gemini)

| # | Function | Change |
|---|----------|--------|
| 1 | `_build_level_logic` | Add direct-short early-return before Invoice Search lookup |
| 2 | `_build_level_logic` | Self Matching: `main_mtc_qty = fmt_qty(rec_qty)` (was `top['mtc_qty']`) |
| 3 | `_build_level_logic` | Self Matching: `sorted_m = []` (was `sorted_m[1:]`) |
| 4 | `_build_level_logic` | Sub-row loop: write `sorted_m[sub_start:]` where `sub_start=1` for non-self |
| 5 | `_build_level_logic` | No-match case: remark = "X units short received directly" |

**Total functions modified: 1** (`_build_level_logic` only)
**Total lines changed: ~15**
**Architecture: unchanged**
