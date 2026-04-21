import pandas as pd, sys

# ── Claims ──
df_c = pd.read_excel(r'c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\Book4-Claim sheet.xlsx', header=0, dtype=str)
print("="*80)
print("CLAIMS SHEET")
print("="*80)
for i, row in df_c.iterrows():
    print(f"\n--- Claim #{i} ---")
    print(f"  Barcode:    {row.get('Barcode','')}")
    print(f"  Invoice:    {row.get('Invoice Number','')}")
    print(f"  PO:         {row.get('Header PO','')}")
    print(f"  ASIN:       {row.get('ASIN','')}")
    print(f"  Inv Qty:    {row.get('Inv Qty','')}")
    print(f"  Missing:    {row.get('Missing QTY','')}")
    print(f"  PQV Value:  {row.get('PQV Value','')}")
    print(f"  SID:        {row.get('Shipment ID','')}")

# ── Tool Output ──
df_t = pd.read_excel(r'c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\MFI_Investigation_20260408_115021- Tool out put.xlsx', header=0, dtype=str)
print("\n" + "="*80)
print("TOOL OUTPUT")
print("="*80)
print(f"Shape: {df_t.shape}")
print(f"Columns: {list(df_t.columns)}")
for i, row in df_t.iterrows():
    print(f"\nRow {i}: Barcode={row.get('Barcode','')} | Inv={row.get('Inv no','')} | SID={row.get('SID','')} | PO={row.get('PO','')} | ASIN={row.get('ASIN','')} | InvQty={row.get('Inv Qty','')} | RecQty={row.get('Rec Qty','')} | MtcQty={row.get('Mtc Qty','')} | MtcInv={row.get('Mtc Inv','')} | Remarks={row.get('Remarks','')} | CP={row.get('CP','')}")

# ── My Investigation ──
df_m = pd.read_excel(r'c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\Book1- My investigation.xlsx', header=0, dtype=str)
print("\n" + "="*80)
print("YOUR MANUAL INVESTIGATION (CORRECT OUTPUT)")
print("="*80)
print(f"Shape: {df_m.shape}")
print(f"Columns: {list(df_m.columns)}")
for i, row in df_m.iterrows():
    vals = ' | '.join([f"{c}={row[c]}" for c in df_m.columns if pd.notna(row[c]) and str(row[c]).strip()])
    print(f"Row {i}: {vals}")
