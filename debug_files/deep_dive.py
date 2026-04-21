import pandas as pd
import json

def clean(s):
    if pd.isna(s): return ''
    return str(s).strip()

def safe_num(v):
    try: return float(str(v).replace(',', ''))
    except: return 0.0

print('Deep Analysis of B012SX56OG')

df_rebni = pd.read_excel('Book2-REBNI.xlsx', header=0, dtype=str)
df_inv = pd.read_excel('Book3-INVOICE.xlsx', header=0, dtype=str)

claim_asin = 'B012SX56OG'
claim_sid_frag = '5694387014973'
claim_po = '2WTJXTFT'
claim_inv_no = '2503702230'

# REBNI Data
rebni_rows = df_rebni[(df_rebni['asin'] == claim_asin) & (df_rebni['shipment_id'].astype(str).str.contains(claim_sid_frag))]
print('\nREBNI ROWS:')
for _, r in rebni_rows.iterrows():
    print(f"  SID={r['shipment_id']} PO={r['po']} ASIN={r['asin']} Rec={r['quantity_unpacked']}")

# Invoice Search Data
# Find columns
sid_col = [c for c in df_inv.columns if 'shipment' in c.lower() and 'id' in c.lower()][0]
po_col = [c for c in df_inv.columns if 'purchase' in c.lower() and 'order' in c.lower()][0]
asin_col = [c for c in df_inv.columns if 'asin' in c.lower() and 'matched' not in c.lower().split()][0] # Adjust if needed
# Actually, the tool uses (sid, po, asin) to look up in the index.
# In Book3-INVOICE, it seems the columns are 'purchase_order', 'asin', 'shipment_id' etc based on previous cat output.

# Let's just grep the entire Invoice file for the (SID, PO, ASIN) combo
# But look at matched_invoice_number
print('\nINVOICE SEARCH ROWS (Matched to this Claim):')
# Filter by the claim PO and SID and ASIN
inv_matches = df_inv[
    (df_inv['asin'] == claim_asin) & 
    (df_inv['purchase_order'] == claim_po) & 
    (df_inv['shipment_id'].astype(str).str.contains(claim_sid_frag))
]

for _, r in inv_matches.iterrows():
    print(f"  MtcInv={r['matched_invoice_number']} MtcQty={r['matched_invoice_quantity']} MtcPO={r['matched_po']} MtcASIN={r['matched_asin']}")

print('\nALL INVOICE SEARCH ROWS FOR THIS ASIN (regardless of shipment):')
asin_matches = df_inv[df_inv['asin'] == claim_asin]
print(f"Found {len(asin_matches)} total matches for {claim_asin}")
# Group by matched invoice and sum quantity
agg = asin_matches.groupby('matched_invoice_number')['matched_invoice_quantity'].apply(lambda x: sum(safe_num(i) for i in x)).sort_values(ascending=False)
print('\nTOP MATCHED INVOICES FOR THIS ASIN:')
print(agg.head(20))
