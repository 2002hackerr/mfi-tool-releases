import pandas as pd
import json

df_i = pd.read_excel('Book3-INVOICE.xlsx', header=0, dtype=str)

# find the col names
mtc_inv_col = [c for c in df_i.columns if 'inv' in c.lower() and 'num' in c.lower()]
mtc_qty_col = [c for c in df_i.columns if 'qty' in c.lower() or 'quantity' in c.lower()]
mtc_po_col = [c for c in df_i.columns if 'po' in c.lower() and not 'purchase' in c.lower()]
pur_po_col = [c for c in df_i.columns if 'po' in c.lower()]
asin_col = [c for c in df_i.columns if 'asin' in c.lower()]

mtc_inv_col = mtc_inv_col[0] if mtc_inv_col else None
mtc_qty_col = mtc_qty_col[0] if mtc_qty_col else None
mtc_po_col = mtc_po_col[0] if mtc_po_col else None
asin_col = asin_col[0] if asin_col else None
pur_po_col = pur_po_col[0] if pur_po_col else None

res_i = df_i[(df_i[asin_col] == 'B012SX56OG')]

records = []
for _, r in res_i.iterrows():
    records.append({
        'invoice': r.get(mtc_inv_col, ''),
        'qty': r.get(mtc_qty_col, ''),
        'po': r.get(mtc_po_col, '') or r.get(pur_po_col, ''),
        'asin': r.get(asin_col, '')
    })

with open('inv_B012SX56OG.json', 'w') as f:
    json.dump(records, f, indent=2)
print(f'Wrote inv_B012SX56OG.json using cols: Inv={mtc_inv_col}, Qty={mtc_qty_col}, PO={mtc_po_col}, ASIN={asin_col}')
