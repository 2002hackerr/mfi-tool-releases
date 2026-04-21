import pandas as pd

df_t = pd.read_excel('MFI_Investigation_20260408_115021- Tool out put.xlsx', header=0, dtype=str).fillna('')
df_m = pd.read_excel('Book1- My investigation.xlsx', header=0, dtype=str).fillna('')

def clean_df(df):
    valid = []
    for _, r in df.iterrows():
        a = str(r.get('ASIN', '')).strip()
        m = str(r.get('Mtc Qty', '')).strip()
        i = str(r.get('Inv no', '')).strip()
        # ignore blanks and headers
        if (a in ['', 'nan', 'ASIN']) and (not m) and (not i):
            continue
        valid.append({
            'ASIN': a,
            'Inv': i,
            'RecQty': str(r.get('Rec Qty', '')).strip(),
            'MtcQty': m,
            'MtcInv': str(r.get('Mtc Inv', '')).strip(),
            'Remarks': str(r.get('Remarks', '')).strip()
        })
    return valid

t_rows = clean_df(df_t)
m_rows = clean_df(df_m)

with open('comparison.txt', 'w', encoding='utf-8') as f:
    f.write(f'Tool rows: {len(t_rows)}, Manual rows: {len(m_rows)}\n\n')
    for i in range(min(50, max(len(t_rows), len(m_rows)))):
        tr = t_rows[i] if i < len(t_rows) else {}
        mr = m_rows[i] if i < len(m_rows) else {}
        
        asi_t, m_inv_t, m_qty_t, r_qty_t = tr.get('ASIN',''), tr.get('MtcInv',''), tr.get('MtcQty',''), tr.get('RecQty','')
        asi_m, m_inv_m, m_qty_m, r_qty_m = mr.get('ASIN',''), mr.get('MtcInv',''), mr.get('MtcQty',''), mr.get('RecQty','')
        
        f.write(f'Row {i}:\n')
        f.write(f'  TOOL:   ASIN={asi_t:10} MtcInv={m_inv_t:15} MtcQty={m_qty_t:5} RecQty={r_qty_t:5} Rem={tr.get("Remarks","")}\n')
        f.write(f'  MANUAL: ASIN={asi_m:10} MtcInv={m_inv_m:15} MtcQty={m_qty_m:5} RecQty={r_qty_m:5} Rem={mr.get("Remarks","")}\n')
        if tr != mr:
            f.write(f'  -> DIFFERENCE\n')
print('Done!')
