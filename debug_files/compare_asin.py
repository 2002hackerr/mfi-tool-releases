import pandas as pd

df_t = pd.read_excel('MFI_Investigation_20260408_115021- Tool out put.xlsx', header=0, dtype=str).fillna('')
df_m = pd.read_excel('Book1- My investigation.xlsx', header=0, dtype=str).fillna('')

def group_by_asin(df):
    groups = {}
    current_asin = None
    curr_group = []
    
    for _, r in df.iterrows():
        header_asin = str(r.get('ASIN', '')).strip()
        inv_no_hdr = str(r.get('Inv no', '')).strip()
        
        # skip completely blank rows
        if not header_asin and not inv_no_hdr and not str(r.get('Mtc Qty','')).strip():
            continue
            
        # Is this a main row?
        if header_asin and header_asin != 'nan' and header_asin != 'ASIN':
            if current_asin:
                groups.setdefault(current_asin, []).append(curr_group)
            current_asin = header_asin
            curr_group = []
            
        if current_asin:
            curr_group.append({
                'ASIN': header_asin,
                'Inv': str(r.get('Inv no', '')).strip(),
                'SubInv': str(r.get('Mtc Inv', '')).strip(),
                'Rec': str(r.get('Rec Qty', '')).strip(),
                'Mtc': str(r.get('Mtc Qty', '')).strip(),
                'Rem': str(r.get('Remarks', '')).strip()
            })
            
    if current_asin:
        groups.setdefault(current_asin, []).append(curr_group)
    return groups

t_g = group_by_asin(df_t)
m_g = group_by_asin(df_m)

with open('asin_comp.txt', 'w', encoding='utf-8') as f:
    for asin in m_g.keys():
        f.write(f'--- ASIN {asin} ---\n')
        f.write('TOOL:\n')
        blocks_t = t_g.get(asin, [])
        for bi, blk in enumerate(blocks_t):
            for ri, row in enumerate(blk):
                f.write(f'  T-{bi}-{ri}: SubInv={row["SubInv"]:15} Rec={row["Rec"]:5} Mtc={row["Mtc"]:5} Rem={row["Rem"]}\n')
                
        f.write('MANUAL:\n')
        blocks_m = m_g.get(asin, [])
        for bi, blk in enumerate(blocks_m):
            for ri, row in enumerate(blk):
                f.write(f'  M-{bi}-{ri}: SubInv={row["SubInv"]:15} Rec={row["Rec"]:5} Mtc={row["Mtc"]:5} Rem={row["Rem"]}\n')
        f.write('\n')

print('Done!')
