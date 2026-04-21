import sys
import os
import pandas as pd

# Path to the new version
sys.path.append(r"c:\Users\Mukesh_Maruthi\MFI_Tool")
from row_ib_investigation_tool_v5_6_1 import InvestigationEngine, load_claims, build_rebni_index, load_rebni, build_invoice_index, load_invoice_search, write_excel, extract_sid, safe_num, clean

def generate_example():
    # Load data (restricted to first few for speed)
    print("Loading data...")
    claims_path = r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\mazharkn_COCOBLU_2503702299-V2169174706.xlsx"
    rebni_path = r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\Book2-REBNI.xlsx"
    inv_path = r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\Book3-INVOICE.xlsx"
    
    df_c = load_claims(claims_path).head(5) # Just 5 rows for example
    rp, rs, rfb = build_rebni_index(load_rebni(rebni_path))
    ip, ifb = build_invoice_index(load_invoice_search(inv_path))
    
    engine = InvestigationEngine(rp, rs, rfb, ip, ifb, lambda x,y,z: None)
    
    all_blocks = []
    print("Running auto-investigation for example...")
    for i, r in df_c.iterrows():
        rows, _ = engine.run_auto(
            clean(r.get('Barcode', '')),
            clean(r.get('Invoice', '')),
            extract_sid(clean(r.get('SID', ''))),
            clean(r.get('PO', '')),
            clean(r.get('ASIN', '')),
            safe_num(r.get('InvQty', 0)),
            safe_num(r.get('PQV', 0))
        )
        all_blocks.append(rows)
    
    out_path = r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\Example_Output_v5_6_1.xlsx"
    write_excel(all_blocks, out_path)
    print(f"Example generated: {out_path}")

if __name__ == "__main__":
    generate_example()
