import pandas as pd
import json

def analyze_tool_output(path):
    print(f"--- Tool Output Analysis ({path}) ---")
    df = pd.read_excel(path)
    print(f"Columns: {df.columns.tolist()}")
    # Filter rows with remarks that look like feedback
    feedback_rows = df[df['Remarks'].fillna('').str.contains('loop|skip|incorrect|mistake|need|missing|should', case=False)]
    for i, row in feedback_rows.iterrows():
        print(f"Row {i}: ASIN={row.get('ASIN')} | Mtc ASIN={row.get('Mtc ASIN')} | Mtc Inv={row.get('Mtc Inv')} | Remarks={row.get('Remarks')}")

def analyze_manual(path):
    print(f"\n--- Manual Investigation Analysis ({path}) ---")
    df = pd.read_excel(path)
    print(f"Columns: {df.columns.tolist()}")
    for i, row in df.head(30).iterrows():
        print(f"Row {i}: {row.to_dict()}")

def analyze_rebni(path):
    print(f"\n--- REBNI File Analysis ({path}) ---")
    # Load first few rows to see actual headers in the file
    df = pd.read_excel(path, header=0, nrows=5)
    print(f"ACTUAL Headers in REBNI file: {df.columns.tolist()}")
    
analyze_tool_output(r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\MFI_Investigation_20260409_143724-V2169174706.xlsx")
analyze_manual(r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\V2169174706+-+My+Investigation.xlsx")
analyze_rebni(r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\Book2-REBNI.xlsx")
