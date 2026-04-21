import pandas as pd
import json

def dump_excel(path, name):
    print(f"--- {name} ({path}) ---")
    try:
        df = pd.read_excel(path, header=0)
        # Convert to records and print first few rows or key comparisons
        cols = df.columns.tolist()
        print(f"Columns: {cols}")
        data = df.to_dict(orient='records')
        for i, row in enumerate(data[:50]): # Read up to 50 rows
            print(f"Row {i}: {row}")
    except Exception as e:
        print(f"Error reading {path}: {e}")

# Read Tool Output (with feedback)
dump_excel(r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\MFI_Investigation_20260409_143724-V2169174706.xlsx", "Tool Output with Feedback")

# Read Manual Investigation
dump_excel(r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\V2169174706+-+My+Investigation.xlsx", "Manual Investigation")

# Read Claiming Sheet
dump_excel(r"c:\Users\Mukesh_Maruthi\MFI_Tool\debug_files\mazharkn_COCOBLU_2503702299-V2169174706.xlsx", "Claiming Sheet")
