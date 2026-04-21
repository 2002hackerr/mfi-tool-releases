$path = 'c:\Users\Mukesh_Maruthi\MFI_Tool\row_ib_investigation_tool_v5_8_12-My_Fix.py'
$raw = Get-Content $path -Raw
# If it was flattened, it likely has no newlines but spaces where newlines should be,
# or it's joined by some other character.
# Let's try to restore via joining by recognized anchors if needed,
# but first, let's try a simple newline restoration if it's just a missing CRLF issue.

$verifiedLines = @(
"\"\"\"MFI Investigation Tool  v5.9.0  |  ROW IB",
"==========================================",
"ROW IB  |  Amazon",
"Developed by Mukesh",
"",
"CHANGES IN v5.8.12:",
"  ✔ [FIX] Removed duplicate CrossPODialog class (resolves v5.3.0 logic reversion)",
"  ✔ [FIX] Resolved \"Engine not running\" error in Lookup Tool (corrected app reference)",
"  ✔ [UI] Enabled Universal Red Highlighting in PreviewPanel (added 'shortage_red' tag)",
"  ✔ [PERF] Optimized Lookup Tool to use high-speed 'inv_asin_map' index",
"  ✔ [STABILITY] Hardened Cross PO \"Guard\" to prevent premature finalization",
"",
"IMPORTING LIBRARIES...",
"\"\"\"",
"import sys",
"import os",
"import math",
"import tkinter as tk",
"from tkinter import ttk, messagebox, filedialog"
# ... The rest of the recovery logic will be applied in the next surgical chunk.
)

Set-Content $path $verifiedLines
Write-Host "Top-level structure restored. Moving to full engine recovery."
