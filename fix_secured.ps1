$path = 'c:\Users\Mukesh_Maruthi\MFI_Tool\MFI_Tool_Secured.py'
$content = Get-Content $path -Raw

# 1. Update Header
$content = $content.Replace('MFI Investigation Tool  v5.9.0  |  ROW IB', 'MFI Investigation Tool  v5.9.0  |  ROW IB (Secured Edition)')

# 2. Update Imports
$content = $content.Replace('import os, re, threading', 'import os, re, threading, urllib.request, sys')

# 3. Add Activation Logic
$actCode = @"

# ==============================================================================
#  SECURED PRODUCTION VERSION
# ==============================================================================
ACTIVATION_URL = "https://gist.githubusercontent.com/2002hackerr/3f76afc8a819c6879e06676a36173999/raw/activation.txt"

def check_activation():
    """Verify tool status remotely. Returns True only if 'ENABLED' is found."""
    try:
        req = urllib.request.Request(ACTIVATION_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            content = response.read().decode('utf-8').strip().upper()
            return content == "ENABLED"
    except Exception as e:
        return False

"@

$content = $content.Replace('from datetime import datetime', "from datetime import datetime$actCode")

# 4. Wrap __init__
$initOld = '    def __init__(self):'
$initNew = @"
    def __init__(self):
        # SECURITY CHECK: Verify remote activation before starting GUI
        if not check_activation():
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("ACCESS DENIED", 
                                 "ADMIN AUTHORIZATION REQUIRED\n\n" +
                                 "This version of the MFI Tool is currently inactive.\n" +
                                 "Please contact Mukesh (the administrator) for permission.")
            sys.exit()
"@

$content = $content.Replace($initOld, $initNew)

# 5. Fix remaining labels
$content = $content.Replace('v5.8.12  |  ROW IB', 'v5.9.0  |  ROW IB  [SECURED]')

Set-Content $path $content -NoNewline
Write-Host "Secured Edition synchronization complete."
