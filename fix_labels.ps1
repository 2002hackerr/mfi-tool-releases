# 1. Fix Personal Edition Label
$p_path = 'c:\Users\Mukesh_Maruthi\MFI_Tool\MFI_Tool_Personal.py'
$p_content = Get-Content $p_path -Raw
$p_content = $p_content.Replace('v5.8.12  |  ROW IB', 'v5.9.0  |  ROW IB')
Set-Content $p_path $p_content -NoNewline

# 2. Fix Secured Edition Redundancy
$s_path = 'c:\Users\Mukesh_Maruthi\MFI_Tool\MFI_Tool_Secured.py'
$s_content = Get-Content $s_path -Raw
# First fix the header redundancy
$s_content = $s_content.Replace('(Secured Edition) (Secured Edition)', '(Secured Edition)')
# Ensure UI label is perfect
$s_content = $s_content.Replace('v5.8.12  |  ROW IB', 'v5.9.0  |  ROW IB  [SECURED]')
Set-Content $s_path $s_content -NoNewline

Write-Host "Label synchronization and branding polish complete."
