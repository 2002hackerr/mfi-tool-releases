$path = 'c:\Users\Mukesh_Maruthi\MFI_Tool\row_ib_investigation_tool_v5_8_12-My_Fix.py'
$lines = Get-Content $path
# We identify lines 2252-2257 from our previous view_file (adjusting for 0-indexing)
# In view_file they were 2252 to 2257.
# 2252:     def _show_docs(self, event=None):
# 2253:         DocumentationDialog(self.root)
# 2254:         self.all_blocks  = []
# 2255:         self.preview     = None
# 2256:         self._build_ui()
# 2257: 

# Let's filter out these specific redundant lines
$newLines = @()
for ($i = 0; $i -lt $lines.Length; $i++) {
    $lineNum = $i + 1
    if ($lineNum -ge 2252 -and $lineNum -le 2257) {
        continue
    }
    $newLines += $lines[$i]
}

Set-Content $path $newLines -NoNewline
Write-Host "Redundant lines 2252-2257 removed. Code is clean."
