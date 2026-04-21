@echo off
echo Installing dependencies...
pip install --quiet pandas openpyxl pyinstaller
echo Building EXE for MFI Investigation Tool v4.2...
pyinstaller --onefile --windowed --name "MFI_Investigation_Tool_v4.2" mfi_tool.py
echo Done! Find MFI_Investigation_Tool_v4.2.exe in the 'dist' folder.
pause
