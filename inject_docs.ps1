$path = 'c:\Users\Mukesh_Maruthi\MFI_Tool\row_ib_investigation_tool_v5_8_12-My_Fix.py'
$content = Get-Content $path -Raw

# 1. Define the DocumentationDialog class
$docsClass = @"

class DocumentationDialog(tk.Toplevel):
    \"\"\"Integrated Help System — v5.9.0 Master Guide\"\"\"
    def __init__(self, parent):
        super().__init__(parent)
        self.title("MFI Tool v5.9.0 | Master Investigation Guide")
        self.geometry("900x700")
        self.configure(bg="#0f0f1a")
        self.resizable(True, True)
        self.lift(); self.focus_force()

        tk.Label(self, text="📕  MFI TOOL MASTER INVESTIGATION GUIDE",
                 bg="#16213e", fg="#4a9eff",
                 font=("Segoe UI", 14, "bold"), height=2).pack(fill="x")

        from tkinter.scrolledtext import ScrolledText
        self.txt = ScrolledText(self, bg="#0f0f1a", fg="#e0e0e0",
                                font=("Calibri", 11), wrap="word",
                                padx=20, pady=20, relief="flat",
                                insertbackground="white")
        self.txt.pack(fill="both", expand=True)

        self.txt.tag_configure("h1", font=("Segoe UI", 14, "bold"), foreground="#f0a500")
        self.txt.tag_configure("h2", font=("Segoe UI", 12, "bold"), foreground="#4a9eff", spacing1=10)
        self.txt.tag_configure("bold", font=("Calibri", 11, "bold"), foreground="white")
        self.txt.tag_configure("tip", font=("Calibri", 11, "italic"), foreground="#f0c060")
        self.txt.tag_configure("warn", font=("Calibri", 11, "bold"), foreground="#ff4444")

        self._fill_docs()
        self.txt.config(state="disabled")

        bf = tk.Frame(self, bg="#0f0f1a"); bf.pack(pady=10)
        tk.Button(bf, text="✔  UNDERSTOOD", command=self.destroy,
                  bg="#2d6a4f", fg="white", font=("Segoe UI", 11, "bold"),
                  padx=30, pady=8, relief="flat", cursor="hand2").pack()

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+`{px}+`{py}")

    def _fill_docs(self):
        def add(txt, tag=None): self.txt.insert("end", txt, tag)
        add("CHAPTER 1: PRE-INVESTIGATION SETUP\n", "h1")
        add("Before clicking Start, you must ensure your data 'Holy Trinity' is properly loaded:\n\n")
        add("1.1 The Claims File (Your Goal)\n", "h2")
        add("The Excel sheet containing the shortages you need to explain.\n")
        add("• Required Columns: ASIN, PQV (Claimed Qty), and preferably a Barcode/SID.\n", "bold")
        add("• Auto-Fix: If your headers are non-standard, use the Header Correction Dialog to map them to the tool's expected fields.\n\n")
        add("1.2 The REBNI Data (The Source)\n", "h2")
        add("This is the Amazon Receiving data. It tells us exactly what was unpacked from a specific SID.\n")
        add("• Logic: The tool strictly uses the first row (rebni_rows[0]) for each SID+Barcode to ensure accuracy.\n\n", "tip")
        add("1.3 The Invoice Search (The Evidence)\n", "h2")
        add("This data lists all shipments associated with your ASIN across different POs and Invoices.\n")
        add("• The engine uses this to find 'Parent-Child' relationships between shipments.\n\n")
        add("CHAPTER 2: OPERATION MODES & TICKET TYPES\n", "h1")
        add("AUTO MODE (Efficiency First)\n", "h2")
        add("• The tool will automatically traverse the entire shortage chain.\n")
        add("• Direct Shortage Stop: If the first match has a shortage >= claimed PQV, the tool stops and recording the findings.\n\n", "bold")
        add("MANUAL MODE (Total Control)\n", "h2")
        add("• Recommended for high-value or complex investigations (e.g., REMASH TT).\n")
        add("• Pauses at every single match discovered for IBC/PBC validation.\n\n")
        add("CHAPTER 3: THE LIVE PREVIEW & 'CONFIRM EDITS'\n", "h1")
        add("The Live Preview Panel is your real-time log. If you notice a data glitch in the Invoice Search:\n\n")
        add("1. Identify the error in the Preview Panel.\n")
        add("2. Double-click the cell and type the correct number.\n")
        add("3. Click the [✔ CONFIRM EDITS] button.\n", "bold")
        add("4. The engine will adopt your manual correction for all future calculations!\n\n", "tip")
        add("CHAPTER 4: CROSS PO OVERAGE LIBRARY\n", "h1")
        add("• CASE 1: No PO exists for the ASIN, but it was received anyway.\n", "warn")
        add("• CASE 2: The PO exists, but the ASIN was never invoiced on it.\n", "warn")
        add("• CASE 3: Rec Qty > Inv Qty. Excess units belong to a different chain.\n", "warn")
        add("• Confirm & Investigate: Trigger a sub-investigation to find where overage units should have gone.\n\n")
        add("CHAPTER 5: ANALYZING YOUR REPORTS\n", "h1")
        add("• 7-Key Signature: Composite key ensures NO duplicate matches ever clutter your report.\n")
        add("• RED HIGHLIGHT: Indicates a confirmed shortage (Inv > Rec).\n", "warn")
        add("• GREEN HIGHLIGHT: Indicates an overage or full receipt.\n")
        add("• Math Formula: The 'Missing Qty' column is an automated Excel formula (=F-G).\n\n")
        add("CHAPTER 6: TROUBLESHOOTING & DICES\n", "h1")
        add("• Missing SID?: If a SID is missing in REBNI, check DICES and paste the number into the popup.\n", "tip")
        add("• Pending Invoices: Use the [📋 VIEW ALL PENDING INVOICES] button in Manual Mode to ensure no matched branches were ignored.\n")
        add("• STOP Button: Safely terminates the recursive engine without crashing the tool.\n", "bold")
"@

# 1. Insert DocumentationDialog class before PendingInvoicesDialog (line 740)
$anchorDialog = 'class PendingInvoicesDialog(tk.Toplevel):'
$content = $content.Replace($anchorDialog, $docsClass + "`n`n" + $anchorDialog)

# 2. Cleanup and correctly structure MFIToolApp
# Correcting branding labels from 5.8.12 to 5.9.0
$content = $content.Replace('v5.8.12  |  ROW IB', 'v5.9.0  |  ROW IB')

# 3. Insert the [📋 DOCUMENTATION] button in UI (Option A)
$btnCode = '        tk.Button(t, text="📋 DOCUMENTATION", command=self._show_docs,
                  bg="#16213e", fg="#f0c060", activebackground="#203864",
                  activeforeground="white", relief="flat", font=("Segoe UI", 9, "bold"),
                  cursor="hand2").pack(side="right", padx=10)'
$anchorLabel = '        tk.Label(t, text="v5.9.0  |  ROW IB",'
$content = $content.Replace($anchorLabel, $btnCode + "`n" + $anchorLabel)

Set-Content $path $content -NoNewline
Write-Host "Documentation integrated and branding synchronized successfully."
