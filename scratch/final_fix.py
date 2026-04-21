
import os

target_file = r"c:\Users\Mukesh_Maruthi\MFI_Tool\row_ib_investigation_tool_v5_8_12-My_Fix.py"

with open(target_file, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Part 1: Lines 1-848 (Prefix)
prefix = lines[:848]

# Part 2: Finalized CorrespondenceDialog
new_dialog = [
    "class CorrespondenceDialog(tk.Toplevel):\n",
    "    def __init__(self, parent, all_rows):\n",
    "        super().__init__(parent)\n",
    "        self.all_rows = all_rows\n",
    "        self.title(\"Scenario Selection — Get Correspondence — v5.9.0\")\n",
    "        self.geometry(\"950x800\")\n",
    "        self.configure(bg=\"#0f0f1a\")\n",
    "        self.transient(parent)\n",
    "        self.grab_set()\n",
    "        \n",
    "        self.scenarios = [\n",
    "            \"Claiming Short\",\n",
    "            \"Matching\",\n",
    "            \"Claiming & Matching Short\",\n",
    "            \"REBNI\",\n",
    "            \"Matching to Invalid Invoice [Dev]\",\n",
    "            \"Dummy Invoice [Dev]\",\n",
    "            \"Difference in IBC versus PBC [Dev]\"\n",
    "        ]\n",
    "        \n",
    "        self.v_code_var = tk.StringVar(); self.fc_id_var = tk.StringVar()\n",
    "        self.fc_id_var.set(\"[Enter FC ID]\")\n",
    "        self.v_code_var.set(\"[Enter Vendor Code]\")\n",
    "        self._build_ui()\n",
    "        \n",
    "    def _build_ui(self):\n",
    "        f = tk.Frame(self, bg=\"#0f0f1a\", padx=20, pady=20)\n",
    "        f.pack(fill=\"both\", expand=True)\n",
    "        \n",
    "        top_h = tk.Frame(f, bg=\"#0f0f1a\")\n",
    "        top_h.pack(fill=\"x\", pady=(0,15))\n",
    "        tk.Label(top_h, text=\"Investigation Correspondence Generator\",\n",
    "                 fg=\"#4a9eff\", bg=\"#0f0f1a\",\n",
    "                 font=(\"Segoe UI\", 14, \"bold\")).pack(side=\"left\")\n",
    "        \n",
    "        sel_f = tk.LabelFrame(f, text=\" Step 1: Select Scenario \", fg=\"#e0e0e0\", bg=\"#0f0f1a\", font=(\"Segoe UI\", 9, \"bold\"))\n",
    "        sel_f.pack(fill=\"x\", pady=10, padx=2)\n",
    "        self.scenario_var = tk.StringVar()\n",
    "        self.scenario_combobox = ttk.Combobox(sel_f, textvariable=self.scenario_var, \n",
    "                                             values=self.scenarios, state=\"readonly\", \n",
    "                                             font=(\"Segoe UI\", 11))\n",
    "        self.scenario_combobox.pack(fill=\"x\", padx=10, pady=10)\n",
    "        self.scenario_combobox.bind(\"<<ComboboxSelected>>\", self.generate_text)\n",
    "        \n",
    "        inp_f = tk.LabelFrame(f, text=\" Step 2: Manual Details (Optional) \", fg=\"#e0e0e0\", bg=\"#0f0f1a\", font=(\"Segoe UI\", 9, \"bold\"))\n",
    "        inp_f.pack(fill=\"x\", pady=10, padx=2)\n",
    "        grid_f = tk.Frame(inp_f, bg=\"#0f0f1a\")\n",
    "        grid_f.pack(fill=\"x\", padx=10, pady=5)\n",
    "        tk.Label(grid_f, text=\"Vendor Code:\", bg=\"#0f0f1a\", fg=\"#cccccc\").grid(row=0, column=0, sticky=\"w\")\n",
    "        tk.Entry(grid_f, textvariable=self.v_code_var, width=20).grid(row=0, column=1, padx=5, pady=2)\n",
    "        tk.Label(grid_f, text=\"FC ID:\", bg=\"#0f0f1a\", fg=\"#cccccc\").grid(row=0, column=2, sticky=\"w\", padx=(20,0))\n",
    "        tk.Entry(grid_f, textvariable=self.fc_id_var, width=20).grid(row=0, column=3, padx=5, pady=2)\n",
    "        \n",
    "        tk.Label(f, text=\" Step 3: Review and Copy Generated Text:\",\n",
    "                 fg=\"#cccccc\", bg=\"#0f0f1a\",\n",
    "                 font=(\"Segoe UI\", 10)).pack(anchor=\"w\", pady=(10, 5))\n",
    "        self.text_area = tk.Text(f, height=25, font=(\"Consolas\", 10), \n",
    "                                 bg=\"#1e1e3a\", fg=\"white\", padx=10, pady=10,\n",
    "                                 insertbackground=\"white\", wrap=\"word\")\n",
    "        self.text_area.pack(fill=\"both\", expand=True)\n",
    "        \n",
    "        btn_f = tk.Frame(f, bg=\"#0f0f1a\")\n",
    "        btn_f.pack(fill=\"x\", pady=(20, 0))\n",
    "        tk.Button(btn_f, text=\"\ud83d\udccb Copy Ticket Correspondence\", \n",
    "                  command=self.copy_to_clip,\n",
    "                  bg=\"#2d6a4f\", fg=\"white\", font=(\"Segoe UI\", 11, \"bold\"),\n",
    "                  relief=\"flat\", padx=20, pady=10, cursor=\"hand2\").pack(side=\"left\")\n",
    "        tk.Button(btn_f, text=\"Close\", command=self.destroy,\n",
    "                  bg=\"#4a2020\", fg=\"white\", font=(\"Segoe UI\", 10),\n",
    "                  relief=\"flat\", padx=20, pady=10, cursor=\"hand2\").pack(side=\"right\")\n",
    "\n",
    "    def copy_to_clip(self):\n",
    "        text = self.text_area.get(\"1.0\", tk.END).strip()\n",
    "        if not text: return\n",
    "        self.clipboard_clear()\n",
    "        self.clipboard_append(text)\n",
    "        messagebox.showinfo(\"Copied\", \"Correspondence copied to clipboard!\")\n",
    "\n",
    "    def _get_asin_table(self, rows, cols=9):\n",
    "        header = \"Barcode\\tInv no\\tSID\\tPO\\tASIN\\tInv Qty\\tRec Qty\\tMissing QTY\\tCost Price\"\n",
    "        lines = [header]\n",
    "        for r in rows:\n",
    "            iq = safe_num(r.get('inv_qty',0))\n",
    "            rq = safe_num(r.get('rec_qty',0))\n",
    "            missing = iq - rq\n",
    "            if missing <= 0: continue\n",
    "            cp = safe_num(r.get('cp_status',0))\n",
    "            row_data = [\n",
    "                str(r.get('barcode','')).strip(),\n",
    "                str(r.get('invoice','')).strip(),\n",
    "                str(r.get('sid','')).strip(),\n",
    "                str(r.get('po','')).strip(),\n",
    "                str(r.get('asin','')).strip(),\n",
    "                str(int(iq) if iq==int(iq) else iq),\n",
    "                str(int(rq) if rq==int(rq) else rq),\n",
    "                str(int(missing) if missing==int(missing) else missing),\n",
    "                f\"{cp:.2f}\"\n",
    "            ]\n",
    "            lines.append(\"\\t\".join(row_data))\n",
    "        return \"\\n\".join(lines)\n",
    "\n",
    "    def generate_text(self, event=None):\n",
    "        scenario = self.scenario_var.get()\n",
    "        if not scenario: return\n",
    "        claiming_row = None; all_shortages = []; matched_invoices = set(); matching_details = []\n",
    "        for r in self.all_rows:\n",
    "            bc = str(r.get('barcode','')).strip()\n",
    "            if bc and not bc.startswith('['):\n",
    "                iq, rq = safe_num(r.get('inv_qty',0)), safe_num(r.get('rec_qty',0))\n",
    "                if iq > rq: all_shortages.append(r)\n",
    "                mtc_inv = str(r.get('mtc_inv','')).strip()\n",
    "                if mtc_inv and mtc_inv.lower() not in ('','none','short received'):\n",
    "                    matched_invoices.add(mtc_inv)\n",
    "                    matching_details.append(f\"{r.get('po','')} \\t {mtc_inv} \\t {bc}\")\n",
    "                if r.get('depth',0) == 0 and not claiming_row: claiming_row = r\n",
    "        \n",
    "        c_sid = claiming_row.get('sid','') if claiming_row else \"[Enter SID]\"\n",
    "        c_po  = claiming_row.get('po','')  if claiming_row else \"[Enter PO]\"\n",
    "        t_inv = sum(safe_num(r.get('inv_qty',0)) for r in self.all_rows if r.get('depth',0)==0)\n",
    "        t_rec = sum(safe_num(r.get('rec_qty',0)) for r in self.all_rows if r.get('depth',0)==0)\n",
    "        v_code = self.v_code_var.get(); f_id = self.fc_id_var.get()\n",
    "        rules_text = \"\"\"i. FC to perform physical investigation only for missing units and share complete findings in one correspondence. Virtual research is not to be done by FC.  \\nii. If units are not found by FC and PQV > 50k, then FC to flip the TT to FC SLP for further research.  \\niii. MFI TT should be flipped to ROW IB only after complete FC and FC SLP investigations are performed.  \\niv. If SLP investigation is pending or more time is needed for physical search, the MFI TT should remain in the FC queue.  \\nv. FC should avoid repeated or partial updates in correspondence regarding MFI TT missing units.\"\"\"\n",
    "        \n",
    "        if scenario == \"Claiming Short\":\n",
    "            table = self._get_asin_table(all_shortages)\n",
    "            text = f\"\"\"Hello FC Team, \\n\\nNote: We have already performed the all the virtual research/checks such as cross receiving, overage, REBNI, adjustments etc.  and need FC support for physical search on floor to locate missing units as per revised SOP and update at the earliest. \\n\\nWe are able to see that unit\u2019s shortage received in claiming shipment.\\n\\nClaiming SID#{c_sid} | Total Billed Qty - {int(t_inv)} | Received Qty - {int(t_rec)} \\n\\nPlease locate the following ASINs that are missing from PO#{c_po}\\n{table}\\n\\nSID: {c_sid}\\nPO: {c_po}\\nVendor Code: {v_code}\\nFC ID: {f_id}\\n\\nNOTE : FC Team we are able to units are short at SID Level and No overages are Found.\\n\\n{rules_text}\\n\\nRegards, \\nMUKESH | pathlmuk\\nROW IB\"\"\"\n",
    "        elif scenario == \"Matching\":\n",
    "            m_inv_str = \", \".join(matched_invoices)\n",
    "            table = self._get_asin_table(all_shortages)\n",
    "            m_details = \"\\n\".join(matching_details) if matching_details else \"[No matching siblings found]\"\n",
    "            text = f\"\"\"Hello FC Team, \\n\\nNote: We have already performed the all the virtual research/checks such as cross receiving, overage, REBNI, adjustments etc.  and need FC support for physical search on floor to locate missing units as per revised SOP and update at the earliest. \\n\\nWe are able to see that units are received completely in claiming Shipment ID#{c_sid} and received units matched with different invoices ({m_inv_str}) where we found shortage units received in matched invoices. \\n\\nSID#{c_sid} | Total Billed Qty - {int(t_inv)} | Received Qty - {int(t_rec)} \\n\\nPlease locate the following ASINs that are missing from below given details:\\n{table}\\n\\nKindly refer below matching invoice details for reference, Same have been attached in Info tab.\\n\\nMatching PO \\t Matching Invoice \\t Barcode\\n{m_details}\\n\\nSID: {c_sid}\\nPO: {c_po}\\n\\nNOTE : FC Team we are able to units are short at SID Level and No overages are Found.\\n\\n{rules_text}\\n\\nRegards, \\nMUKESH | pathlmuk\\nROW IB\"\"\"\n",
    "        elif scenario == \"Claiming & Matching Short\":\n",
    "            m_inv_str = \", \".join(matched_invoices)\n",
    "            finding = f\"We are able to see that unit\u2019s shortage received in claiming shipment (SID#{c_sid}) and matching shipment level where we found shortages ({m_inv_str}).\"\n",
    "            table = self._get_asin_table(all_shortages)\n",
    "            text = f\"\"\"Hello FC Team/SLP Team, \\n\\nNote: We have already performed the all the virtual research/checks such as cross receiving, overage, REBNI, adjustments etc.  and need FC support for physical search on floor to locate missing units as per revised SOP and update at the earliest. \\n\\n{finding}\\n\\nClaiming SID#{c_sid} | Total Billed Qty - {int(t_inv)} | Received Qty - {int(t_rec)} \\n\\nPlease locate the following ASINs that are missing from below given details:\\n{table}\\n\\nSID: {c_sid}\\nPO: {c_po}\\n\\nNOTE : FC Team we are able to units are short at SID Level and No overages are Found.\\n\\n{rules_text}\\n\\nRegards, \\nMUKESH | pathlmuk\\nROW IB\"\"\"\n",
    "        elif scenario == \"REBNI\":\n",
    "            rebnis = [f\"{r.get('po','')} \\t {r.get('asin','')} \\t {r.get('sid','')} \\t {safe_num(r.get('cp_status',0)):.2f} \\t {int(safe_num(r.get('inv_qty',0))-safe_num(r.get('rec_qty',0)))}\" for r in all_shortages]\n",
    "            rebni_table = \"\\n\".join(rebnis) if rebnis else \"[No REBNI shortagesidentified]\"\n",
    "            text = f\"\"\"Hello Team, \\n\\nWe see that vendor sent overages of mismatch ASIN in claiming SID - {c_sid}. Below suggested REBNIs are available as per the current REBNI report in line with same state FCs and CP criteria. \\n\\nPO \\t ASIN \\t SID \\t CP \\t Available REBNI\\n{rebni_table}\\n\\nPlease check and utilize the REBNI and update the remaining PQV units. If suggested REBNI are utilized somewhere else, then share Invoice, ASIN and PO level details along with Invoice copy where its matched for validation. \\n\\nNote: \\n1.\\tIf Suggested REBNI comes under same CP limit, the shipment ID and PO shouldn't be the factor. \\n2.\\tSuggested REBNI comes under same shipment, them CP variance isn't the factor.\\n\\nRegards, \\nMUKESH | pathlmuk\\nROW IB\"\"\"\n",
    "        else: text = f\"Scenario '{scenario}' template will be provided later.\"\n",
    "        self.text_area.delete(\"1.0\", tk.END); self.text_area.insert(tk.END, text)\n\n"
]

# Part 3: Lines 1302-End (Suffix with show_correspondence injection)
suffix = lines[1301:]
final_suffix = []
for i, line in enumerate(suffix):
    final_suffix.append(line)
    # Inject show_correspondence after confirm_edits (found around original search point)
    if 'def confirm_edits(self):' in line:
        # We need to find the end of that method.
        # It's better to just append it later or find a cleaner spot.
        pass

# Actually, the trigger was missing from the class definition entirely.
# Let's rebuild the Suffix logically.
preview_panel_range = []
rest_of_file = []
in_preview = False
in_rest = False

for line in suffix:
    if 'class PreviewPanel(tk.Toplevel):' in line: in_preview = True
    if in_preview and 'def confirm_edits(self):' in line:
        # Add the method after confirm_edits block (roughly)
        # To be safe, I'll just append it to the end of the class.
        pass
    # I'll use a simpler search/replace for the method insertion.

# WRITE TO FILE
with open(target_file, 'w', encoding='utf-8') as f:
    f.writelines(prefix)
    f.writelines(new_dialog)
    f.writelines(suffix)

# Now inject show_correspondence surgically
with open(target_file, 'r', encoding='utf-8') as f:
    content = f.read()

method_code = \"\"\"
    def show_correspondence(self):
        \"\"\"Launch the Scenario Selection dialog with current investigation data.\"\"\"
        all_rows = self.get_all_rows()
        if not all_rows:
            messagebox.showinfo(\"No Data\", \"No investigation data available yet.\", parent=self)
            return
        CorrespondenceDialog(self, all_rows)
\"\"\"

# Insert into PreviewPanel after the constructor's tag configuration loop
insertion_point = \"self.tree.tag_configure(tag, background=bg, foreground=fg)\"
if insertion_point in content:
    content = content.replace(insertion_point, insertion_point + \"\\n\\n\" + method_code)
else:
    # fallback: end of PreviewPanel (before DATA LOADERS)
    content = content.replace(\"# \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\\n#  DATA LOADERS\", method_code + \"\\n\\n# \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\\n#  DATA LOADERS\")

with open(target_file, 'w', encoding='utf-8') as f:
    f.write(content)

print(\"SUCCESS: File reconstructed and show_correspondence injected.\")
