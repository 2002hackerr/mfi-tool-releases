
import os

target = r"c:\Users\Mukesh_Maruthi\MFI_Tool\row_ib_investigation_tool_v5_8_12-My_Fix.py"

with open(target, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Part 1: Stable prefix (Lines 1-848)
# Indices 0 to 847
prefix = lines[:848]

# Part 2: Clean CorrespondenceDialog logic
clean_dialog = [
    "class CorrespondenceDialog(tk.Toplevel):\n",
    "    def __init__(self, parent, all_rows):\n",
    "        super().__init__(parent)\n",
    "        self.all_rows = all_rows\n",
    "        self.title('Scenario Selection — Get Correspondence — v5.9.0')\n",
    "        self.geometry('950x800')\n",
    "        self.configure(bg='#0f0f1a')\n",
    "        self.transient(parent)\n",
    "        self.grab_set()\n",
    "        \n",
    "        self.scenarios = [\n",
    "            'Claiming Short',\n",
    "            'Matching',\n",
    "            'Claiming & Matching Short',\n",
    "            'REBNI',\n",
    "            'Matching to Invalid Invoice [Dev]',\n",
    "            'Dummy Invoice [Dev]',\n",
    "            'Difference in IBC versus PBC [Dev]'\n",
    "        ]\n",
    "        \n",
    "        self.v_code_var = tk.StringVar(); self.fc_id_var = tk.StringVar()\n",
    "        self.fc_id_var.set('[Enter FC ID]')\n",
    "        self.v_code_var.set('[Enter Vendor Code]')\n",
    "        self._build_ui()\n",
    "        \n",
    "    def _build_ui(self):\n",
    "        f = tk.Frame(self, bg='#0f0f1a', padx=20, pady=20)\n",
    "        f.pack(fill='both', expand=True)\n",
    "        top_h = tk.Frame(f, bg='#0f0f1a')\n",
    "        top_h.pack(fill='x', pady=(0,15))\n",
    "        tk.Label(top_h, text='Investigation Correspondence Generator',\n",
    "                 fg='#4a9eff', bg='#0f0f1a',\n",
    "                 font=('Segoe UI', 14, 'bold')).pack(side='left')\n",
    "        sel_f = tk.LabelFrame(f, text=' Step 1: Select Scenario ', fg='#e0e0e0', bg='#0f0f1a', font=('Segoe UI', 9, 'bold'))\n",
    "        sel_f.pack(fill='x', pady=10, padx=2)\n",
    "        self.scenario_var = tk.StringVar()\n",
    "        self.scenario_combobox = ttk.Combobox(sel_f, textvariable=self.scenario_var, \n",
    "                                             values=self.scenarios, state='readonly', \n",
    "                                             font=('Segoe UI', 11))\n",
    "        self.scenario_combobox.pack(fill='x', padx=10, pady=10)\n",
    "        self.scenario_combobox.bind('<<ComboboxSelected>>', self.generate_text)\n",
    "        inp_f = tk.LabelFrame(f, text=' Step 2: Manual Details (Optional) ', fg='#e0e0e0', bg='#0f0f1a', font=('Segoe UI', 9, 'bold'))\n",
    "        inp_f.pack(fill='x', pady=10, padx=2)\n",
    "        grid_f = tk.Frame(inp_f, bg='#0f0f1a')\n",
    "        grid_f.pack(fill='x', padx=10, pady=5)\n",
    "        tk.Label(grid_f, text='Vendor Code:', bg='#0f0f1a', fg='#cccccc').grid(row=0, column=0, sticky='w')\n",
    "        tk.Entry(grid_f, textvariable=self.v_code_var, width=20).grid(row=0, column=1, padx=5, pady=2)\n",
    "        tk.Label(grid_f, text='FC ID:', bg='#0f0f1a', fg='#cccccc').grid(row=0, column=2, sticky='w', padx=(20,0))\n",
    "        tk.Entry(grid_f, textvariable=self.fc_id_var, width=20).grid(row=0, column=3, padx=5, pady=2)\n",
    "        tk.Label(f, text=' Step 3: Review and Copy Generated Text:',\n",
    "                 fg='#cccccc', bg='#0f0f1a',\n",
    "                 font=('Segoe UI', 10)).pack(anchor='w', pady=(10, 5))\n",
    "        self.text_area = tk.Text(f, height=25, font=('Consolas', 10), \n",
    "                                 bg='#1e1e3a', fg='white', padx=10, pady=10,\n",
    "                                 insertbackground='white', wrap='word')\n",
    "        self.text_area.pack(fill='both', expand=True)\n",
    "        btn_f = tk.Frame(f, bg='#0f0f1a')\n",
    "        btn_f.pack(fill='x', pady=(20, 0))\n",
    "        tk.Button(btn_f, text='\ud83d\udccb Copy Ticket Correspondence', \n",
    "                  command=self.copy_to_clip,\n",
    "                  bg='#2d6a4f', fg='white', font=('Segoe UI', 11, 'bold'),\n",
    "                  relief='flat', padx=20, pady=10, cursor='hand2').pack(side='left')\n",
    "        tk.Button(btn_f, text='Close', command=self.destroy,\n",
    "                  bg='#4a2020', fg='white', font=('Segoe UI', 10),\n",
    "                  relief='flat', padx=20, pady=10, cursor='hand2').pack(side='right')\n",
    "\n",
    "    def copy_to_clip(self):\n",
    "        text = self.text_area.get('1.0', tk.END).strip()\n",
    "        if not text: return\n",
    "        self.clipboard_clear()\n",
    "        self.clipboard_append(text)\n",
    "        messagebox.showinfo('Copied', 'Correspondence copied to clipboard!')\n",
    "\n",
    "    def _get_asin_table(self, rows, cols=9):\n",
    "        header = 'Barcode\\tInv no\\tSID\\tPO\\tASIN\\tInv Qty\\tRec Qty\\tMissing QTY\\tCost Price'\n",
    "        lines = [header]\n",
    "        for r in rows:\n",
    "            iq, rq = safe_num(r.get('inv_qty',0)), safe_num(r.get('rec_qty',0))\n",
    "            missing = iq - rq\n",
    "            if missing <= 0: continue\n",
    "            cp = safe_num(r.get('cp_status',0))\n",
    "            row_data = [str(r.get('barcode','')).strip(), str(r.get('invoice','')).strip(), str(r.get('sid','')).strip(),\n",
    "                        str(r.get('po','')).strip(), str(r.get('asin','')).strip(), str(int(iq) if iq==int(iq) else iq),\n",
    "                        str(int(rq) if rq==int(rq) else rq), str(int(missing) if missing==int(missing) else missing), f'{cp:.2f}']\n",
    "            lines.append('\\t'.join(row_data))\n",
    "        return '\\n'.join(lines)\n",
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
    "        c_sid = claiming_row.get('sid','') if claiming_row else '[Enter SID]'\n",
    "        c_po  = claiming_row.get('po','')  if claiming_row else '[Enter PO]'\n",
    "        t_inv = sum(safe_num(r.get('inv_qty',0)) for r in self.all_rows if r.get('depth',0)==0)\n",
    "        t_rec = sum(safe_num(r.get('rec_qty',0)) for r in self.all_rows if r.get('depth',0)==0)\n",
    "        v_code = self.v_code_var.get(); f_id = self.fc_id_var.get()\n",
    "        rules = 'i. FC to perform physical investigation only for missing units and share complete findings. \\nii. If >50k PQV, flip to FC SLP. \\niii. Flip to ROW IB only after FC/SLP complete.'\n",
    "        \n",
    "        if scenario == 'Claiming Short':\n",
    "            table = self._get_asin_table(all_shortages)\n",
    "            text = f'Hello FC Team, \\n\\nClaiming SID#{c_sid} | Billed Qty - {int(t_inv)} | Received Qty - {int(t_rec)} \\nPlease locate ASINs missing from PO#{c_po}\\n{table}\\nSID: {c_sid}\\nPO: {c_po}\\nVendor Code: {v_code}\\nFC ID: {f_id}\\n\\n{rules}\\n\\nRegards, \\nMUKESH | pathlmuk\\nROW IB'\n",
    "        elif scenario == 'Matching':\n",
    "            m_inv_str = ', '.join(matched_invoices); table = self._get_asin_table(all_shortages)\n",
    "            m_details = '\\n'.join(matching_details) if matching_details else '[No matches]'\n",
    "            text = f'Hello FC Team, \\n\\nSID#{c_sid} | Billed Qty - {int(t_inv)} | Received Qty - {int(t_rec)} \\nMatched with invoices ({m_inv_str}). Locate missing units:\\n{table}\\nMatching details:\\nPO \\t Invoice \\t Barcode\\n{m_details}\\nSID: {c_sid}\\nPO: {c_po}\\n\\n{rules}\\n\\nRegards, \\nMUKESH | pathlmuk\\nROW IB'\n",
    "        elif scenario == 'Claiming & Matching Short':\n",
    "            m_inv_str = ', '.join(matched_invoices); table = self._get_asin_table(all_shortages)\n",
    "            text = f'Hello FC/SLP Team, \\n\\nShortage found in claiming SID#{c_sid} and matched level ({m_inv_str}).\\nClaiming SID#{c_sid} | Billed Qty - {int(t_inv)} | Received Qty - {int(t_rec)} \\nLocate missing units:\\n{table}\\n\\n{rules}\\n\\nRegards, \\nMUKESH | pathlmuk\\nROW IB'\n",
    "        elif scenario == 'REBNI':\n",
    "            rebnis = [f\"{r.get('po','')} \\t {r.get('asin','')} \\t {r.get('sid','')} \\t {safe_num(r.get('cp_status',0)):.2f} \\t {int(safe_num(r.get('inv_qty',0))-safe_num(r.get('rec_qty',0)))}\" for r in all_shortages]\n",
    "            rebni_table = '\\n'.join(rebnis) if rebnis else '[No REBNI shortages]'\n",
    "            text = f'Hello Team, \\n\\nSuggested REBNIs available for overages in SID - {c_sid}.\\nPO \\t ASIN \\t SID \\t CP \\t Available REBNI\\n{rebni_table}\\nPlease check and utilize.\\n\\nRegards, \\nMUKESH | pathlmuk\\nROW IB'\n",
    "        else: text = f'Scenario {scenario} template provided later.'\n",
    "        self.text_area.delete('1.0', tk.END); self.text_area.insert(tk.END, text)\n\n"
]

# Part 3: Stable suffix (Lines 1302 to End)
# Indices 1301 to end
suffix = lines[1301:]

with open(target, 'w', encoding='utf-8') as f:
    f.writelines(prefix)
    f.writelines(clean_dialog)
    f.writelines(suffix)

# Part 4: Inject show_correspondence method into PreviewPanel
with open(target, 'r', encoding='utf-8') as f:
    content = f.read()

method = \"\"\"
    def show_correspondence(self):
        \"\"\"Launch the Scenario Selection dialog with current investigation data.\"\"\"
        all_rows = self.get_all_rows()
        if not all_rows:
            messagebox.showinfo(\"No Data\", \"No investigation data available yet.\", parent=self)
            return
        CorrespondenceDialog(self, all_rows)
\"\"\"

# Insert after confirm_edits (around line 1590)
if 'def confirm_edits(self):' in content and 'CorrespondenceDialog(self, all_rows)' not in content:
    # We'll split by DATA LOADERS and insert before it
    parts = content.split('# \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\\n#  DATA LOADERS')
    if len(parts) == 2:
        new_content = parts[0] + method + \"\\n\\n# \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\\n#  DATA LOADERS\" + parts[1]
        with open(target, 'w', encoding='utf-8') as f:
            f.write(new_content)
    else:
        print(\"Could not find DATA LOADERS split point\")
else:
    print(\"Show correspondence already exists or confirm_edits missing\")
