
import os

target_path = r"c:\Users\Mukesh_Maruthi\MFI_Tool\row_ib_investigation_tool_v5_8_12-My_Fix.py"

with open(target_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Prefix: 1-848 (Indices 0 to 847)
prefix = lines[:848]

# Suffix: 1302 to End (Indices 1301 to end)
suffix = lines[1301:]

# Midfold: Finalized CorrespondenceDialog
midfold = [
    "class CorrespondenceDialog(tk.Toplevel):\n",
    "    def __init__(self, parent, all_rows):\n",
    "        super().__init__(parent)\n",
    "        self.all_rows = all_rows\n",
    "        self.title('Scenario Selection — Get Correspondence — v5.9.0')\n",
    "        self.geometry('950x800')\n",
    "        self.configure(bg='#0f0f1a')\n",
    "        self.transient(parent)\n",
    "        self.grab_set()\n",
    "        self.scenarios = ['Claiming Short', 'Matching', 'Claiming & Matching Short', 'REBNI', 'Matching to Invalid Invoice [Dev]', 'Dummy Invoice [Dev]', 'Difference in IBC versus PBC [Dev]']\n",
    "        self.v_code_var = tk.StringVar(); self.fc_id_var = tk.StringVar()\n",
    "        self.fc_id_var.set('[Enter FC ID]'); self.v_code_var.set('[Enter Vendor Code]')\n",
    "        self._build_ui()\n",
    "    def _build_ui(self):\n",
    "        f = tk.Frame(self, bg='#0f0f1a', padx=20, pady=20); f.pack(fill='both', expand=True)\n",
    "        top_h = tk.Frame(f, bg='#0f0f1a'); top_h.pack(fill='x', pady=(0,15))\n",
    "        tk.Label(top_h, text='Investigation Correspondence Generator', fg='#4a9eff', bg='#0f0f1a', font=('Segoe UI', 14, 'bold')).pack(side='left')\n",
    "        sel_f = tk.LabelFrame(f, text=' Step 1: Select Scenario ', fg='#e0e0e0', bg='#0f0f1a', font=('Segoe UI', 9, 'bold')); sel_f.pack(fill='x', pady=10, padx=2)\n",
    "        self.scenario_var = tk.StringVar()\n",
    "        self.scenario_combobox = ttk.Combobox(sel_f, textvariable=self.scenario_var, values=self.scenarios, state='readonly', font=('Segoe UI', 11))\n",
    "        self.scenario_combobox.pack(fill='x', padx=10, pady=10)\n",
    "        self.scenario_combobox.bind('<<ComboboxSelected>>', self.generate_text)\n",
    "        inp_f = tk.LabelFrame(f, text=' Step 2: Manual Details (Optional) ', fg='#e0e0e0', bg='#0f0f1a', font=('Segoe UI', 9, 'bold')); inp_f.pack(fill='x', pady=10, padx=2)\n",
    "        grid_f = tk.Frame(inp_f, bg='#0f0f1a'); grid_f.pack(fill='x', padx=10, pady=5)\n",
    "        tk.Label(grid_f, text='Vendor Code:', bg='#0f0f1a', fg='#cccccc').grid(row=0, column=0, sticky='w')\n",
    "        tk.Entry(grid_f, textvariable=self.v_code_var, width=20).grid(row=0, column=1, padx=5, pady=2)\n",
    "        tk.Label(grid_f, text='FC ID:', bg='#0f0f1a', fg='#cccccc').grid(row=0, column=2, sticky='w', padx=(20,0))\n",
    "        tk.Entry(grid_f, textvariable=self.fc_id_var, width=20).grid(row=0, column=3, padx=5, pady=2)\n",
    "        tk.Label(f, text=' Step 3: Review and Copy Generated Text:', fg='#cccccc', bg='#0f0f1a', font=('Segoe UI', 10)).pack(anchor='w', pady=(10, 5))\n",
    "        self.text_area = tk.Text(f, height=20, font=('Consolas', 10), bg='#1e1e3a', fg='white', padx=10, pady=10, insertbackground='white', wrap='word'); self.text_area.pack(fill='both', expand=True)\n",
    "        btn_f = tk.Frame(f, bg='#0f0f1a'); btn_f.pack(fill='x', pady=(20, 0))\n",
    "        tk.Button(btn_f, text='Copy Correspondence', command=self.copy_to_clip, bg='#2d6a4f', fg='white', font=('Segoe UI', 11, 'bold'), relief='flat', padx=20, pady=10, cursor='hand2').pack(side='left')\n",
    "        tk.Button(btn_f, text='Close', command=self.destroy, bg='#4a2020', fg='white', font=('Segoe UI', 10), relief='flat', padx=20, pady=10, cursor='hand2').pack(side='right')\n",
    "    def copy_to_clip(self):\n",
    "        self.clipboard_clear(); self.clipboard_append(self.text_area.get('1.0', tk.END).strip()); messagebox.showinfo('Copied', 'Copied to clipboard!')\n",
    "    def _get_asin_table(self, rows):\n",
    "        lines = ['Barcode\\tInv no\\tSID\\tPO\\tASIN\\tInv Qty\\tRec Qty\\tMissing QTY\\tCost Price']\n",
    "        for r in rows:\n",
    "            iq, rq, cp = safe_num(r.get('inv_qty',0)), safe_num(r.get('rec_qty',0)), safe_num(r.get('cp_status',0))\n",
    "            if iq > rq: lines.append(f\"{r.get('barcode','')}\\t{r.get('invoice','')}\\t{r.get('sid','')}\\t{r.get('po','')}\\t{r.get('asin','')}\\t{int(iq)}\\t{int(rq)}\\t{int(iq-rq)}\\t{cp:.2f}\")\n",
    "        return '\\n'.join(lines)\n",
    "    def generate_text(self, event=None):\n",
    "        s = self.scenario_var.get(); claiming = None; shorts = []\n",
    "        for r in self.all_rows:\n",
    "            if not str(r.get('barcode','')).startswith('['):\n",
    "                if safe_num(r.get('inv_qty',0)) > safe_num(r.get('rec_qty',0)): shorts.append(r)\n",
    "                if r.get('depth',0) == 0: claiming = r\n",
    "        sid, po = claiming.get('sid',''), claiming.get('po','')\n",
    "        rules = 'Rules: FC investigation for missing units only. If >50k, flip to SLP. Flip to IB last.'\n",
    "        if s == 'Claiming Short':\n",
    "            t = self._get_asin_table(shorts)\n",
    "            txt = f'Hello FC, Locate missing units from SID#{sid}/PO#{po}.\\n{t}\\n{rules}'\n",
    "        elif s == 'REBNI':\n",
    "            txt = f'Hello Team, REBNIs suggested for overages in SID#{sid}.\\n{rules}'\n",
    "        else: txt = f'Template for {s} generated later.'\n",
    "        self.text_area.delete('1.0', tk.END); self.text_area.insert(tk.END, txt)\n"
]

with open(target_path, 'w', encoding='utf-8') as f:\n    f.writelines(prefix + midfold + suffix)
