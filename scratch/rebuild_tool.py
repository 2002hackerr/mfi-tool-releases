
import os

target = r"c:\Users\Mukesh_Maruthi\MFI_Tool\row_ib_investigation_tool_v5_8_12-My_Fix.py"

with open(target, "r", encoding="utf-8") as f:
    lines = f.readlines()

# Prefix: Lines 1-848 (Indices 0 to 847)
prefix = lines[:848]

# Suffix: Lines 1302+ (Indices 1301 onwards)
# Line 1302 starts "class PreviewPanel(tk.Toplevel):"
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
    "        \n",
    "        self.scenarios = [\n",
    "            'Claiming Short',\n",
    "            'Matching',\n",
    "            'Claiming & Matching Short',\n",
    "            'REBNI'\n",
    "        ]\n",
    "        \n",
    "        self.v_code_var = tk.StringVar(value='[Enter Vendor Code]')\n",
    "        self.fc_id_var = tk.StringVar(value='[Enter FC ID]')\n",
    "        self._build_ui()\n",
    "        \n",
    "    def _build_ui(self):\n",
    "        f = tk.Frame(self, bg='#0f0f1a', padx=20, pady=20)\n",
    "        f.pack(fill='both', expand=True)\n",
    "        \n",
    "        # Scenario Selection\n",
    "        sel_f = tk.LabelFrame(f, text=' Step 1: Select Scenario ', fg='#e0e0e0', bg='#0f0f1a')\n",
    "        sel_f.pack(fill='x', pady=10)\n",
    "        self.scenario_var = tk.StringVar()\n",
    "        self.scenario_cb = ttk.Combobox(sel_f, textvariable=self.scenario_var, values=self.scenarios, state='readonly')\n",
    "        self.scenario_cb.pack(fill='x', padx=10, pady=10)\n",
    "        self.scenario_cb.bind('<<ComboboxSelected>>', self.generate_text)\n",
    "        \n",
    "        # Manual Inputs\n",
    "        inp_f = tk.LabelFrame(f, text=' Step 2: Details ', fg='#e0e0e0', bg='#0f0f1a')\n",
    "        inp_f.pack(fill='x', pady=10)\n",
    "        tk.Entry(inp_f, textvariable=self.v_code_var, width=20).pack(side='left', padx=10, pady=5)\n",
    "        tk.Entry(inp_f, textvariable=self.fc_id_var, width=20).pack(side='left', padx=10, pady=5)\n",
    "        \n",
    "        # Text Area\n",
    "        self.text_area = tk.Text(f, height=25, font=('Consolas', 10), bg='#1e1e3a', fg='white')\n",
    "        self.text_area.pack(fill='both', expand=True)\n",
    "        \n",
    "        # Control Button\n",
    "        tk.Button(f, text='Copy Correspondence', command=self.copy_to_clip, bg='#2d6a4f', fg='white').pack(pady=10)\n",
    "        \n",
    "    def copy_to_clip(self):\n",
    "        self.clipboard_clear()\n",
    "        self.clipboard_append(self.text_area.get('1.0', tk.END).strip())\n",
    "        messagebox.showinfo('Copied', 'Correspondence copied to clipboard!')\n",
    "\n",
    "    def generate_text(self, event=None):\n",
    "        scenario = self.scenario_var.get()\n",
    "        claiming_row = next((r for r in self.all_rows if r.get('depth', 0) == 0), {})\n",
    "        sid = claiming_row.get('sid', '[SID]')\n",
    "        po = claiming_row.get('po', '[PO]')\n",
    "        v_code = self.v_code_var.get()\n",
    "        \n",
    "        txt = f'Hello FC Team,\\n\\nScenario: {scenario}\\nSID: {sid}\\nPO: {po}\\nVendor: {v_code}\\n\\nRegards,\\nMUKESH\\nROW IB'\n",
    "        self.text_area.delete('1.0', tk.END)\n",
    "        self.text_area.insert(tk.END, txt)\n"
]

# Write reconstructed file
with open(target, "w", encoding="utf-8") as f:
    f.writelines(prefix + midfold + suffix)

# Inject show_correspondence into PreviewPanel
with open(target, "r", encoding="utf-8") as f:
    content = f.read()

method_injection = \"\"\"
    def show_correspondence(self):
        all_rows = self.get_all_rows()
        if not all_rows:
            messagebox.showinfo('No Data', 'No investigation data available.')
            return
        CorrespondenceDialog(self, all_rows)
\"\"\"

# Insert before DATA LOADERS section
insertion_marker = \"# \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\\n#  DATA LOADERS\"
if insertion_marker in content:
    content = content.replace(insertion_marker, method_injection + \"\\n\\n\" + insertion_marker)

with open(target, \"w\", encoding=\"utf-8\") as f:
    f.write(content)

print(\"Recovery complete.\")
