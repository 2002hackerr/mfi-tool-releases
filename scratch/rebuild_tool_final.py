
import os

target = r"c:\Users\Mukesh_Maruthi\MFI_Tool\row_ib_investigation_tool_v5_8_12-My_Fix.py"

with open(target, "r", encoding="utf-8") as f:
    lines = f.readlines()

# Indices for 849-1301 (Line numbers 849-1301)
# prefix: indices 0-847 (848 lines)
prefix = lines[:848]
# suffix: indices 1301 onwards (Line 1302 onwards)
suffix = lines[1301:]

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
    "        self.scenarios = ['Claiming Short', 'Matching', 'Claiming & Matching Short', 'REBNI']\n",
    "        self.v_code_var = tk.StringVar(value='[Vendor Code]')\n",
    "        self.fc_id_var = tk.StringVar(value='[FC ID]')\n",
    "        self._build_ui()\n",
    "    def _build_ui(self):\n",
    "        f = tk.Frame(self, bg='#0f0f1a', padx=20, pady=20); f.pack(fill='both', expand=True)\n",
    "        sel_f = tk.LabelFrame(f, text=' Step 1: Select Scenario ', fg='#e0e0e0', bg='#0f0f1a')\n",
    "        sel_f.pack(fill='x', pady=10)\n",
    "        self.scenario_var = tk.StringVar()\n",
    "        self.scenario_cb = ttk.Combobox(sel_f, textvariable=self.scenario_var, values=self.scenarios, state='readonly')\n",
    "        self.scenario_cb.pack(fill='x', padx=10, pady=10)\n",
    "        self.scenario_cb.bind('<<ComboboxSelected>>', self.generate_text)\n",
    "        inp_f = tk.LabelFrame(f, text=' Step 2: Details ', fg='#e0e0e0', bg='#0f0f1a')\n",
    "        inp_f.pack(fill='x', pady=10)\n",
    "        tk.Entry(inp_f, textvariable=self.v_code_var).pack(side='left', padx=10, pady=5)\n",
    "        tk.Entry(inp_f, textvariable=self.fc_id_var).pack(side='left', padx=10, pady=5)\n",
    "        self.text_area = tk.Text(f, height=25, font=('Consolas', 10), bg='#1e1e3a', fg='white')\n",
    "        self.text_area.pack(fill='both', expand=True)\n",
    "        tk.Button(f, text='Copy Correspondence', command=self.copy_to_clip, bg='#2d6a4f', fg='white').pack(pady=10)\n",
    "    def copy_to_clip(self):\n",
    "        self.clipboard_clear(); self.clipboard_append(self.text_area.get('1.0', 'end').strip()); messagebox.showinfo('Copied', 'Done')\n",
    "    def generate_text(self, event=None):\n",
    "        s = self.scenario_var.get(); row = next((r for r in self.all_rows if r.get('depth',0)==0), {})\n",
    "        sid, po = row.get('sid','?'), row.get('po','?')\n",
    "        txt = f\"Hello FC,\\n\\nScenario: {s}\\nSID: {sid}\\nPO: {po}\\n\\nRegards,\\nMUKESH\\nROW IB\"\n",
    "        self.text_area.delete('1.0', 'end'); self.text_area.insert('1.0', txt)\n"
]

with open(target, 'w', encoding='utf-8') as f:\n    f.writelines(prefix + midfold + suffix)

with open(target, 'r', encoding='utf-8') as f:\n    content = f.read()

method = \"\"\"
    def show_correspondence(self):
        all_rows = self.get_all_rows()
        if not all_rows:
            messagebox.showinfo('No Data', 'No investigation data available.')
            return
        CorrespondenceDialog(self, all_rows)
\"\"\"
marker = \"# \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\\n#  DATA LOADERS\"
if marker in content:\n    content = content.replace(marker, method + '\\n\\n' + marker)

with open(target, 'w', encoding='utf-8') as f:\n    f.write(content)
