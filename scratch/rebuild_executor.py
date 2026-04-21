
import os
import tkinter as tk
from tkinter import messagebox, ttk

target = r"c:\Users\Mukesh_Maruthi\MFI_Tool\row_ib_investigation_tool_v5_8_12-My_Fix.py"

def rebuild():
    with open(target, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    prefix = lines[:848]
    suffix = lines[1301:]
    
    mid = [
        "class CorrespondenceDialog(tk.Toplevel):\n",
        "    def __init__(self, parent, all_rows):\n",
        "        super().__init__(parent)\n",
        "        self.all_rows = all_rows\n",
        "        self.title('Scenario Selection — v5.8.12')\n",
        "        self.geometry('900x750')\n",
        "        self.configure(bg='#0f0f1a')\n",
        "        self.transient(parent)\n",
        "        self.grab_set()\n",
        "        self.scenarios = ['Claiming Short', 'Matching', 'Claiming & Matching Short', 'REBNI']\n",
        "        self.v_code = tk.StringVar(value='[Vendor Code]')\n",
        "        self.fc_id = tk.StringVar(value='[FC ID]')\n",
        "        self._build_ui()\n",
        "    def _build_ui(self):\n",
        "        f = tk.Frame(self, bg='#0f0f1a', padx=20, pady=20); f.pack(fill='both', expand=True)\n",
        "        self.sv = tk.StringVar()\n",
        "        self.cb = ttk.Combobox(f, textvariable=self.sv, values=self.scenarios, state='readonly')\n",
        "        self.cb.pack(fill='x', pady=10)\n",
        "        self.cb.bind('<<ComboboxSelected>>', self.gen)\n",
        "        self.ta = tk.Text(f, height=22, font=('Consolas', 10), bg='#1e1e3a', fg='white')\n",
        "        self.ta.pack(fill='both', expand=True)\n",
        "        tk.Button(f, text='Copy Correspondence', command=self.cp, bg='#2d6a4f', fg='white').pack(pady=10)\n",
        "    def cp(self):\n",
        "        self.clipboard_clear(); self.clipboard_append(self.ta.get('1.0', 'end').strip()); messagebox.showinfo('Copied', 'Done')\n",
        "    def gen(self, e=None):\n",
        "        s = self.sv.get(); r = next((x for x in self.all_rows if x.get('depth',0)==0), {})\n",
        "        sid, po = r.get('sid','?'), r.get('po','?')\n",
        "        txt = f\"Hello FC,\\n\\nScenario: {s}\\nSID: {sid}\\nPO: {po}\\n\\nRegards,\\nMUKESH\\nROW IB\"\n",
        "        self.ta.delete('1.0', 'end'); self.ta.insert('1.0', txt)\n"
    ]
    
    with open(target, 'w', encoding='utf-8') as f:
        f.writelines(prefix + mid + suffix)
    
    with open(target, 'r', encoding='utf-8') as f:
        content = f.read()
    
    method = """
    def show_correspondence(self):
        all_rows = self.get_all_rows()
        if not all_rows:
            messagebox.showinfo('No Data', 'No investigation data available.')
            return
        CorrespondenceDialog(self, all_rows)
"""
    marker = "# \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\\n#  DATA LOADERS"
    if marker in content:
        content = content.replace(marker, method + "\\n\\n" + marker)
        with open(target, 'w', encoding='utf-8') as f:
            f.write(content)

if __name__ == "__main__":
    rebuild()
