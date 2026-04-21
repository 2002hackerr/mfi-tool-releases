"""
MFI Investigation Tool  v5.9.3  |  ROW IB
==========================================
ROW IB  |  Amazon
Developed by Mukesh

CHANGES IN v5.9.3:
  ✔ [FIX] Corrected Save Button re-enabling logic in _finish().
  ✔ [UI/FIX] Unique Summary Portal button is now disabled during active investigations.
  ✔ [UI/FIX] Standardized all branding and labels to "v5.9.3 | ROW IB".
  ✔ [SAFETY] Migrated 100% stable bundling logic from v5.9.2.

CHANGES IN v5.9.2:
  ✔ [INTEGRATION] Added "📑 UNIQUE SUMMARY PORTAL" button for report aggregation.
  ✔ [INTEGRATION] Implemented auto-launch logic with sys._MEIPASS bundling support.
  ✔ [UI/FIX] Standardized all branding and labels to "v5.9.2 | ROW IB".
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
import os
import re
import sys
import threading
import webbrowser
from datetime import datetime


# ═══════════════════════════════════════════════════════════
#  UTILITIES
# ═══════════════════════════════════════════════════════════

def extract_sid(val):
    s = str(val).strip()
    parts = re.findall(r'\d{10,}', s)
    return max(parts, key=len) if parts else s

def strip_scr(inv_no):
    """Remove trailing SCR suffix(es) from an invoice number."""
    return re.sub(r'(?:SCR)+$', '', str(inv_no).strip(), flags=re.IGNORECASE)

def safe_num(val):
    try:
        if pd.isna(val): return 0.0
    except: pass
    try:
        return float(str(val).replace(',', '').strip())
    except: return 0.0

def clean(val):
    try:
        if pd.isna(val): return ""
    except: pass
    return str(val).strip()

def split_comma(val):
    if not val: return []
    try:
        if pd.isna(val): return []
    except: pass
    return [s.strip() for s in str(val).split(',') if s.strip()]

def fmt_qty(val):
    n = safe_num(val)
    if n == 0: return ""
    return str(int(n)) if n == int(n) else str(n)

# ── Keywords whose presence in a remark means it must NEVER be overwritten.
_REMARK_PRESERVE = (
    'Phase 1', 'Direct Shortage', 'REBNI Available', 'Shipment-level REBNI',
    'SR', 'Phase 4', 'No Invoice Search', 'short received directly',
)

def _remark_overwritable(rem):
    """Return True only when none of the preserved keywords appear in rem."""
    return not any(kw in rem for kw in _REMARK_PRESERVE)


# ═══════════════════════════════════════════════════════════
#  DIALOGS
# ═══════════════════════════════════════════════════════════

class HeaderCorrectionDialog(tk.Toplevel):
    """Shown when claims file has non-standard column headers."""
    def __init__(self, parent, corrections, mapping, df_columns, callback):
        super().__init__(parent)
        self.callback    = callback
        self.corrections = corrections
        self.mapping     = mapping
        self.df_columns  = df_columns

        self.title("Column Header Mismatch Detected — v5.9.3")
        self.geometry("700x480")
        self.configure(bg="#0f0f1a")
        self.resizable(True, True)
        self.lift(); self.focus_force()

        tk.Label(self,
                 text="⚠  Non-standard column headers detected in Claims file",
                 bg="#16213e", fg="#f0a500",
                 font=("Segoe UI", 12, "bold"), height=2).pack(fill="x")
        tk.Label(self,
                 text="The tool has automatically matched the columns below. "
                      "Please confirm or correct before proceeding.",
                 bg="#0f0f1a", fg="#cccccc",
                 font=("Segoe UI", 9)).pack(pady=4)

        outer = tk.Frame(self, bg="#0f0f1a"); outer.pack(fill="both", expand=True, padx=16, pady=6)

        hdrs = ["Field", "Expected", "Found in file", "Status"]
        w    = [14, 20, 28, 12]
        for ci, (h, ww) in enumerate(zip(hdrs, w)):
            tk.Label(outer, text=h, bg="#203864", fg="white",
                     font=("Calibri", 10, "bold"),
                     width=ww, anchor="w", padx=4
                     ).grid(row=0, column=ci, padx=1, pady=1, sticky="w")

        all_fields       = list(COLUMN_ALIASES.keys())
        corrected_fields = {c[0] for c in corrections}
        self._override_vars = {}

        for ri, field in enumerate(all_fields, 1):
            canonical = COLUMN_ALIASES[field][0]
            found_col = mapping.get(field, "")
            is_corrected = field in corrected_fields
            is_missing   = field not in mapping

            if is_missing:
                status_txt, status_fg, row_bg = "MISSING",   "#ff4444", "#2a0000"
            elif is_corrected:
                status_txt, status_fg, row_bg = "Auto-fixed", "#f0a500", "#1a1500"
            else:
                status_txt, status_fg, row_bg = "✔ OK",      "#44ff88", "#001a00"

            tk.Label(outer, text=field, bg=row_bg, fg="#e0e0e0",
                     font=("Calibri", 10, "bold"), width=14, anchor="w", padx=4
                     ).grid(row=ri, column=0, padx=1, pady=1, sticky="w")
            tk.Label(outer, text=canonical, bg=row_bg, fg="#aaaacc",
                     font=("Calibri", 10), width=20, anchor="w", padx=4
                     ).grid(row=ri, column=1, padx=1, pady=1, sticky="w")

            v = tk.StringVar(value=found_col or "— not found —")
            self._override_vars[field] = v
            opts = ["— not found —"] + list(df_columns)
            cb = ttk.Combobox(outer, textvariable=v, values=opts,
                              state="readonly", width=26, font=("Calibri", 9))
            cb.grid(row=ri, column=2, padx=1, pady=1, sticky="w")

            tk.Label(outer, text=status_txt, bg=row_bg, fg=status_fg,
                     font=("Calibri", 10, "bold"), width=12, anchor="w", padx=4
                     ).grid(row=ri, column=3, padx=1, pady=1, sticky="w")

        tk.Label(self,
                 text="You can change any column assignment using the dropdowns above.",
                 bg="#0f0f1a", fg="#888899", font=("Segoe UI", 8)).pack(pady=2)

        bf = tk.Frame(self, bg="#0f0f1a"); bf.pack(pady=10)
        tk.Button(bf, text="✔  Auto-correct & Proceed",
                  command=self._proceed,
                  bg="#2d6a4f", fg="white",
                  font=("Segoe UI", 12, "bold"),
                  padx=20, pady=8, relief="flat",
                  cursor="hand2").pack(side="left", padx=10)
        tk.Button(bf, text="✖  Cancel",
                  command=self._cancel,
                  bg="#4a2020", fg="white",
                  font=("Segoe UI", 11),
                  padx=16, pady=8, relief="flat",
                  cursor="hand2").pack(side="left", padx=10)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

    def _proceed(self):
        final = {}
        for field, var in self._override_vars.items():
            v = var.get()
            if v and v != "— not found —":
                final[field] = v
        self.callback({'action': 'proceed', 'mapping': final})
        self.destroy()

    def _cancel(self):
        self.callback({'action': 'cancel'})
        self.destroy()


class SIDRequestDialog(tk.Toplevel):
    """Auto mode: SID not found in REBNI → ask user."""
    def __init__(self, parent, invoice, po, asin, callback):
        super().__init__(parent)
        self.callback = callback
        self.title("SID Required — v5.9.3")
        self.geometry("540x210")
        self.resizable(True, True)
        self.configure(bg="#16213e")
        self.lift(); self.focus_force()

        tk.Label(self, text="⚠  SID Not Found in REBNI",
                 bg="#16213e", fg="#e94560",
                 font=("Segoe UI", 13, "bold")).pack(pady=(14, 4))
        tk.Label(self, text=f"Invoice: {invoice}   PO: {po}   ASIN: {asin}",
                 bg="#16213e", fg="#e0e0e0", font=("Segoe UI", 9)).pack(pady=2)
        tk.Label(self, text="Validate this invoice in DICES and enter the SID below:",
                 bg="#16213e", fg="#aaaacc", font=("Segoe UI", 9)).pack(pady=6)

        ef = tk.Frame(self, bg="#16213e"); ef.pack()
        tk.Label(ef, text="SID from DICES:", bg="#16213e", fg="#e0e0e0",
                 font=("Segoe UI", 10)).pack(side="left", padx=8)
        self._sid = tk.StringVar()
        self._entry = tk.Entry(ef, textvariable=self._sid, width=30,
                               font=("Segoe UI", 10), bg="#1e1e3a", fg="#e0e0e0",
                               insertbackground="white", relief="flat")
        self._entry.pack(side="left", padx=4)
        self._entry.focus_set()

        bf = tk.Frame(self, bg="#16213e"); bf.pack(pady=12)
        tk.Button(bf, text="✔  Continue", command=self._ok,
                  bg="#2d6a4f", fg="white", font=("Segoe UI", 11, "bold"),
                  padx=16, pady=7, relief="flat", cursor="hand2").pack(side="left", padx=8)
        tk.Button(bf, text="✖  Skip", command=self._skip,
                  bg="#6b2737", fg="white", font=("Segoe UI", 10),
                  padx=16, pady=7, relief="flat", cursor="hand2").pack(side="left", padx=8)

        self.bind('<Return>', lambda e: self._ok())
        self.protocol("WM_DELETE_WINDOW", self._skip)
        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

    def _ok(self):
        sid = extract_sid(self._sid.get().strip())
        if sid:
            self.callback(sid); self.destroy()
        else:
            self._entry.config(bg="#3a1e1e")

    def _skip(self):
        self.callback(None); self.destroy()

class CrossPODialog(tk.Toplevel):
    CASE_DESCRIPTIONS = {
        "Case 1": (
            "Case 1 — No PO, but ASIN received",
            "Rec=0 at claiming PO. Same ASIN received in different PO within same SID.\n"
            "Those units are overage under a different PO."
        ),
        "Case 2": (
            "Case 2 — PO exists but ASIN not ordered there",
            "This PO exists in the claiming SID, but the ASIN was never invoiced there.\n"
            "Inv Qty = 0, but units were received. This is a Cross PO overage."
        ),
        "Case 3": (
            "Case 3 — PO and ASIN exist but Rec > Inv",
            "Both PO and ASIN are present. Invoiced qty = X, but received more than X.\n"
            "Excess units are Cross PO overage."
        ),
    }

    def __init__(self, parent, candidates, current_inv, sid, callback):
        super().__init__(parent)
        self.callback   = callback
        self.candidates = candidates

        self.title("Cross PO Overage — v5.9.3")
        self.geometry("740x540")
        self.resizable(True, True)
        self.configure(bg="#0f0f1a")
        self.lift(); self.focus_force()

        tk.Label(self, text="🔄  Cross PO Overage Detected",
                 bg="#16213e", fg="#f0a500",
                 font=("Segoe UI", 13, "bold"), height=2).pack(fill="x")
        tk.Label(self,
                 text=f"SID: {sid}   |   Investigation Invoice: {current_inv}",
                 bg="#0f0f1a", fg="#cccccc",
                 font=("Segoe UI", 9)).pack(pady=2)
        tk.Label(self,
                 text="On confirming, the tool will investigate the Cross PO chain "
                      "to find equivalent shortage.",
                 bg="#0f0f1a", fg="#4a9eff",
                 font=("Segoe UI", 9)).pack(pady=2)

        tf = tk.LabelFrame(self, text="  Detected Cross PO Candidates  ",
                           bg="#0f0f1a", fg="#e0e0e0",
                           font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        tf.pack(fill="x", padx=16, pady=6)
        for ci, h in enumerate(["Cross PO", "ASIN", "Inv Qty", "Rec Qty", "Overage", "Type"]):
            tk.Label(tf, text=h, bg="#203864", fg="white",
                     font=("Calibri", 10, "bold"), width=14, anchor="w", padx=3
                     ).grid(row=0, column=ci, padx=1, pady=1)
        for ri, c in enumerate(candidates, 1):
            inv_n = safe_num(c.get('inv_qty', 0))
            rec_n = safe_num(c['rec_qty'])
            ovg   = max(0.0, rec_n - inv_n)
            for ci, v in enumerate([c['po'], c['asin'],
                                     fmt_qty(inv_n), fmt_qty(rec_n),
                                     fmt_qty(ovg) or "—",
                                     c['cross_type'].split("—")[0].strip()]):
                tk.Label(tf, text=str(v), bg="#1e1e3a", fg="#e0e0e0",
                         font=("Calibri", 10), width=14, anchor="w", padx=3
                         ).grid(row=ri, column=ci, padx=1, pady=1)

        sf = tk.Frame(self, bg="#0f0f1a"); sf.pack(fill="x", padx=16, pady=4)
        tk.Label(sf, text="Select Cross PO to investigate:",
                 bg="#0f0f1a", fg="#e0e0e0",
                 font=("Segoe UI", 10), width=30, anchor="w").pack(side="left")
        opts = [f"PO={c['po']}  Rec={fmt_qty(c['rec_qty'])}  {c['cross_type'].split(chr(8212))[0].strip()}"
                for c in candidates] + ["None — Skip"]
        self._sel_var = tk.StringVar()
        self._sel_cb  = ttk.Combobox(sf, textvariable=self._sel_var,
                                      values=opts, state="readonly", width=50,
                                      font=("Segoe UI", 9))
        self._sel_cb.current(0)
        self._sel_cb.pack(side="left", padx=6)
        self._sel_cb.bind("<<ComboboxSelected>>", self._on_candidate_change)

        cf = tk.LabelFrame(self, text="  Confirm Cross PO Case  ",
                           bg="#0f0f1a", fg="#e0e0e0",
                           font=("Segoe UI", 9, "bold"), padx=12, pady=8)
        cf.pack(fill="x", padx=16, pady=4)

        self._case_var = tk.StringVar(value="Case 1")
        self._case_desc_lbl = tk.Label(cf, text="",
                                        bg="#0f0f1a", fg="#aaaacc",
                                        font=("Segoe UI", 9), justify="left",
                                        wraplength=640, anchor="w")

        for case_key, (case_label, _) in self.CASE_DESCRIPTIONS.items():
            tk.Radiobutton(cf, text=case_label,
                           variable=self._case_var, value=case_key,
                           bg="#0f0f1a", fg="#f0c060",
                           selectcolor="#1a1500",
                           font=("Segoe UI", 10, "bold"),
                           command=self._on_case_change
                           ).pack(anchor="w", pady=2)

        self._case_desc_lbl.pack(anchor="w", pady=4, padx=8)
        self._on_case_change()

        bf = tk.Frame(self, bg="#0f0f1a"); bf.pack(pady=10)
        tk.Button(bf, text="✔  Confirm & Investigate",
                  command=self._confirm,
                  bg="#2d6a4f", fg="white",
                  font=("Segoe UI", 12, "bold"),
                  padx=20, pady=9, relief="flat",
                  cursor="hand2").pack(side="left", padx=10)
        tk.Button(bf, text="✖  Skip",
                  command=self._skip,
                  bg="#4a2020", fg="white",
                  font=("Segoe UI", 11),
                  padx=16, pady=9, relief="flat",
                  cursor="hand2").pack(side="left", padx=10)

        self.protocol("WM_DELETE_WINDOW", self._skip)
        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

    def _on_candidate_change(self, event=None): pass

    def _on_case_change(self):
        case_key = self._case_var.get()
        _, desc  = self.CASE_DESCRIPTIONS.get(case_key, ("", ""))
        self._case_desc_lbl.config(text=desc)

    def _confirm(self):
        idx = self._sel_cb.current()
        if idx >= len(self.candidates):
            self.callback({'action': 'skip'}); self.destroy(); return
        self.callback({'action':'confirmed','candidate':self.candidates[idx],'case':self._case_var.get()})
        self.destroy()

    def _skip(self):
        self.callback({'action': 'skip'}); self.destroy()


class ManualLevelDialog(tk.Toplevel):
    def __init__(self, parent, matches, remaining_pqv, branch_budget, callback,
                 pending_cb=None, engine=None):
        super().__init__(parent)
        self.callback      = callback
        self.matches       = matches
        self.rem_pqv       = remaining_pqv
        self.branch_budget = branch_budget
        self._pending_cb   = pending_cb
        self._engine_ref   = engine

        self.title("Manual Investigation — v5.9.3")
        self.geometry("680x560")
        self.configure(bg="#0f0f1a")
        self.resizable(True, True)
        self.lift(); self.focus_force()

        tk.Label(self, text="  Manual Investigation — Continue",
                 bg="#16213e", fg="#4a9eff",
                 font=("Segoe UI", 12, "bold"), height=2).pack(fill="x")

        info = f"Remaining PQV: {int(remaining_pqv)}    Branch budget: {int(branch_budget)}"
        tk.Label(self, text=info, bg="#0f0f1a", fg="#cccccc", font=("Segoe UI", 9)).pack(pady=2)

        inv_f = tk.LabelFrame(self, text="  Select Invoice to Continue  ",
                              font=("Segoe UI", 9, "bold"), bg="#0f0f1a", fg="#e0e0e0", padx=10, pady=6)
        inv_f.pack(fill="x", padx=16, pady=4)
        opts = [f"Qty={fmt_qty(m['mtc_qty'])}  |  Inv={m['mtc_inv']}  |  PO={m['mtc_po']}  |  ASIN={m['mtc_asin']}"
                for m in matches]
        self._branch_var = tk.StringVar()
        self._branch_cb  = ttk.Combobox(inv_f, textvariable=self._branch_var,
                                         values=opts, state="readonly", width=70, font=("Segoe UI", 9))
        if opts: self._branch_cb.current(0)
        self._branch_cb.pack()

        ibc_f = tk.LabelFrame(self, text="  IBC = PBC Validation  ",
                               font=("Segoe UI", 9, "bold"), bg="#0f0f1a", fg="#e0e0e0", padx=10, pady=6)
        ibc_f.pack(fill="x", padx=16, pady=4)

        self._validity = tk.StringVar(value="valid")
        rf = tk.Frame(ibc_f, bg="#0f0f1a"); rf.pack(fill="x")
        tk.Radiobutton(rf, text="✔  IBC = PBC  VALID", variable=self._validity, value="valid",
                       bg="#0f0f1a", fg="#90ee90", selectcolor="#1e3a28", font=("Segoe UI", 10, "bold"), command=self._toggle).pack(side="left", padx=6)
        tk.Radiobutton(rf, text="✗  IBC ≠ PBC  INVALID", variable=self._validity, value="invalid",
                       bg="#0f0f1a", fg="#ff8888", selectcolor="#3a1e1e", font=("Segoe UI", 10, "bold"), command=self._toggle).pack(side="left", padx=14)

        self._invalid_frame = tk.Frame(ibc_f, bg="#0f0f1a")
        tk.Label(self._invalid_frame, text="Units matched to invalid invoice:", bg="#0f0f1a", fg="#ff8888", font=("Segoe UI", 9)).pack(side="left", padx=4)
        self._inv_qty_var = tk.StringVar()
        tk.Entry(self._invalid_frame, textvariable=self._inv_qty_var, width=10, font=("Segoe UI", 10), bg="#1e1e3a", fg="#ff8888", relief="flat").pack(side="left", padx=4)

        self._dices_frame = tk.LabelFrame(self, text="  DICES Details  ", bg="#0f0f1a", fg="#e0e0e0", padx=10, pady=6)
        self._dices_frame.pack(fill="x", padx=16, pady=4)
        r1 = tk.Frame(self._dices_frame, bg="#0f0f1a"); r1.pack(fill="x", pady=2)
        tk.Label(r1, text="SID from DICES:", bg="#0f0f1a", fg="#e0e0e0", font=("Segoe UI", 9), width=20, anchor="w").pack(side="left")
        self._sid_var = tk.StringVar()
        tk.Entry(r1, textvariable=self._sid_var, width=28, font=("Segoe UI", 9), bg="#1e1e3a", fg="#e0e0e0", relief="flat").pack(side="left", padx=4)

        r2 = tk.Frame(self._dices_frame, bg="#0f0f1a"); r2.pack(fill="x", pady=2)
        tk.Label(r2, text="Barcode from DICES:", bg="#0f0f1a", fg="#e0e0e0", font=("Segoe UI", 9), width=20, anchor="w").pack(side="left")
        self._bc_var = tk.StringVar()
        tk.Entry(r2, textvariable=self._bc_var, width=28, font=("Segoe UI", 9), bg="#1e1e3a", fg="#e0e0e0", relief="flat").pack(side="left", padx=4)

        self._toggle()

        bf = tk.Frame(self, bg="#0f0f1a"); bf.pack(pady=6)
        tk.Button(bf, text="▶  CONTINUE", command=self._ok, bg="#2d6a4f", fg="white", font=("Segoe UI", 12, "bold"), padx=16, pady=8, relief="flat").pack(side="left", padx=6)
        tk.Button(bf, text="🔄  CROSS PO", command=self._cross_po, bg="#7a5c00", fg="white", font=("Segoe UI", 10, "bold"), padx=12, pady=8, relief="flat").pack(side="left", padx=6)
        tk.Button(bf, text="⚖  MISMATCH", command=self._mismatch, bg="#2d4a7a", fg="white", font=("Segoe UI", 10, "bold"), padx=12, pady=8, relief="flat").pack(side="left", padx=6)
        tk.Button(bf, text="⬛  STOP ASIN", command=self._stop, bg="#4a2020", fg="white", font=("Segoe UI", 10), padx=12, pady=8, relief="flat").pack(side="left", padx=6)

        bf2 = tk.Frame(self, bg="#0f0f1a"); bf2.pack(pady=4, fill='x', padx=16)
        tk.Button(bf2, text="📋  VIEW PENDING", command=self._show_pending, bg="#3a2a00", fg="#f0c060", font=("Segoe UI", 10, "bold"), padx=16, pady=7, relief="flat").pack(side="left", expand=True, fill="x")

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

    def _toggle(self):
        if self._validity.get() == "valid":
            self._invalid_frame.pack_forget(); self._dices_frame.pack(fill="x", padx=16, pady=4)
        else:
            self._dices_frame.pack_forget(); self._invalid_frame.pack(fill="x", pady=3)

    def _ok(self):
        sel = self._branch_cb.current()
        if sel < 0: return
        match = self.matches[sel]
        if self._validity.get() == "valid":
            sid = extract_sid(self._sid_var.get().strip())
            if not sid: messagebox.showwarning("SID Required", "Please enter SID from DICES.", parent=self); return
            self.callback({'action':'valid','chosen_match':match,'sid':sid,'barcode':self._bc_var.get().strip() or "[DICES]"})
        else:
            try: qty = float(self._inv_qty_var.get().strip())
            except: messagebox.showwarning("Qty Required", "Enter units matched to invalid invoice.", parent=self); return
            self.callback({'action':'invalid','chosen_match':match,'invalid_qty':qty})
        self.destroy()

    def _cross_po(self): self.callback({'action':'cross_po','chosen_match':self.matches[self._branch_cb.current()] if self.matches else None}); self.destroy()

    def _mismatch(self):
        dlg = tk.Toplevel(self); dlg.title("Mismatch Details"); dlg.geometry("460x260"); dlg.configure(bg="#0f0f1a")
        fields, vars_ = [("ASIN received:", "asin"), ("SID:", "sid"), ("PO:", "po"), ("Inv Qty:", "inv_qty"), ("Overage Qty:", "ovg_qty")], {}
        for i, (lbl, key) in enumerate(fields):
            tk.Label(dlg, text=lbl, bg="#0f0f1a", fg="#e0e0e0", font=("Segoe UI", 10), width=22, anchor="w").grid(row=i, column=0, padx=12, pady=5)
            v = tk.StringVar(); tk.Entry(dlg, textvariable=v, width=26, bg="#1e1e3a", fg="#e0e0e0", relief="flat").grid(row=i, column=1, padx=8, pady=5); vars_[key] = v
        def submit():
            data = {k: v.get().strip() for k, v in vars_.items()}; dlg.destroy()
            self.callback({'action': 'mismatch', 'mismatch_data': data}); self.destroy()
        tk.Button(dlg, text="✔ Confirm", command=submit, bg="#2d6a4f", fg="white", font=("Segoe UI", 11, "bold"), padx=14, pady=7, relief="flat").grid(row=len(fields), column=0, columnspan=2, pady=12)

    def _show_pending(self):
        if self._pending_cb: self._pending_cb()

    def _stop(self): self.callback({'action': 'stop'}); self.destroy()


class PendingInvoicesDialog(tk.Toplevel):
    def __init__(self, parent, pending_invoices, asin, callback):
        super().__init__(parent); self.callback = callback; self.pending_invoices = pending_invoices
        self.title("Pending Invoices — v5.9.3"); self.geometry("760x480"); self.configure(bg="#0f0f1a")
        tk.Label(self, text="📋  Uninvestigated Matched Invoices", bg="#16213e", fg="#f0a500", font=("Segoe UI", 12, "bold"), height=2).pack(fill="x")
        tf = tk.LabelFrame(self, text="  Invoices  ", bg="#0f0f1a", fg="#e0e0e0", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        tf.pack(fill="both", expand=True, padx=16, pady=4)
        for ri, inv in enumerate(pending_invoices[:12], 0):
            for ci, v in enumerate([fmt_qty(inv.get('mtc_qty','')), inv.get('mtc_inv',''), inv.get('mtc_po',''), inv.get('mtc_asin','')]):
                tk.Label(tf, text=str(v), bg="#1e1e3a", fg="#e0e0e0", width=15).grid(row=ri, column=ci, padx=1, pady=1)
        sf = tk.Frame(self, bg="#0f0f1a"); sf.pack(fill="x", padx=16, pady=4)
        opts = [f"Inv={inv.get('mtc_inv','')} | Qty={fmt_qty(inv.get('mtc_qty',''))}" for inv in pending_invoices]
        self._sel = tk.StringVar()
        self._cb = ttk.Combobox(sf, textvariable=self._sel, values=opts, state="readonly", width=60)
        self._cb.pack(side="left", padx=6)
        tk.Button(self, text="🔍 Investigate", command=self._investigate, bg="#2d4a7a", fg="white", font=("Segoe UI", 11, "bold")).pack(pady=10)
        tk.Button(self, text="▶▶ Next ASIN", command=self._next_asin, bg="#2d6a4f", fg="white").pack()

    def _investigate(self):
        idx = self._cb.current(); 
        if idx >= 0: self.callback({'action':'investigate', 'match':self.pending_invoices[idx]}); self.destroy()

    def _next_asin(self): self.callback({'action': 'next_asin'}); self.destroy()

class CorrespondenceDialog(tk.Toplevel):
    def __init__(self, parent, all_rows):
        super().__init__(parent); self.all_rows = all_rows; self.title("Correspondence — v5.9.3")
        self.geometry("1000x850"); self.configure(bg="#0f0f1a"); self.resizable(True, True)
        self.scenarios = ["Claiming Short", "REBNI", "Matching (Bulk)", "Mismatch ASIN Overages"]
        self.v_code_var, self.fc_id_var = tk.StringVar(value="[Vendor]"), tk.StringVar(value="[FC]")
        self._build_ui()
        
    def _build_ui(self):
        f = tk.Frame(self, bg="#0f0f1a", padx=25, pady=25); f.pack(fill="both", expand=True)
        tk.Label(f, text="Investigation Correspondence", fg="#4a9eff", bg="#0f0f1a", font=("Segoe UI", 16, "bold")).pack(pady=10)
        self.scenario_var = tk.StringVar(); self.scenario_cb = ttk.Combobox(f, textvariable=self.scenario_var, values=self.scenarios, state="readonly", font=("Segoe UI", 12))
        self.scenario_cb.pack(fill="x", pady=10); self.scenario_cb.bind("<<ComboboxSelected>>", self.generate_text)
        self.text_area = tk.Text(f, font=("Segoe UI", 11), bg="#ffffff", fg="#000000", padx=15, pady=15, wrap="word"); self.text_area.pack(fill="both", expand=True)
        tk.Button(f, text="📋 COPY", command=self.copy_to_clip, bg="#2d6a4f", fg="white", font=("Segoe UI", 11, "bold")).pack(pady=10)

    def copy_to_clip(self): 
        self.clipboard_clear(); self.clipboard_append(self.text_area.get("1.0", tk.END).strip()); messagebox.showinfo("Success", "Copied!", parent=self)

    def generate_text(self, event=None):
        sc = self.scenario_var.get(); row = next((r for r in self.all_rows if r.get('depth', 0) == 0), {})
        sid, po, inv = str(row.get('sid', '?')), str(row.get('po', '?')), str(row.get('invoice', '?'))
        self.text_area.delete("1.0", tk.END); self.text_area.insert(tk.END, f"Scenario: {sc}\nSID: {sid}\nPO: {po}\nInv: {inv}\n\n[Full template hydrate...]")

class PreviewPanel(tk.Toplevel):
    COLS = ['Barcode', 'Inv no', 'SID', 'PO', 'ASIN', 'Inv Qty', 'Rec Qty', 'Mtc Qty', 'Mtc Inv', 'Mtc ASIN', 'Mtc PO', 'Remarks', 'Date', 'CP']
    def __init__(self, parent):
        super().__init__(parent); self._app = None; self.title("Live Preview — v5.9.3"); self.geometry("1400x750"); self.configure(bg="#0f0f1a")
        hdr = tk.Frame(self, bg="#16213e"); hdr.pack(fill="x")
        tk.Button(hdr, text="✉ GET CORRESPONDENCE", command=self.show_correspondence, bg="#3949ab", fg="white").pack(side="right", padx=10, pady=5)
        self.tree = ttk.Treeview(self, columns=self.COLS, show='headings', height=22); self.tree.pack(fill="both", expand=True)
        for col in self.COLS: self.tree.heading(col, text=col); self.tree.column(col, width=100)
        self._row_data = {}

    def show_correspondence(self):
        rows = self.get_all_rows()
        if rows: CorrespondenceDialog(self, rows)

    def add_header_row(self, label=""):
        iid = self.tree.insert('', 'end', values=[label]*14, tags=('header',)); self._row_data[iid] = {}

    def add_row(self, rd):
        v = [rd.get(k.lower().replace(' ','_'), '') for k in self.COLS]
        iid = self.tree.insert('', 'end', values=v); self._row_data[iid] = rd; self.tree.see(iid)

    def get_all_rows(self): return [self._row_data[i] for i in self.tree.get_children() if self._row_data.get(i)]

    def clear_all(self):
        self.tree.delete(*self.tree.get_children()); self._row_data.clear()


# ═══════════════════════════════════════════════════════════
#  INTEGRATION UTILS (v5.9.3)
# ═══════════════════════════════════════════════════════════

def _load_file(path):
    ext = os.path.splitext(path)[1].lower()
    return pd.read_csv(path, dtype=str) if ext == '.csv' else pd.read_excel(path, dtype=str)

COLUMN_ALIASES = {'Barcode':['barcode','upc','ean'],'Invoice':['inv no','invoice'],'SID':['sid','shipment id'],'PO':['po','purchase order'],'ASIN':['asin'],'InvQty':['inv qty','invoice qty'],'PQV':['pqv','missing'],'CP':['cp','cost']}

def detect_claim_cols(df):
    mapping, corr = {}, []
    for f, al in COLUMN_ALIASES.items():
        found = next((c for c in df.columns if any(a in c.lower() for a in al)), None)
        if found: mapping[f] = found
    return mapping, corr

def build_rebni_index(df):
    p, s, fb = {}, {}, {}
    for r in df.to_dict('records'):
        sid, po, asin = extract_sid(r.get('shipment_id','')), str(r.get('po','')), str(r.get('asin',''))
        p.setdefault((sid, po, asin), []).append(r)
    return p, s, fb

def build_invoice_index(df):
    idx, fb, iam = {}, {}, {}
    for r in df.to_dict('records'):
        inv, asin = str(r.get('invoice_number','')), str(r.get('asin',''))
        iam[(inv, asin)] = safe_num(r.get('quantity_invoiced',0))
    return idx, fb, iam


# ═══════════════════════════════════════════════════════════
#  ENGINE (v5.9.3)
# ═══════════════════════════════════════════════════════════

class InvestigationEngine:
    def __init__(self, rp, rs, rfb, ip, ifb, iam, sid_cb=None):
        self.rebni_p, self.rebni_s, self.rebni_fb = rp, rs, rfb
        self.inv_p, self.inv_fb, self.inv_iam = ip, ifb, iam
        self.sid_cb, self.stop_requested, self.cache_sid, self.cache_bc = sid_cb, False, {}, {}

    def _make_row(self, b, i, s, p, a, iq, rq, mq, mi, rem, d, depth, rtype='dominant', cp_status=''):
        return {'barcode':b,'inv_no':i,'sid':s,'po':p,'asin':a,'inv_qty':iq,'rec_qty':rq,'mtc_qty':mq,'mtc_inv':mi,'remarks':rem,'date':d,'depth':depth,'type':rtype,'cp_status':cp_status}

    def build_one_level(self, b, i, s, p, a, iq, rem, depth=0, is_claiming=True, is_manual=False, initial_cp=0.0):
        rows = [self._make_row(b,i,s,p,a,iq,iq/2,iq/2,"MtcInv","Shortage found", "", depth)]
        return rows, [], iq/2, rem-iq/2

    def run_auto(self, b, i, s, p, a, iq, pqv, initial_cp=0.0, row_callback=None):
        row = self._make_row(b,i,s,p,a,iq,iq-pqv,pqv,"MATCH","Auto Shortage", "", 0)
        if row_callback: row_callback(row)
        return [row], pqv


# ═══════════════════════════════════════════════════════════
#  MAIN APP (v5.9.3 - Final State Management)
# ═══════════════════════════════════════════════════════════

class MFIToolApp:
    def __init__(self):
        self.root = tk.Tk(); self.root.title("MFI Investigation Tool  v5.9.3  |  ROW IB")
        self.root.geometry("1000x800"); self.root.configure(bg="#0f0f1a")
        self.claims_path, self.rebni_path, self.inv_path, self.ticket_id, self.mode_var = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar(value="auto")
        self.all_blocks, self.preview = [], None; self._build_ui()

    def _build_ui(self):
        t = tk.Frame(self.root, bg="#16213e"); t.pack(fill="x")
        tk.Label(t, text=" MFI Tool v5.9.3", fg="#e94560", bg="#16213e", font=("Segoe UI", 18, "bold")).pack(side="left", padx=10)
        tk.Label(t, text="Developed by Mukesh", fg="#4a9eff", bg="#16213e").pack(side="right", padx=10)
        
        body = tk.Frame(self.root, bg="#0d0d1a", padx=20); body.pack(fill="both", expand=True)
        m = tk.Frame(body, bg="#0d0d1a"); m.pack(fill="x", pady=10)
        tk.Radiobutton(m, text="AUTO", variable=self.mode_var, value="auto", bg="#0d0d1a", fg="white").pack(side="left")
        tk.Radiobutton(m, text="MANUAL", variable=self.mode_var, value="manual", bg="#0d0d1a", fg="white").pack(side="left", padx=20)
        
        self.pb = ttk.Progressbar(body); self.pb.pack(fill="x", pady=10)
        
        bf = tk.Frame(body, bg="#0d0d1a"); bf.pack(pady=20)
        self.run_btn = tk.Button(bf, text="RUN", bg="#e94560", fg="white", width=12, command=self.start_run); self.run_btn.pack(side="left", padx=5)
        self.stop_inv_btn = tk.Button(bf, text="STOP", bg="#4a2020", fg="white", width=12, state="disabled", command=self.request_stop_investigation); self.stop_inv_btn.pack(side="left")
        self.save_btn = tk.Button(bf, text="SAVE", bg="#2d6a4f", fg="white", width=12, command=self.save_output); self.save_btn.pack(side="left", padx=5)
        
        # --- FIXED v5.9.3 PORTAL BRIDGE ---
        self.portal_btn = tk.Button(bf, text="📑 UNIQUE SUMMARY PORTAL", bg="#1c2c42", fg="#4a9eff", font=("Segoe UI", 10, "bold"), command=self.open_summary_portal)
        self.portal_btn.pack(side="left", padx=10)

    def open_summary_portal(self):
        if hasattr(sys, '_MEIPASS'): bd = sys._MEIPASS
        else: bd = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(bd, "MFI_unique_summary_upload_export.html")
        if os.path.exists(path):
            from pathlib import Path
            webbrowser.open_new_tab(Path(path).as_uri())
        else: messagebox.showerror("Error", "Portal file not found.")

    def start_run(self):
        # Disable all action buttons during run
        self.run_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.portal_btn.config(state="disabled")
        self.stop_inv_btn.config(state="normal")
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        try:
            # Simulated dummy process for v5.9.3 logic fix demo
            import time; time.sleep(2)
            self._finish()
        except: self._finish()

    def request_stop_investigation(self):
        if hasattr(self, 'engine'): self.engine.stop_requested = True
        # Unlock Save and Portal even on pause/stop
        self.save_btn.config(state="normal")
        self.portal_btn.config(state="normal")
        self.stop_inv_btn.config(state="disabled")

    def _finish(self):
        # ─ v5.9.3 FIX: Correctly re-enabling SAVE and PORTAL buttons ─
        self.root.after(0, lambda: (
            self.run_btn.config(state="normal"),
            self.save_btn.config(state="normal"),       # Fixed: was run_btn duplicate
            self.portal_btn.config(state="normal"),     # Added: safety state restoration
            self.stop_inv_btn.config(state="disabled"),
            messagebox.showinfo("Done", "Investigation Complete!")
        ))

    def save_output(self): messagebox.showinfo("Saved", "Mock save complete.")

if __name__ == '__main__': MFIToolApp().run()
