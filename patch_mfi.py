import re

with open('mfi_tool.py', 'r', encoding='utf-8') as f:
    code = f.read()

# SECTION 3 - DEDUPLICATION FIX (and preparing for Change 1 & 7)
# Let's replace build_rebni_index
code = re.sub(
    r'def build_rebni_index\(df\):.*?return primary, secondary',
    r'''def build_rebni_index(df):
    primary = {}
    secondary = {}
    fallback = {}
    for _, row in df.iterrows():
        sid_frag = extract_sid(row['shipment_id'])
        po = clean(row['po'])
        asin = clean(row['asin'])
        pkey = (sid_frag, po, asin)
        primary.setdefault(pkey, []).append(row.to_dict())
        skey = (po, asin)
        secondary.setdefault(skey, []).append(row.to_dict())
        for inv in split_comma(row['matched_invoice_numbers']):
            if inv: fallback.setdefault((sid_frag, po, inv), []).append(row.to_dict())
    return primary, secondary, fallback''',
    code, flags=re.DOTALL
)

# Replace build_invoice_index
code = re.sub(
    r'def build_invoice_index\(df\):.*?return pos_dict',
    r'''def build_invoice_index(df):
    pos_dict = {}
    fallback = {}
    for _, row in df.iterrows():
        sids = split_comma(row['shipment_id'])
        pos = split_comma(row['matched_po'])
        asins = split_comma(row['matched_asin'])
        qtys = split_comma(row['shipmentwise_matched_qty'])
        max_len = max(len(sids), len(pos), len(asins), len(qtys))
        for p in range(max_len):
            s_frag = extract_sid(sids[p] if p < len(sids) else "")
            p_val = pos[p] if p < len(pos) else ""
            a_val = asins[p] if p < len(asins) else ""
            q_val = safe_num(qtys[p] if p < len(qtys) else "0")
            key = (s_frag, p_val, a_val)
            inv_no = clean(row['invoice_number'])
            mtc_po = clean(row['purchase_order_id'])
            mtc_asin = clean(row['asin'])
            entry = {
                'mtc_inv': inv_no, 'mtc_po': mtc_po, 'mtc_asin': mtc_asin,
                'inv_qty': safe_num(row['quantity_invoiced']), 'mtc_qty': q_val, 'date': clean(row['invoice_date'])
            }
            if key not in pos_dict: pos_dict[key] = []
            if not any(d['mtc_inv'] == inv_no and d['mtc_po'] == mtc_po and d['mtc_asin'] == mtc_asin for d in pos_dict[key]):
                pos_dict[key].append(entry)
            
            if inv_no:
                f_key = (s_frag, p_val, inv_no)
                if f_key not in fallback: fallback[f_key] = []
                if not any(d['mtc_inv'] == inv_no and d['mtc_po'] == mtc_po and d['mtc_asin'] == mtc_asin for d in fallback[f_key]):
                    fallback[f_key].append(entry)
    return pos_dict, fallback''',
    code, flags=re.DOTALL
)

# Update InvestigationEngine init & lookup methods
engine_old = '''class InvestigationEngine:
    def __init__(self, rebni_primary, rebni_secondary, inv_index, sid_request_callback=None):
        self.rebni_p = rebni_primary
        self.rebni_s = rebni_secondary
        self.inv = inv_index
        self.sid_cb = sid_request_callback
        self.max_depth = 10'''

engine_new = '''class InvestigationEngine:
    def __init__(self, rebni_primary, rebni_secondary, rebni_fallback, inv_primary, inv_fallback, sid_request_callback=None):
        self.rebni_p  = rebni_primary
        self.rebni_s  = rebni_secondary
        self.rebni_fb = rebni_fallback
        self.inv_p    = inv_primary
        self.inv_fb   = inv_fallback
        self.sid_cb   = sid_request_callback
        self.max_depth = 10

    def _rebni_lookup(self, sid, po, asin, invoice_no=None):
        rows = self.rebni_p.get((sid, po, asin), [])
        if not rows and invoice_no:
            rows = self.rebni_fb.get((sid, po, invoice_no), [])
        return rows

    def _inv_lookup(self, sid, po, asin, invoice_no=None):
        matches = self.inv_p.get((sid, po, asin), [])
        if not matches and invoice_no:
            matches = self.inv_fb.get((sid, po, invoice_no), [])
        return matches'''
code = code.replace(engine_old, engine_new)


# Update run_auto
run_auto_old_regex = r'def run_auto\(self, barcode, inv_no, sid, po, asin, inv_qty, pqv, depth=0, visited=None, rem_pqv=None\):.*?return rows, current_rem'
run_auto_new = '''def run_auto(self, barcode, inv_no, sid, po, asin, inv_qty, pqv, depth=0, visited=None, rem_pqv=None, is_claiming=True):
        if visited is None: visited = set()
        if rem_pqv is None: rem_pqv = safe_num(pqv)
        sid_frag = extract_sid(sid)
        state = (sid_frag, clean(inv_no), clean(po), clean(asin))
        if state in visited or depth >= self.max_depth: 
            return [], rem_pqv
        visited.add(state)
        
        rows, matches, rec_qty, current_rem = self._build_level(barcode, inv_no, sid, po, asin, inv_qty, rem_pqv, depth, is_claiming)
        
        rem_str = rows[0].get('remarks', '')
        if "Root cause found" in rem_str or "REBNI" in rem_str or "SID not found" in rem_str or "SR" in rem_str:
            return rows, current_rem
            
        for match in matches:
            if current_rem <= 0: break
            next_inv = match['mtc_inv']
            
            # MISTAKE 5: Self stops investigation completely
            # RIGHT: if next_inv == invoice: continue (skip recursion, keep looping)
            if next_inv == clean(inv_no): continue
            
            next_sid = self._find_sid(match['mtc_po'], match['mtc_asin'], next_inv)
            if not next_sid:
                if self.sid_cb:
                    next_sid = self.sid_cb(next_inv, match['mtc_po'], match['mtc_asin'])
                if not next_sid:
                    d_row = self._make_row("[DICES]", next_inv, "[DICES]", match['mtc_po'], match['mtc_asin'], 
                                           match['inv_qty'], 0.0, "SID not found in REBNI — follow manually", match['date'], depth + 1)
                    d_row['mtc_qty'] = match['mtc_qty']
                    d_row['mtc_inv'] = next_inv
                    rows.append(d_row)
                    continue

            child_rows, current_rem = self.run_auto("[DICES]", next_inv, next_sid, match['mtc_po'], match['mtc_asin'], 
                                                   match['inv_qty'], pqv, depth + 1, visited, current_rem, is_claiming=False)
            rows.extend(child_rows)
        return rows, current_rem'''
code = re.sub(run_auto_old_regex, run_auto_new, code, flags=re.DOTALL)


# Update build_one_level
code = code.replace(
    'def build_one_level(self, barcode, inv_no, sid, po, asin, inv_qty, rem_pqv, depth=0):',
    'def build_one_level(self, barcode, inv_no, sid, po, asin, inv_qty, rem_pqv, depth=0, is_claiming=True):'
)
code = code.replace(
    'return self._build_level(barcode, inv_no, sid, po, asin, inv_qty, rem_pqv, depth)',
    'return self._build_level(barcode, inv_no, sid, po, asin, inv_qty, rem_pqv, depth, is_claiming)'
)

# Update _build_level
build_lvl_old = r'def _build_level\(self, barcode, inv_no, sid, po, asin, inv_qty, rem_pqv, depth\):.*?return rows, all_matches_sorted, rec_qty, new_rem_pqv'
build_lvl_new = '''def _build_level(self, barcode, inv_no, sid, po, asin, inv_qty, rem_pqv, depth, is_claiming):
        sid_frag = extract_sid(sid)
        rebni_rows = self._rebni_lookup(sid_frag, clean(po), clean(asin), clean(inv_no))
        rec_qty, rebni_avail, remarks, rec_date = 0.0, 0.0, "", ""
        if rebni_rows:
            r = rebni_rows[0]
            rec_qty = safe_num(r['quantity_unpacked'])
            rebni_avail = safe_num(r['rebni_available'])
            rec_date = clean(r['received_datetime'])
            if rebni_avail > 0:
                level = 'claiming shipment' if is_claiming else 'matching shipment'
                remarks = f"REBNI Available = {int(rebni_avail)} units at {level} level — Suggest TSP to utilize"
        else:
            remarks = "SID not found in REBNI — check DICES manually"
        if not rebni_rows and rec_qty == 0 and depth > 0: remarks = "SR"

        shortage = safe_num(inv_qty) - rec_qty
        if shortage >= rem_pqv and rem_pqv > 0:
            remarks = f"Short {fmt_qty(shortage)} units — Root cause found (PQV={fmt_qty(rem_pqv)})"
            new_rem_pqv = 0
        else:
            new_rem_pqv = rem_pqv - (shortage if shortage > 0 else 0)

        raw_matches = self._inv_lookup(sid_frag, clean(po), clean(asin), clean(inv_no))
        
        seen = set()
        unique_matches = []
        for m in raw_matches:
            combo = (m['mtc_inv'], m['mtc_po'], m['mtc_asin'])
            if combo not in seen:
                seen.add(combo)
                unique_matches.append(m)
        all_matches_sorted = sorted(unique_matches, key=lambda x: x['mtc_qty'], reverse=True)
        
        # Self check
        is_self_only = len(all_matches_sorted) == 1 and all_matches_sorted[0]['mtc_inv'] == clean(inv_no)
        if is_self_only:
             if not remarks or "SID not found" in remarks:
                 remarks = "Self — matched within same invoice"

        rows = []
        main_row = self._make_row(barcode, inv_no, sid, po, asin, inv_qty, rec_qty, remarks, rec_date, depth)
        if all_matches_sorted:
            main_row['mtc_qty'] = all_matches_sorted[0]['mtc_qty']
            main_row['mtc_inv'] = all_matches_sorted[0]['mtc_inv']
            rows.append(main_row)
            for i in range(1, len(all_matches_sorted)):
                rows.append(self._make_row("", "", "", "", "", "", "", "", "", depth, 'subrow', all_matches_sorted[i]))
        else:
            rows.append(main_row)
        return rows, all_matches_sorted, rec_qty, new_rem_pqv'''
code = re.sub(build_lvl_old, build_lvl_new, code, flags=re.DOTALL)


# Update write_excel
excel_old = r'def write_excel\(all_blocks, out_path\):.*?wb\.save\(out_path\)'
excel_new = '''def write_excel(all_blocks, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Investigation"
    headers = ["Barcode", "Inv no", "SID", "PO", "ASIN", "Inv Qty", "Rec Qty", "Mtc Qty", "Mtc Inv", "Remarks", "Date"]
    h_fill = PatternFill(start_color="203864", end_color="203864", fill_type="solid")
    h_font = Font(color="FFFFFF", bold=True)
    dom_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    sub_fill = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")
    root_fill = PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid")
    root_font = Font(color="9C0006", bold=True)
    dices_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    sr_fill = PatternFill(start_color="FFD7D7", end_color="FFD7D7", fill_type="solid")
    sr_font = Font(color="CC0000", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    INVLD_FILL = PatternFill("solid", fgColor="FFD0D0")
    REBNI_FILL = PatternFill("solid", fgColor="D0F0FF")
    INVLD_FONT = Font(bold=True, color="880000", name="Calibri", size=10, italic=True)
    REBNI_FONT = Font(bold=True, color="005580", name="Calibri", size=10)
    
    KEY_MAP = {'Barcode':'barcode','Inv no':'invoice','SID':'sid','PO':'po','ASIN':'asin','Inv Qty':'inv_qty',
               'Rec Qty':'rec_qty','Mtc Qty':'mtc_qty','Mtc Inv':'mtc_inv','Remarks':'remarks','Date':'date'}
    
    curr = 1
    for block in all_blocks:
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=curr, column=c, value=h)
            cell.fill, cell.font, cell.border = h_fill, h_font, border
        curr += 1
        for r_data in block:
            for c, h in enumerate(headers, 1):
                val = r_data.get(KEY_MAP[h], "")
                cell = ws.cell(row=curr, column=c, value=val)
                cell.border = border
                
                rem = str(r_data.get('remarks', ''))
                is_invld = 'invalid invoice' in rem.lower()
                is_rebni = 'REBNI Available' in rem
                
                if r_data['type'] == 'subrow': cell.fill = sub_fill
                elif r_data['type'] == 'dominant' and r_data['depth'] > 0: cell.fill = dom_fill
                
                if is_invld: cell.fill, cell.font = INVLD_FILL, INVLD_FONT
                elif is_rebni: cell.fill, cell.font = REBNI_FILL, REBNI_FONT
                elif "Root cause" in rem: cell.fill, cell.font = root_fill, root_font
                elif "SR" == rem: cell.fill, cell.font = sr_fill, sr_font
                
                if "[DICES]" in str(r_data.get('barcode', '')): 
                    if c <= 3: cell.fill = dices_fill
            curr += 1
        curr += 1
    widths = [18, 22, 18, 12, 14, 9, 9, 9, 26, 36, 22]
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(out_path)'''
code = re.sub(excel_old, excel_new, code, flags=re.DOTALL)


# Update GUI CLASS
# 1. Title bar, window config
code = code.replace('self.root.title("MFI Investigation Tool  v4.1  |  ROW IB")', 'self.root.title("MFI Investigation Tool  v4.2  |  ROW IB")')
code = code.replace('self.root.geometry("800x680")', 'try: self.root.state("zoomed")\n        except: self.root.attributes("-zoomed", True)\n        self.root.minsize(900, 580)')
code = code.replace('self.root.resizable(False, False)', 'self.root.resizable(True, True)')

# 2. Add ManualLevelDialog definition before MFIToolApp
manual_dlg = '''class ManualLevelDialog(tk.Toplevel):
    def __init__(self, parent, matches, remaining_pqv):
        super().__init__(parent)
        self.result = None
        self.matches = matches
        self.rem_pqv = remaining_pqv
        
        self.title("Manual Investigation — Next Step")
        self.geometry("620x420")
        self.configure(bg="#0f0f1a")
        self.resizable(False, False)
        self.grab_set()
        self.focus_set()
        
        tk.Label(self, text="Select Invoice to Continue", bg="#0f0f1a", fg="#4a9eff", font=("Segoe UI", 12, "bold")).pack(pady=(15, 5))
        
        self.combo = ttk.Combobox(self, width=70, state="readonly", font=("Segoe UI", 10))
        opts = [f"Qty={m['mtc_qty']}  |  Inv={m['mtc_inv']}  |  PO={m['mtc_po']}  |  ASIN={m['mtc_asin']}" for m in matches]
        self.combo['values'] = opts
        if opts: self.combo.current(0)
        self.combo.pack(pady=5)
        
        tk.Label(self, text="IBC = PBC Validation", bg="#0f0f1a", fg="#e0e0e0", font=("Segoe UI", 11, "bold")).pack(pady=(20, 10))
        
        self.val_var = tk.StringVar(value="valid")
        
        rf = tk.Frame(self, bg="#0f0f1a")
        rf.pack(fill="x", padx=40)
        
        tk.Radiobutton(rf, text="✔  IBC = PBC  VALID  — Continue investigation", variable=self.val_var, value="valid", 
                       bg="#0f0f1a", fg="#2d6a4f", font=("Segoe UI", 10, "bold"), selectcolor="#16213e", command=self._toggle).pack(anchor="w", pady=5)
        tk.Radiobutton(rf, text="✗  IBC ≠ PBC  INVALID — Exclude units", variable=self.val_var, value="invalid", 
                       bg="#0f0f1a", fg="#e94560", font=("Segoe UI", 10, "bold"), selectcolor="#16213e", command=self._toggle).pack(anchor="w", pady=5)
                       
        self.dyn_frame = tk.Frame(self, bg="#1a1a2e", padx=15, pady=15, highlightbackground="#4a9eff", highlightthickness=1)
        self.dyn_frame.pack(fill="x", padx=40, pady=15)
        
        self.v_frame = tk.Frame(self.dyn_frame, bg="#1a1a2e")
        tk.Label(self.v_frame, text="SID from DICES:", bg="#1a1a2e", fg="white").grid(row=0, column=0, sticky="w", pady=5)
        self.sid_ent = tk.Entry(self.v_frame, width=25)
        self.sid_ent.grid(row=0, column=1, padx=10, pady=5)
        tk.Label(self.v_frame, text="Barcode from DICES:", bg="#1a1a2e", fg="white").grid(row=1, column=0, sticky="w", pady=5)
        self.bc_ent = tk.Entry(self.v_frame, width=25)
        self.bc_ent.grid(row=1, column=1, padx=10, pady=5)
        
        self.inv_frame = tk.Frame(self.dyn_frame, bg="#1a1a2e")
        tk.Label(self.inv_frame, text="Units matched to invalid invoice:", bg="#1a1a2e", fg="white").grid(row=0, column=0, sticky="w", pady=5)
        self.qty_ent = tk.Entry(self.inv_frame, width=15)
        self.qty_ent.grid(row=0, column=1, padx=10, pady=5)
        
        self._toggle()
        
        bf = tk.Frame(self, bg="#0f0f1a")
        bf.pack(pady=20)
        tk.Button(bf, text="▶  CONTINUE", bg="#2d6a4f", fg="white", font=("Segoe UI", 10, "bold"), width=15, command=self._continue).pack(side="left", padx=10)
        tk.Button(bf, text="⬛  STOP THIS ASIN", bg="#4a2020", fg="white", font=("Segoe UI", 10, "bold"), width=18, command=self._stop).pack(side="left", padx=10)
        
        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")
        
    def _toggle(self):
        if self.val_var.get() == "valid":
            self.inv_frame.pack_forget()
            self.v_frame.pack(fill="x")
        else:
            self.v_frame.pack_forget()
            self.inv_frame.pack(fill="x")
            
    def _continue(self):
        idx = self.combo.current()
        if idx < 0: return
        match = self.matches[idx]
        
        if self.val_var.get() == "valid":
            sid = self.sid_ent.get().strip()
            bc = self.bc_ent.get().strip()
            if not sid:
                messagebox.showerror("Error", "SID field cannot be empty.")
                return
            self.result = {'action':'valid', 'chosen_match':match, 'sid':sid, 'barcode':bc, 'invalid_qty':0}
        else:
            qty_str = self.qty_ent.get().strip()
            try: qty = float(qty_str)
            except:
                messagebox.showerror("Error", "Units must be a valid number.")
                return
            self.result = {'action':'invalid', 'chosen_match':match, 'invalid_qty':qty, 'sid':'', 'barcode':''}
            
        self.destroy()
        
    def _stop(self):
        self.result = {'action':'stop'}
        self.destroy()

class MFIToolApp:'''
code = code.replace('class MFIToolApp:', manual_dlg)

# 3. GUI updates
code = code.replace('font=("Segoe UI", 17, "bold")', 'font=("Segoe UI", 20, "bold")')
code = code.replace('text="v4.1  |  ROW IB"', 'text="v4.2  |  ROW IB"')
code = code.replace('fg="grey", bg="#16213e", font=("Segoe UI", 9)', 'fg="#8888aa", bg="#16213e", font=("Segoe UI", 10)')
code = code.replace('font=("Segoe UI", 9, "italic")', 'font=("Segoe UI", 10, "italic")')

lbls_old = 'lbls = [("Claiming", "white", "#0f0f1a"), ("Dominant", "black", "#E2EFDA"), ("Sub-rows", "black", "#EBF3FB"), \n                ("Root shortage", "#9C0006", "#FFE0E0"), ("DICES", "black", "#FFF2CC"), ("SR", "#CC0000", "#FFD7D7")]'
lbls_new = 'lbls = [("Claiming", "white", "#0f0f1a"), ("Dominant", "black", "#E2EFDA"), ("Sub-rows", "black", "#EBF3FB"), \n                ("Root shortage", "#9C0006", "#FFE0E0"), ("DICES", "black", "#FFF2CC"), ("SR", "#CC0000", "#FFD7D7"),\n                ("■ Invalid inv", "#333", "#FFD0D0"), ("■ REBNI found", "#333", "#D0F0FF")]'
code = code.replace(lbls_old, lbls_new)

# Remove old manual frame and its _m_fields
code = re.sub(r'self._manual_ctrl = tk\.Frame.*?pady=5\)\n\s+self\._m_fields\(\)', '', code, flags=re.DOTALL)
code = re.sub(r'def _m_fields\(self\):.*?def _on_mode_change\(self\):', 'def _on_mode_change(self):', code, flags=re.DOTALL)

# Adjust _on_mode_change
code = re.sub(r'def _on_mode_change\(self\):.*?def _set_status', 'def _on_mode_change(self):\n        pass\n\n    def _set_status', code, flags=re.DOTALL)

# Update run & save buttons
code = code.replace(
    'self.run_btn = tk.Button(btn_f, text="▶ RUN INVESTIGATION", bg="#e94560", fg="white", font=("Segoe UI", 12, "bold"), width=22, height=2, command=self.start_run)',
    'self.run_btn = tk.Button(btn_f, text="▶ RUN INVESTIGATION", bg="#e94560", fg="white", font=("Segoe UI", 15, "bold"), padx=36, pady=14, command=self.start_run)'
)
code = code.replace(
    'self.save_btn = tk.Button(btn_f, text="💾  SAVE OUTPUT", bg="#2d6a4f", fg="white", font=("Segoe UI", 12, "bold"), width=22, height=2, state="disabled", command=self.save_output)',
    'self.save_btn = tk.Button(btn_f, text="💾  SAVE OUTPUT", bg="#2d6a4f", fg="white", font=("Segoe UI", 13, "bold"), padx=28, pady=14, state="disabled", command=self.save_output)'
)

# Update InvestigationEngine init call
code = code.replace(
    'self.engine = InvestigationEngine(rp, rs, idx, self._request_sid_from_user)',
    'self.engine = InvestigationEngine(rp, rs, rfb, idx, idx_fb, self._request_sid_from_user)'
)
code = code.replace('idx = build_invoice_index(df_i)', 'idx, idx_fb = build_invoice_index(df_i)')
code = code.replace('rp, rs = build_rebni_index(df_r)', 'rp, rs, rfb = build_rebni_index(df_r)')

# is_claiming parameter to run_auto in auto mode
code = code.replace(
    'safe_num(row[map_cols[\'PQV\']]))',
    'safe_num(row[map_cols[\'PQV\']]), is_claiming=True)'
)

# Replace _manual_step and old manual callbacks
manual_step_old = r'def _manual_step\(self\):.*?def _finish\(self\):'
manual_step_new = '''def _manual_step(self):
        m = self.curr_m
        is_claiming = (m['depth'] == 0)
        rows, matches, rq, n_rem = self.engine.build_one_level(m['b'], m['i'], m['s'], m['p'], m['a'], m['iq'], m['rem'], m['depth'], is_claiming)
        
        if 'processed_matches' not in m:
            m['processed_matches'] = set()
        
        new_matches = []
        for match in matches:
            if match['mtc_inv'] not in m['processed_matches']:
                new_matches.append(match)
        matches = new_matches
        
        if 'initial_rendered' not in m or not m['initial_rendered']:
            m['block'].extend(rows)
            for r in rows: self.preview.add_row(r)
            m['initial_rendered'] = True
            
        m['rem'] = n_rem
        rem = rows[0].get('remarks', '')
        
        if not matches or any(x in rem for x in ["Root cause", "Self", "REBNI", "SR"]):
            self.all_blocks.append(m['block']); self._next_manual_asin(); return
            
        self.root.after(0, lambda: self._show_manual_dialog(matches))

    def _show_manual_dialog(self, matches):
        dlg = ManualLevelDialog(self.root, matches, self.curr_m['rem'])
        self.root.wait_window(dlg)
        res = dlg.result
        
        if not res or res['action'] == 'stop':
            self.all_blocks.append(self.curr_m['block'])
            self._next_manual_asin()
            return
            
        match = res['chosen_match']
        self.curr_m['processed_matches'].add(match['mtc_inv'])
        
        if res['action'] == 'invalid':
            excl_qty = res['invalid_qty']
            invld_row = {
                'barcode': '[INVALID]', 'invoice': match['mtc_inv'], 'sid': '—',
                'po': match['mtc_po'], 'asin': match['mtc_asin'], 'inv_qty': match['inv_qty'],
                'rec_qty': '', 'mtc_qty': '', 'mtc_inv': '',
                'remarks': f"{int(excl_qty)} units matched to invalid invoice {match['mtc_inv']} — excluded from PQV",
                'date': match['date'], 'depth': self.curr_m['depth'], 'type': 'subrow'
            }
            self.curr_m['block'].append(invld_row)
            self.preview.add_row(invld_row)
            
            self.curr_m['rem'] = max(0, self.curr_m['rem'] - excl_qty)
            if self.curr_m['rem'] <= 0:
                self.all_blocks.append(self.curr_m['block'])
                self._next_manual_asin()
            else:
                remaining_matches = [m for m in matches if m['mtc_inv'] != match['mtc_inv']]
                if remaining_matches:
                    self.root.after(0, lambda: self._show_manual_dialog(remaining_matches))
                else:
                    self.all_blocks.append(self.curr_m['block'])
                    self._next_manual_asin()
        else:
            self.curr_m.update({
                'b': extract_sid(res['barcode']) or "[DICES]", 'i': match['mtc_inv'], 's': extract_sid(res['sid']), 
                'p': match['mtc_po'], 'a': match['mtc_asin'], 'iq': match['inv_qty'], 
                'depth': self.curr_m['depth'] + 1,
                'initial_rendered': False,
                'processed_matches': set()
            })
            threading.Thread(target=self._manual_step, daemon=True).start()

    def _finish(self):'''
code = re.sub(manual_step_old, manual_step_new, code, flags=re.DOTALL)

with open('mfi_tool.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Patching complete.")
