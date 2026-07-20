"""
P&L Forecast Automation Tool
============================
Tkinter desktop app for updating the FY P&L Summary workbook's external file
links and manual-input cells for a chosen forecasting cycle, without ever
breaking a link.

Run:  python app.py
Needs: openpyxl  (pip install openpyxl)
"""
import os
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import link_scanner
import file_matcher
import manual_scanner
import xml_updater

SCENARIOS = ['2+10', '5+7', '7+5', '10+2']


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('P&L Forecast Automation Tool')
        self.geometry('1180x760')
        self.configure(bg='#f4f6f8')

        self.base_file = tk.StringVar()
        self.input_folder = tk.StringVar()
        self.scenario = tk.StringVar(value=SCENARIOS[1])
        self.links = {}
        self.match_results = {}
        self.manual_cells = []
        self.override_paths = {}   # idx -> manually browsed path
        self.log_lines = []

        self._build_style()
        self._build_layout()

    # ---------- UI SCAFFOLDING ----------
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        style.configure('TButton', font=('Segoe UI', 10), padding=6)
        style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'))
        style.configure('Treeview', font=('Segoe UI', 9), rowheight=24)
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))
        style.configure('Scenario.TButton', font=('Segoe UI', 11, 'bold'), padding=10)

    def _build_layout(self):
        top = tk.Frame(self, bg='#1f2d3d', height=56)
        top.pack(fill='x')
        tk.Label(top, text='P&L Forecast Automation Tool', bg='#1f2d3d', fg='white',
                 font=('Segoe UI', 15, 'bold')).pack(side='left', padx=16, pady=10)

        body = tk.Frame(self, bg='#f4f6f8')
        body.pack(fill='both', expand=True, padx=12, pady=10)

        self._build_step1(body)
        self._build_step2(body)

        self.nb = ttk.Notebook(body)
        self.nb.pack(fill='both', expand=True, pady=(10, 0))
        self.tab_links = tk.Frame(self.nb, bg='white')
        self.tab_manual = tk.Frame(self.nb, bg='white')
        self.nb.add(self.tab_links, text='  External Links  ')
        self.nb.add(self.tab_manual, text='  Manual Inputs  ')
        self._build_links_tab()
        self._build_manual_tab()

        bottom = tk.Frame(self, bg='#f4f6f8')
        bottom.pack(fill='x', padx=12, pady=8)
        self.status = tk.Label(bottom, text='Ready.', bg='#f4f6f8', anchor='w', fg='#333')
        self.status.pack(side='left', fill='x', expand=True)
        ttk.Button(bottom, text='Generate Updated File', style='Accent.TButton',
                   command=self.on_generate).pack(side='right')

    def _build_step1(self, parent):
        frame = tk.LabelFrame(parent, text=' Step 1 — Base workbook & forecasting scenario ',
                               bg='white', font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        frame.pack(fill='x')

        row1 = tk.Frame(frame, bg='white')
        row1.pack(fill='x', pady=4)
        tk.Label(row1, text='Base P&L file:', bg='white', width=16, anchor='w').pack(side='left')
        tk.Entry(row1, textvariable=self.base_file).pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(row1, text='Browse…', command=self.on_browse_base).pack(side='left')
        ttk.Button(row1, text='Scan Links', command=self.on_scan_links).pack(side='left', padx=(6, 0))

        row2 = tk.Frame(frame, bg='white')
        row2.pack(fill='x', pady=(8, 0))
        tk.Label(row2, text='Forecast scenario:', bg='white', width=16, anchor='w').pack(side='left')
        for s in SCENARIOS:
            b = ttk.Button(row2, text=f'F{s}', style='Scenario.TButton',
                            command=lambda s=s: self.on_pick_scenario(s))
            b.pack(side='left', padx=4)
        self.scenario_label = tk.Label(row2, text=f'Selected: F{self.scenario.get()}',
                                        bg='white', fg='#0a5', font=('Segoe UI', 10, 'bold'))
        self.scenario_label.pack(side='left', padx=14)

    def _build_step2(self, parent):
        frame = tk.LabelFrame(parent, text=' Step 2 — Input folder (where this month\'s source files live) ',
                               bg='white', font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        frame.pack(fill='x', pady=(8, 0))
        row = tk.Frame(frame, bg='white')
        row.pack(fill='x')
        tk.Label(row, text='Input folder:', bg='white', width=16, anchor='w').pack(side='left')
        tk.Entry(row, textvariable=self.input_folder).pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(row, text='Browse…', command=self.on_browse_folder).pack(side='left')
        ttk.Button(row, text='Auto-Match Files', command=self.on_auto_match).pack(side='left', padx=(6, 0))

    def _build_links_tab(self):
        cols = ('idx', 'filename', 'sheets', 'refs', 'matched', 'score')
        headers = ['#', 'Linked file (as saved in workbook)', 'Used on sheet(s)', 'Ref count',
                   'Matched file in your folder', 'Match confidence']
        self.tree_links = ttk.Treeview(self.tab_links, columns=cols, show='headings', selectmode='browse')
        for c, h in zip(cols, headers):
            self.tree_links.heading(c, text=h)
        self.tree_links.column('idx', width=40, anchor='center')
        self.tree_links.column('filename', width=320)
        self.tree_links.column('sheets', width=200)
        self.tree_links.column('refs', width=70, anchor='center')
        self.tree_links.column('matched', width=340)
        self.tree_links.column('score', width=90, anchor='center')
        self.tree_links.pack(fill='both', expand=True, side='left')
        self.tree_links.bind('<Double-1>', self.on_link_row_double_click)

        vsb = ttk.Scrollbar(self.tab_links, orient='vertical', command=self.tree_links.yview)
        self.tree_links.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')

        self.tree_links.tag_configure('ok', background='#eaffea')
        self.tree_links.tag_configure('missing', background='#ffecec')

        tk.Label(self.tab_links, text='Double-click a row to manually browse for / override its matched file.',
                 bg='white', fg='#666').pack(side='bottom', anchor='w', padx=6, pady=4)

    def _build_manual_tab(self):
        top = tk.Frame(self.tab_manual, bg='white')
        top.pack(fill='x')
        tk.Label(top, text="Cells found where some periods are formulas/links and others are typed in "
                            "directly (e.g. the Connectivity row in Relief/Reclass). Double-click Value to edit.",
                 bg='white', fg='#666', wraplength=1000, justify='left').pack(anchor='w', padx=6, pady=6)

        cols = ('sheet', 'cell', 'label', 'value')
        headers = ['Sheet', 'Cell', 'Row label', 'Value (double-click to edit)']
        self.tree_manual = ttk.Treeview(self.tab_manual, columns=cols, show='headings', selectmode='browse')
        for c, h in zip(cols, headers):
            self.tree_manual.heading(c, text=h)
        self.tree_manual.column('sheet', width=140)
        self.tree_manual.column('cell', width=80, anchor='center')
        self.tree_manual.column('label', width=260)
        self.tree_manual.column('value', width=200, anchor='e')
        self.tree_manual.pack(fill='both', expand=True, side='left')
        self.tree_manual.bind('<Double-1>', self.on_manual_row_double_click)

        vsb = ttk.Scrollbar(self.tab_manual, orient='vertical', command=self.tree_manual.yview)
        self.tree_manual.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')

    # ---------- ACTIONS ----------
    def on_browse_base(self):
        path = filedialog.askopenfilename(filetypes=[('Excel workbook', '*.xlsx')])
        if path:
            self.base_file.set(path)

    def on_browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.input_folder.set(path)

    def on_pick_scenario(self, s):
        self.scenario.set(s)
        self.scenario_label.config(text=f'Selected: F{s}')
        if self.match_results:
            self.on_auto_match()  # re-score with the new scenario preference

    def on_scan_links(self):
        path = self.base_file.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning('No file', 'Pick a valid base .xlsx file first.')
            return
        self.status.config(text='Scanning external links…')
        self.update_idletasks()

        def work():
            try:
                links = link_scanner.scan_workbook(path)
                manual = manual_scanner.scan_manual_cells(path)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror('Scan failed', str(e)))
                return
            self.after(0, lambda: self._on_scan_done(links, manual))

        threading.Thread(target=work, daemon=True).start()

    def _on_scan_done(self, links, manual):
        self.links = links
        self.manual_cells = manual
        self._refresh_links_table()
        self._refresh_manual_table()
        self.status.config(text=f'Found {len(links)} external links and {len(manual)} manual-input candidate cells.')

    def on_auto_match(self):
        folder = self.input_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning('No folder', 'Pick a valid input folder first.')
            return
        if not self.links:
            messagebox.showwarning('Scan first', 'Scan the base file for links before matching.')
            return
        self.status.config(text='Matching files by name…')
        self.update_idletasks()

        def work():
            results = file_matcher.match_links_to_folder(
                self.links, folder, preferred_scenario=self.scenario.get()
            )
            self.after(0, lambda: self._on_match_done(results))

        threading.Thread(target=work, daemon=True).start()

    def _on_match_done(self, results):
        self.match_results = results
        # keep any manual overrides the user already set
        for idx, ov in self.override_paths.items():
            if idx in results:
                results[idx]['best'] = ov
                results[idx]['score'] = 1.0
        self._refresh_links_table()
        matched = sum(1 for r in results.values() if r['best'])
        self.status.config(text=f'Matched {matched} / {len(results)} links automatically.')

    def _refresh_links_table(self):
        self.tree_links.delete(*self.tree_links.get_children())
        for idx in sorted(self.links):
            link = self.links[idx]
            sheets = ', '.join(sorted(set(u[0] for u in link.used_in))) or '(unused)'
            match = self.match_results.get(idx, {})
            best = match.get('best', '')
            score = match.get('score', 0)
            tag = 'ok' if best else 'missing'
            self.tree_links.insert('', 'end', iid=str(idx), values=(
                idx, link.filename, sheets, len(link.used_in),
                best or '(not matched — double-click to browse)',
                f'{score:.0%}' if best else '-',
            ), tags=(tag,))

    def on_link_row_double_click(self, _event):
        sel = self.tree_links.selection()
        if not sel:
            return
        idx = int(sel[0])
        path = filedialog.askopenfilename(
            title=f'Select replacement for: {self.links[idx].filename}',
            filetypes=[('Excel workbook', '*.xlsx;*.xlsm;*.xls')],
        )
        if path:
            self.override_paths[idx] = path
            self.match_results.setdefault(idx, {})['best'] = path
            self.match_results[idx]['score'] = 1.0
            self._refresh_links_table()

    def _refresh_manual_table(self):
        self.tree_manual.delete(*self.tree_manual.get_children())
        for i, c in enumerate(self.manual_cells):
            self.tree_manual.insert('', 'end', iid=str(i), values=(
                c.sheet, c.cell, c.row_label or '(no label)', c.current_value,
            ))

    def on_manual_row_double_click(self, event):
        row_id = self.tree_manual.identify_row(event.y)
        col = self.tree_manual.identify_column(event.x)
        if not row_id or col != '#4':
            return
        i = int(row_id)
        c = self.manual_cells[i]
        self._edit_cell_popup(row_id, c)

    def _edit_cell_popup(self, row_id, cell_obj):
        top = tk.Toplevel(self)
        top.title(f'Edit {cell_obj.sheet}!{cell_obj.cell}')
        top.geometry('320x120')
        tk.Label(top, text=f'{cell_obj.row_label or cell_obj.sheet}  ({cell_obj.cell})').pack(pady=(12, 4))
        var = tk.StringVar(value=cell_obj.current_value)
        entry = tk.Entry(top, textvariable=var, font=('Segoe UI', 11))
        entry.pack(pady=4, padx=12, fill='x')
        entry.focus_set()

        def save():
            val = var.get().strip()
            try:
                float(val)
            except ValueError:
                messagebox.showerror('Invalid number', 'Enter a numeric value.')
                return
            cell_obj.current_value = val
            self.tree_manual.item(row_id, values=(
                cell_obj.sheet, cell_obj.cell, cell_obj.row_label or '(no label)', val,
            ))
            top.destroy()

        ttk.Button(top, text='Save', command=save).pack(pady=8)
        top.bind('<Return>', lambda e: save())

    # ---------- GENERATE ----------
    def on_generate(self):
        path = self.base_file.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning('No file', 'Pick a valid base .xlsx file first.')
            return
        if not self.links:
            messagebox.showwarning('Scan first', 'Scan the base file for links first.')
            return

        missing = [idx for idx in self.links if not self.match_results.get(idx, {}).get('best')]
        if missing:
            names = '\n'.join(f'  [{i}] {self.links[i].filename}' for i in missing[:15])
            proceed = messagebox.askyesno(
                'Some links unmatched',
                f'{len(missing)} external link(s) have no matched replacement file:\n\n{names}\n\n'
                'These links will be LEFT EXACTLY AS THEY ARE in the output (never broken), '
                'but will still point at their old source. Continue anyway?'
            )
            if not proceed:
                return

        out_dir = filedialog.askdirectory(title='Choose a folder to save the updated file')
        if not out_dir:
            return
        scenario = self.scenario.get()
        base_name = os.path.splitext(os.path.basename(path))[0]
        base_name = re.sub(r'F\d+\+\d+', f'F{scenario}', base_name) if re.search(r'F\d+\+\d+', base_name) \
            else f'{base_name}_F{scenario}'
        out_path = os.path.join(out_dir, f'{base_name}.xlsx')
        if os.path.exists(out_path):
            i = 2
            while os.path.exists(out_path):
                out_path = os.path.join(out_dir, f'{base_name} ({i}).xlsx')
                i += 1

        self.status.config(text='Generating updated workbook…')
        self.update_idletasks()

        def work():
            log = []
            try:
                idx_to_path = {idx: r['best'] for idx, r in self.match_results.items() if r.get('best')}
                xml_updater.relink_external_files(path, out_path, idx_to_path)
                xml_updater.refresh_cached_values(out_path, self.links, idx_to_path, log=log.append)

                sheet_map = xml_updater.sheet_name_to_file_map(out_path)
                updates = [(c.sheet, c.cell, c.current_value) for c in self.manual_cells]
                xml_updater.write_manual_values(out_path, sheet_map, updates, log=log.append)

                import zipfile
                with zipfile.ZipFile(out_path) as z:
                    bad = z.testzip()
                if bad:
                    raise RuntimeError(f'Output archive is corrupt at {bad}')
            except Exception as e:
                self.after(0, lambda: messagebox.showerror('Generation failed', str(e)))
                self.after(0, lambda: self.status.config(text='Failed.'))
                return
            self.after(0, lambda: self._on_generate_done(out_path, log))

        threading.Thread(target=work, daemon=True).start()

    def _on_generate_done(self, out_path, log):
        self.status.config(text=f'Done: {out_path}')
        note = ''
        if log:
            note = '\n\nNotes:\n' + '\n'.join(log[:10])
        messagebox.showinfo(
            'Updated file created',
            f'Saved:\n{out_path}\n\n'
            'Open it in Excel and choose "Update Values" if prompted, so every '
            'formula recalculates against the newly linked files.' + note
        )


if __name__ == '__main__':
    App().mainloop()
