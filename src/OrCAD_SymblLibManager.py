import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
import csv
import uuid
import tkinter.font as tkfont

class OrcadLibrarySpreadsheet:
    def __init__(self, root):
        self.root = root
        self.root.title("OrCAD Symbol Library Manager")
        self.root.iconbitmap('cpu.ico')

        self.template_tree = None
        self.template_path = None
        self.csv_data = []
        self.props = []
        self.filtered_ids = []
        self.updated_parts = set()
        self.parts = []
        self.current_filter = ''
        self.auto_fit = False
        self.sort_directions = {}
        self.create_widgets()

    def create_widgets(self):
        # --- Top Frame: Toolbar ---
        top = ttk.Frame(self.root, borderwidth=1, relief="solid")
        top.pack(fill=tk.X, padx=5, pady=5)

        # Compact button style
        button_style = ttk.Style()
        button_style.configure("Compact.TButton", font=("Segoe UI", 9), padding=(6, 2))

        # Button definitions
        buttons = [
            ("Load XML", self.load_xml),
            ("Load XML as Template", self.load_xml_template),
            ("Export CSV", self.export_csv),
            ("Import CSV", self.import_csv),
            ("Validate CSV", self.validate_csv),
            ("Save XML", self.save_xml),
            ("View History", self.show_update_history),
            ("Compare CSV to Template", self.compare_to_template),
            ("Toggle Auto-Fit", self.toggle_column_fit)
        ]

        # Add buttons
        for text, cmd in buttons:
            ttk.Button(top, text=text, command=cmd, style="Compact.TButton").pack(side=tk.LEFT, padx=2, pady=2)

        # Strict Save checkbox
        self.strict_save = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Strict Save", variable=self.strict_save).pack(side=tk.LEFT, padx=6, pady=2)

        # Search bar (right-aligned)
        self.search_var = tk.StringVar()
        ttk.Label(top, text="Search:").pack(side=tk.RIGHT, padx=(4, 2))
        search_entry = ttk.Entry(top, textvariable=self.search_var, width=20)
        search_entry.pack(side=tk.RIGHT, padx=(0, 4))
        search_entry.bind('<Return>', lambda e: self.apply_search())

        # --- Table Frame ---
        outer = ttk.Frame(self.root, borderwidth=2, relief="groove")
        outer.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        tc = tk.Frame(outer, bd=2, relief="sunken", bg="gray80")
        tc.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        tf = ttk.Frame(tc)
        tf.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)

        # Treeview styles
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=24, font=("Arial", 10),
                        background="white", fieldbackground="white",
                        bordercolor="gray80", relief="solid", borderwidth=1)
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'),
                        background="#e6e6e6", relief="solid", borderwidth=1)
        style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

        # Table widget
        self.table = ttk.Treeview(tf, show='headings', selectmode='extended')
        self.table.grid(row=0, column=0, sticky='nsew')
        self.table.bind('<ButtonRelease-1>', self.update_status)
        self.table.bind('<Double-1>', self.edit_cell)
        self.table.bind('<MouseWheel>', self._on_mousewheel)

        # Scrollbars
        vsb = ttk.Scrollbar(tf, orient='vertical', command=self.scroll_y_by_lines)
        hsb = ttk.Scrollbar(tf, orient='horizontal', command=self.scroll_x_by_columns)
        self.table.configure(yscroll=vsb.set, xscroll=hsb.set)
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        # --- Bottom Status Bar ---
        bottom = ttk.Frame(self.root)
        bottom.pack(fill=tk.X, padx=2, pady=(0, 2))

        self.status_var = tk.StringVar(value="No data loaded.")
        ttk.Label(bottom, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w') \
            .pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(bottom, text="Version 1.0", relief=tk.SUNKEN, anchor='e') \
            .pack(side=tk.RIGHT)

        ttk.Label(bottom, text="Developed By: Prem Mistry ©", relief=tk.SUNKEN, anchor='e') \
            .pack(side=tk.RIGHT)

    def edit_cell(self, event):
        row_id = self.table.identify_row(event.y)
        col = self.table.identify_column(event.x)
        if not row_id or col == '#0':
            return

        bbox = self.table.bbox(row_id, col)
        if not bbox:
            # cell isn't currently visible (e.g. scrolled out of view)
            return
        x, y, w, h = bbox

        old = self.table.set(row_id, self.table['columns'][int(col[1:]) - 1])
        entry = ttk.Entry(self.table)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, old)
        entry.focus()
        entry.bind('<Return>', lambda e: self._save_edit(entry, row_id, int(col[1:]) - 1))
        entry.bind('<FocusOut>', lambda e: self._save_edit(entry, row_id, int(col[1:]) - 1))

    def compare_to_template(self):
        if not self.template_tree:
            return messagebox.showerror("No Template", "Load template first.")
        if not self.csv_data:
            return messagebox.showerror("No CSV", "Import CSV first.")

        template_parts = self._extract_template_parts(self.template_tree)
        added_report = []
        removed_report = []

        for row in self.csv_data:
            pname = row[0]
            if not pname or pname not in template_parts:
                continue
            csv_props = {k.strip(): v.strip() for k, v in zip(self.props, row[1:]) if k.strip()}
            tmpl_props = template_parts[pname]

            csv_keys = {k.lower() for k in csv_props}
            tmpl_keys = {k.lower() for k in tmpl_props}

            added = sorted({k for k in csv_props if k.lower() not in tmpl_keys})
            removed = sorted({k for k in tmpl_props if k.lower() not in csv_keys})

            if added or removed:
                added_report.append(f"{pname}: +{', '.join(added) if added else '-'}")
                removed_report.append(f"{pname}: -{', '.join(removed) if removed else '-'}")

        if not added_report and not removed_report:
            return messagebox.showinfo("Compare", "CSV and template match exactly.")

        win = tk.Toplevel(self.root)
        win.title("Compare CSV to Template")
        txt = tk.Text(win, wrap='word');
        txt.pack(fill='both', expand=True)
        txt.insert('end', "Added Fields:\n" + "\n".join(added_report) + "\n\n")
        txt.insert('end', "Removed Fields:\n" + "\n".join(removed_report))
        txt.config(state='disabled')

    def scroll_y_by_lines(self, *args):
        if args[0] == 'scroll':
            lines = int(args[1]) * 3000  # Adjust vertical scroll speed (was 5000)
            self.table.yview_scroll(lines, 'units')
        else:
            self.table.yview(*args)

    def scroll_x_by_columns(self, *args):
        if args[0] == 'scroll':
            columns = int(args[1]) * 2000  # Faster horizontal scroll
            self.table.xview_scroll(columns, 'units')
        else:
            self.table.xview(*args)


    def _on_mousewheel(self, event):
        # On Windows, event.delta is usually 120 per scroll
        self.table.yview_scroll(int(-1 * (event.delta / 40)), 'units')

    def _save_edit(self, entry, row_id, ci):
        new = entry.get(); entry.destroy()
        col = self.table['columns'][ci]
        self.table.set(row_id, col, new)
        try:
            pkg, libpart = self.parts[int(row_id)]
            if pkg.find('Defn') is not None and col in pkg.find('Defn').attrib:
                pkg.find('Defn').set(col, new)
            else:
                for sup in libpart.findall('.//SymbolUserProp/Defn'):
                    if sup.get('name') == col:
                        sup.set('val', new)
                        break
                else:
                    nv = libpart.find('NormalView')
                    if nv is None:
                        nv = ET.SubElement(libpart, 'NormalView')
                    sup = ET.SubElement(nv, 'SymbolUserProp')
                    ET.SubElement(sup, 'Defn', name=col, val=new)
            self.updated_parts.add(self.table.set(row_id, 'PartName'))
        except: pass
        self.update_status()

    def toggle_column_fit(self):
        self.auto_fit = not self.auto_fit
        self.fit_columns_to_content()
        mode = "Auto-fit" if self.auto_fit else "Manual"
        self.status_var.set(f"Column mode: {mode}")

    def fit_columns_to_content(self):
        font = tkfont.Font()
        for col in self.table['columns']:
            if self.auto_fit:
                max_width = font.measure(col) + 20  # start with header width
                for iid in self.table.get_children():
                    val = str(self.table.set(iid, col))
                    width = font.measure(val) + 10
                    if width > max_width:
                        max_width = width
                self.table.column(col, width=max_width, stretch=False)
            else:
                self.table.column(col, width=120, stretch=True)

    def load_xml(self):
        path = filedialog.askopenfilename(filetypes=[("XML Files", "*.xml")])
        if not path: return
        self.tree = ET.parse(path)
        self.filename = path
        root = self.tree.getroot()
        self.parts = [(pkg, pkg.find('LibPart')) for pkg in root.findall('.//Package') if pkg.find('LibPart') is not None]
        self._extract_props()
        self.populate_table()
        self.status_var.set(f"Loaded {len(self.parts)} parts")

    def populate_table(self):
        cols = ['PartName'] + self.props
        # configure columns and enable resizing
        self.table.config(columns=cols, displaycolumns=cols)
        for c in cols:
            self.table.heading(c, text=c, command=lambda col=c: self.sort_by_column(col))
            self.table.column(c, width=120, anchor='w', stretch=True)


        # clear previous rows
        self.table.delete(*self.table.get_children())
        self.filtered_ids.clear()

        # populate with unique IDs to avoid duplicates
        for pkg, libpart in self.parts:
            defn = pkg.find('Defn')
            lp_defn = libpart.find('Defn')
            pname = lp_defn.get('CellName') if lp_defn is not None else defn.get('name')
            # gather all attributes and user props
            attr_map = defn.attrib if defn is not None else {}
            user_map = {sup.get('name'): sup.get('val', '') for sup in libpart.findall('.//SymbolUserProp/Defn')}
            merged = {**attr_map, **user_map}
            row = [pname] + [merged.get(p, '') for p in self.props]
            iid = str(uuid.uuid4())
            self.table.insert('', 'end', iid=iid, values=row)
            self.filtered_ids.append(iid)


    def sort_by_column(self, col):
        items = [(self.table.set(k, col), k) for k in self.table.get_children('')]

        # Determine sort direction
        descending = self.sort_directions.get(col, False)
        try:
            # Try to sort as numbers first
            items.sort(key=lambda x: float(x[0]), reverse=descending)
        except ValueError:
            # Fall back to string sort
            items.sort(key=lambda x: x[0].lower(), reverse=descending)

        for index, (val, k) in enumerate(items):
            self.table.move(k, '', index)

        # Toggle sort direction
        self.sort_directions[col] = not descending

    def load_xml_template(self):
        path = filedialog.askopenfilename(filetypes=[("XML Files", "*.xml")])
        if not path: return
        self.template_tree = ET.parse(path)
        self.template_path = path
        messagebox.showinfo("Template Loaded", f"Template loaded from {path}")
        self.status_var.set(f"Template: {path}")

    def apply_search(self):
        term = self.search_var.get().strip().lower()
        self.current_filter = term  # store for status
        for iid in self.filtered_ids:
            values = self.table.item(iid)['values']
            # Make everything a string first
            line = ' '.join(str(v) for v in values).lower()
            if term in line:
                self.table.reattach(iid, '', 'end')
            else:
                self.table.detach(iid)
        self.update_status()

    def export_csv(self):
        # 1) Grab current columns and rows from the Treeview
        cols = list(self.table['columns'])
        rows = [self.table.item(iid)['values'] for iid in self.table.get_children()]

        # 2) Ask for filename and write
        file = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV Files", "*.csv")])
        if not file:
            return
        with open(file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            writer.writerows(rows)

        self.status_var.set(f"Exported {len(rows)} rows to CSV")

    def import_csv(self):
        file = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file: return
        with open(file, newline='', encoding='utf-8') as f:
            rows = list(csv.reader(f))
        if not rows or "PartName" not in rows[0]: return messagebox.showerror("CSV Error", "Missing 'PartName'")
        header = rows[0]
        data = [[str(cell).strip() for cell in row] for row in rows[1:]]
        self.csv_data = data; self.props = header[1:]
        self.table.config(columns=header, displaycolumns=header)
        for c in header:
            self.table.heading(c, text=c); self.table.column(c, width=120, stretch=True)
        self.table.delete(*self.table.get_children()); self.filtered_ids.clear()
        for row in data:
            iid = str(uuid.uuid4())
            self.table.insert('', 'end', iid=iid, values=row)
            self.filtered_ids.append(iid)
        self.status_var.set(f"Imported {len(data)} rows")

    def validate_csv(self):
        if not self.template_tree:
            messagebox.showerror("No Template", "Load template first.")
            return

        # template part names (case‐insensitive)
        tmpl = self._extract_template_parts(self.template_tree)
        tmpl_keys = {name.lower() for name in tmpl}

        # table part names
        table_parts = [self.table.set(iid, 'PartName').strip()
                       for iid in self.table.get_children()]
        updated = [p for p in table_parts if p.lower() in tmpl_keys]
        added = [p for p in table_parts if p.lower() not in tmpl_keys]

        msg = (
            f"CSV Validation Results:\n\n"
            f"  To Update (in template): {len(updated)}\n"
            f"    Examples: {updated[:5]}\n\n"
            f"  To Add (new parts): {len(added)}\n"
            f"    Examples: {added[:5]}"
        )
        messagebox.showinfo("Validate CSV", msg)

        # reflect in status bar
        self.current_filter = f"Validated (upd {len(updated)}, add {len(added)})"
        self.update_status()

    def save_xml(self):
        if not self.template_tree:
            messagebox.showerror("No Template", "Load template first.")
            return

        tree = self.template_tree
        root = tree.getroot()
        cols = list(self.table['columns'])
        update_map = {}

        # 1) Extract updated values from table and normalize keys
        for iid in self.table.get_children():
            values = self.table.item(iid)['values']
            partname = str(values[0]).strip()
            key = partname.lower()

            prop_map = dict(zip(cols[1:], values[1:]))

            # Identify Defn attributes from any known part
            defn_keys = set()
            for pkg, libpart in self.parts:
                if pkg is not None and pkg.find('Defn') is not None:
                    defn_keys.update(pkg.find('Defn').attrib.keys())

            props_defn = {k: prop_map[k] for k in prop_map if k in defn_keys}
            props_sup = {k: prop_map[k] for k in prop_map if k not in defn_keys}
            update_map[key] = (props_defn, props_sup)

        template_parts = self._extract_template_parts(tree)
        count = 0

        # 2) Update template with normalized keys
        for pkg in root.findall('.//Package'):
            libpart = pkg.find('LibPart')
            if libpart is None:
                continue

            pkg_defn = pkg.find('Defn')
            lib_defn = libpart.find('Defn')
            if pkg_defn is None or lib_defn is None:
                continue

            pname = lib_defn.get('CellName')
            if not pname:
                continue

            key = str(pname).strip().lower()
            if key not in update_map:
                print(f"⚠ Skipping: '{pname}' not found in CSV update map")
                continue

            props_defn, props_sup = update_map[key]

            # ✅ Always update all <Package><Defn> attributes (mandatory fields)
            for k, v in props_defn.items():
                pkg_defn.set(k, str(v).strip() if v is not None else "")

            # ✅ Keep only CellName in <LibPart><Defn>
            for attrib in list(lib_defn.attrib):
                if attrib != 'CellName':
                    del lib_defn.attrib[attrib]

            if 'CellName' in props_defn:
                lib_defn.set('CellName', props_defn['CellName'])

            # ✅ Ensure <NormalView> exists
            nv = libpart.find('NormalView')
            if nv is None:
                nv = ET.SubElement(libpart, 'NormalView')

            # ✅ Remove old SymbolUserProps
            for old in list(nv.findall('SymbolUserProp')):
                nv.remove(old)

            # ✅ Add SymbolUserProps from template mask (preserve order)
            mask = template_parts.get(pname, {})
            for prop_name, orig_val in mask.items():
                val = props_sup.get(prop_name, orig_val)
                if self.strict_save.get() and not val:
                    continue
                sup = ET.SubElement(nv, 'SymbolUserProp')
                ET.SubElement(sup, 'Defn', name=str(prop_name), val=str(val))

            # ✅ Add new SymbolUserProps not in template or Defn
            used_keys = {k.lower() for k in mask}
            existing_defn_keys = {k.lower() for k in props_defn}
            for k, v in props_sup.items():
                if not v or k.lower() in used_keys or k.lower() in existing_defn_keys:
                    continue
                sup = ET.SubElement(nv, 'SymbolUserProp')
                ET.SubElement(sup, 'Defn', name=str(k), val=str(v))

            count += 1

        # 3) Save output file
        fp = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML", "*.xml")])
        if fp:
            ET.indent(tree, space="  ")
            tree.write(fp, encoding='utf-8', xml_declaration=True)
            self.status_var.set(f"Saved {count} parts")
            messagebox.showinfo("Saved", f"Updated {count} parts to XML")

    def show_update_history(self):
        if not self.updated_parts:
            messagebox.showinfo("History", "No updates recorded.")
            return
        win = tk.Toplevel(self.root)
        win.title("Update History")
        text = tk.Text(win, wrap='word')
        text.pack(fill='both', expand=True)
        for p in sorted(self.updated_parts):
            text.insert('end', p + "\n")
        text.config(state='disabled')

    def update_status(self, event=None):
        total = len(self.filtered_ids)
        selected = len(self.table.selection())
        updated = len(self.updated_parts)
        filt = getattr(self, 'current_filter', '')
        filter_text = f" | Filter: '{filt}'" if filt else ''
        self.status_var.set(f"Total: {total} | Selected: {selected} | Updated: {updated}{filter_text}")

    def _extract_props(self):
        names = set()
        for pkg, libpart in self.parts:
            defn = pkg.find('Defn'); names.update(defn.attrib.keys())
            for sup in libpart.findall('.//SymbolUserProp/Defn'):
                names.add(sup.get('name'))
        self.props = sorted(names)

    def _extract_template_parts(self, tree):
        parts = {}; root = tree.getroot()
        for pkg in root.findall('.//Package'):
            libpart = pkg.find('LibPart')
            defn = libpart.find('Defn') if libpart is not None else None
            pname = defn.get('CellName') if defn is not None else None
            if not pname: continue
            sup = {s.get('name'): s.get('val', '') for s in libpart.findall('.//SymbolUserProp/Defn')}
            parts[pname] = sup
        return parts

if __name__ == '__main__':
    root = tk.Tk()
    OrcadLibrarySpreadsheet(root)
    root.mainloop()