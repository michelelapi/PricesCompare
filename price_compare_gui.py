import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import json
import os

CONFIG_FILE = 'config.json'

class PriceCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Confronta prezzi')
        self.files = []
        self.config = self.load_config()
        self.setup_ui()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {'csv_separator': ',', 'decimal_separator': '.', 'thousands_separator': ','}

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            messagebox.showerror('Error', f'Impossibile salvare la configurazione: {e}')

    def setup_ui(self):
        # Create a menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='File', menu=file_menu)
        file_menu.add_command(label='Salva risultati come CSV', command=self.save_results, state='disabled')
        file_menu.add_command(label='Salva risultati temporanei', command=self.save_temporary_results, state='disabled')
        file_menu.add_command(label='Carica risultati da CSV', command=self.load_results_from_csv)
        file_menu.add_separator()
        file_menu.add_command(label='Esci', command=self.root.quit)
        # Config menu
        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='Configurazione', menu=config_menu)
        config_menu.add_command(label='Impostazioni', command=self.open_config_dialog)
        # Store menu items for enabling/disabling
        self.menu_save_results = file_menu.entryconfig('Salva risultati come CSV', state='disabled')
        self.menu_save_temp = file_menu.entryconfig('Salva risultati temporanei', state='disabled')
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # --- Button row for file selection ---
        btn_row = ttk.Frame(frame)
        btn_row.pack(pady=5)

        self.select_btn = ttk.Button(btn_row, text='Scegli i listini da confrontare', command=self.select_files)
        self.select_btn.pack(side=tk.LEFT, padx=2)

        # Remove the add_btn (Aggiungi listino)
        # self.add_btn = ttk.Button(btn_row, text='Aggiungi listino', command=self.add_file)
        # self.add_btn.pack(side=tk.LEFT, padx=2)

        self.clear_btn = ttk.Button(btn_row, text='Svuota lista', command=self.clear_files)
        self.clear_btn.pack(side=tk.LEFT, padx=2)

        # Remove old placement of select_btn
        # self.select_btn = ttk.Button(frame, text='Scegli i listini da confrontare', command=self.select_files)
        # self.select_btn.pack(pady=5)

        self.files_label = ttk.Label(frame, text='Nessun listino selezionato.')
        self.files_label.pack(pady=5)

        self.result_box = tk.Text(frame, height=15, width=60)
        self.result_box.pack(pady=5)
        self.result_box.insert(tk.END, 'I risultati appariranno qui.')
        self.result_box.config(state=tk.DISABLED)
        # Remove button placements for save_btn, save_temp_btn, load_btn, config_btn
        # self.save_btn = ttk.Button(frame, text='Save Results as CSV', command=self.save_results, state=tk.DISABLED)
        # self.save_btn.pack(pady=5)
        # self.save_temp_btn = ttk.Button(frame, text='Save Temporary Results', command=self.save_temporary_results, state=tk.DISABLED)
        # self.save_temp_btn.pack(pady=5)
        # self.load_btn = ttk.Button(frame, text='Load Results From CSV', command=self.load_results_from_csv)
        # self.load_btn.pack(pady=5)
        self.file_menu = file_menu

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx;*.xls')])
        if files:
            # Only add files not already in the list
            new_files = [f for f in files if f not in self.files]
            duplicate_files = [f for f in files if f in self.files]
            if duplicate_files:
                messagebox.showinfo('Info', 'Listino già selezionato')
            if not new_files:
                # If all selected files are duplicates, just update label and return
                self.files_label.config(text=f'File selezionati: {len(self.files)}')
                return
            self.files.extend(new_files)
            self.files_label.config(text=f'File selezionati: {len(self.files)}')
            if not hasattr(self, 'file_column_mappings') or not isinstance(self.file_column_mappings, list):
                self.file_column_mappings = []
            for file in new_files:
                try:
                    item_col, price_col, description_col, header_idx = self.ask_column_mapping_with_header(file)
                    if item_col is None or price_col is None or description_col is None:
                        messagebox.showwarning('Warning', f'Salto file: {file}')
                        continue
                    self.file_column_mappings.append({
                        'file': file,
                        'item_col': item_col,
                        'price_col': price_col,
                        'description_col': description_col,
                        'header_idx': header_idx
                    })
                except Exception as e:
                    messagebox.showerror('Error', f'Impossibile leggere {file}: {e}')
            if self.file_column_mappings:
                self.compare_and_display()
            else:
                self.result_box.config(state=tk.NORMAL)
                self.result_box.delete(1.0, tk.END)
                self.result_box.insert(tk.END, 'Nessun file/colonna valido selezionato.')
                self.result_box.config(state=tk.DISABLED)
                self.file_menu.entryconfig('Salva risultati come CSV', state='disabled')
                self.file_menu.entryconfig('Salva risultati temporanei', state='disabled')
        else:
            self.files_label.config(text=f'File selezionati: {len(self.files)}')

    # Remove add_file method

    def clear_files(self):
        self.files = []
        self.file_column_mappings = []
        self.files_label.config(text='Nessun listino selezionato.')
        # Remove result widgets if any
        if hasattr(self, 'result_tree') and self.result_tree:
            self.result_tree.destroy()
        if hasattr(self, 'tree_scroll_x') and self.tree_scroll_x:
            self.tree_scroll_x.destroy()
        if hasattr(self, 'tree_scroll_y') and self.tree_scroll_y:
            self.tree_scroll_y.destroy()
        if hasattr(self, 'result_frame') and self.result_frame:
            self.result_frame.destroy()
        if hasattr(self, 'search_frame') and self.search_frame:
            self.search_frame.destroy()
        if hasattr(self, 'total_label') and self.total_label:
            self.total_var.set("")
            self.total_label.pack_forget()
        self.result_box.config(state=tk.NORMAL)
        self.result_box.delete(1.0, tk.END)
        self.result_box.insert(tk.END, 'I risultati appariranno qui.')
        self.result_box.config(state=tk.DISABLED)
        self.result_box.pack(pady=5)
        # Disable menu items for saving
        self.file_menu.entryconfig('Salva risultati come CSV', state='disabled')
        self.file_menu.entryconfig('Salva risultati temporanei', state='disabled')

    def ask_column_mapping_with_header(self, file):
        import pandas as pd
        # Read the first 20 rows for preview
        try:
            preview_df = pd.read_excel(file, nrows=20, header=None)
        except Exception as e:
            messagebox.showerror('Error', f'Impossibile visualizzare {file}: {e}')
            return None, None, None, None
        dialog = tk.Toplevel(self.root)
        dialog.title(f'Seleziona la riga di intestazione per {file.split("/")[-1]}')
        dialog.grab_set()
        tk.Label(dialog, text=f'Seleziona la riga di intestazione per il file:\n{file}').pack(pady=5)
        # Use Treeview for table-like display
        columns = [f"Col {i+1}" for i in range(preview_df.shape[1])]
        style = ttk.Style()
        style.configure('HeaderTreeview.Heading', background='#d9ead3', foreground='black', font=('Arial', 10, 'bold'))
        style.map('HeaderTreeview.Heading', background=[('active', '#b6d7a8')])
        style.configure('header_evenrow', background='#f2f2f2')
        style.configure('header_oddrow', background='#ffffff')
        tree = ttk.Treeview(dialog, columns=columns, show='headings', height=15)
        for i, col in enumerate(columns):
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor='center')
        for idx, row in preview_df.iterrows():
            values = [str(x) if not pd.isna(x) else "" for x in row.values]
            tag = 'header_evenrow' if idx % 2 == 0 else 'header_oddrow'
            tree.insert('', 'end', iid=idx, values=values, tags=(tag,))
        tree.tag_configure('header_evenrow', background='#f2f2f2')
        tree.tag_configure('header_oddrow', background='#ffffff')
        tree.pack(pady=5)
        # Select first row by default
        tree.selection_set(tree.get_children()[0])
        header_row_idx = [0]
        def on_tree_select(event=None):
            sel = tree.selection()
            if sel:
                header_row_idx[0] = int(sel[0])
        tree.bind('<<TreeviewSelect>>', on_tree_select)
        def on_ok_header():
            dialog.destroy()
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text='OK', command=on_ok_header).pack(side=tk.LEFT, padx=5)
        dialog.wait_window()
        header_idx = header_row_idx[0]
        columns = [str(x) for x in preview_df.iloc[header_idx].values]
        # Now show the column selection dialog as before
        dialog2 = tk.Toplevel(self.root)
        dialog2.title(f'Seleziona le colonne per {file.split("/")[-1]}')
        dialog2.grab_set()
        tk.Label(dialog2, text=f'Seleziona le colonne per il file:\n{file}').pack(pady=5)
        tk.Label(dialog2, text='Codice articolo/EAN:').pack()
        item_var = tk.StringVar(dialog2)
        item_var.set(columns[0] if columns else '')
        item_menu = ttk.Combobox(dialog2, textvariable=item_var, values=columns, state='readonly')
        item_menu.pack(pady=2)
        tk.Label(dialog2, text='Prezzo:').pack()
        price_var = tk.StringVar(dialog2)
        price_var.set(columns[0] if columns else '')
        price_menu = ttk.Combobox(dialog2, textvariable=price_var, values=columns, state='readonly')
        price_menu.pack(pady=2)
        tk.Label(dialog2, text='Descrizione:').pack()
        description_var = tk.StringVar(dialog2)
        description_var.set(columns[0] if columns else '')
        description_menu = ttk.Combobox(dialog2, textvariable=description_var, values=columns, state='readonly')
        description_menu.pack(pady=2)
        result = {'item': None, 'price': None, 'description': None}
        def on_ok():
            result['item'] = item_var.get()
            result['price'] = price_var.get()
            result['description'] = description_var.get()
            dialog2.destroy()
        def on_cancel():
            dialog2.destroy()
        btn_frame2 = ttk.Frame(dialog2)
        btn_frame2.pack(pady=5)
        ttk.Button(btn_frame2, text='OK', command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame2, text='Cancel', command=on_cancel).pack(side=tk.LEFT, padx=5)
        dialog2.wait_window()
        return result['item'], result['price'], result['description'], header_idx

    def compare_and_display(self):
        # Read all items and prices
        all_items = []  # List of dicts: {item, price, file, original_item}
        for mapping in self.file_column_mappings:
            try:
                df = pd.read_excel(
                    mapping['file'],
                    usecols=[mapping['item_col'], mapping['price_col'], mapping['description_col']],
                    header=mapping['header_idx']
                )
                for _, row in df.iterrows():
                    # Try to treat item as a number if possible
                    item_cell = row[mapping['item_col']]
                    if pd.isna(item_cell):
                        original_item = ''
                        item_key = ''
                    else:
                        try:
                            # Try integer first, then float
                            if float(item_cell).is_integer():
                                original_item = int(item_cell)
                            else:
                                original_item = int(float(item_cell))
                            item_key = str(original_item)
                        except Exception:
                            original_item = str(item_cell)
                            item_key = original_item.strip().lower()
                    price_cell = row[mapping['price_col']]
                    description_cell = row[mapping['description_col']]
                    # If price is missing or empty, treat as zero
                    if pd.isna(price_cell) or str(price_cell).strip() == '':
                        price = 0.0
                    else:
                        price = price_cell
                    all_items.append({'item_key': item_key, 'price': price, 'file': mapping['file'], 'original_item': original_item, 'description': description_cell})
            except Exception as e:
                messagebox.showerror('Error', f'Impossibile processare {mapping["file"]}: {e}')
        # Find lowest price for each item (case-insensitive, stripped)
        best_prices = {}
        for entry in all_items:
            item_key = entry['item_key']
            if item_key not in best_prices or entry['price'] < best_prices[item_key]['price']:
                best_prices[item_key] = {'price': entry['price'], 'file': entry['file'], 'original_item': entry['original_item'], 'description': entry['description']}
        # Prepare results for display
        results = []
        for item_key, data in sorted(best_prices.items(), key=lambda x: str(x[1]['description']).lower()):
            results.append({'item': data['original_item'], 'price': data['price'], 'description': data['description'], 'file': data['file']})
        self.display_results(results)
        self.comparison_results = results  # Save for CSV export
        # Enable menu items for saving if results exist
        if hasattr(self, 'root') and hasattr(self, 'menu_save_results') and hasattr(self, 'menu_save_temp'):
            state = 'normal' if results else 'disabled'
            # Enable menu items for saving
            self.file_menu.entryconfig('Salva risultati come CSV', state=state)
            self.file_menu.entryconfig('Salva risultati temporanei', state=state)

    def display_results(self, results):
        """
        This function is responsible for displaying the comparison results in a Treeview widget.
        It first clears any existing result widgets, then checks if there are any results to display.
        If there are no results, it displays a message indicating so. If there are results, it hides
        the text box and creates a new Treeview widget with scrollbars to display the results.
        The results are stored for future searching and the Treeview is populated with the results.
        """
        # Remove old result widgets if any
        if hasattr(self, 'result_tree') and self.result_tree:
            self.result_tree.destroy()
        if hasattr(self, 'tree_scroll_x') and self.tree_scroll_x:
            self.tree_scroll_x.destroy()
        if hasattr(self, 'tree_scroll_y') and self.tree_scroll_y:
            self.tree_scroll_y.destroy()
        if hasattr(self, 'result_frame') and self.result_frame:
            self.result_frame.destroy()
        if hasattr(self, 'search_frame') and self.search_frame:
            self.search_frame.destroy()
        # If no results, show a message
        if not results:
            self.result_box.config(state=tk.NORMAL)
            self.result_box.delete(1.0, tk.END)
            self.result_box.insert(tk.END, 'Nessun risultato da visualizzare.')
            self.result_box.config(state=tk.DISABLED)
            return
        # Hide the text box
        self.result_box.config(state=tk.NORMAL)
        self.result_box.delete(1.0, tk.END)
        self.result_box.config(state=tk.DISABLED)
        self.result_box.pack_forget()
        # Store all results for searching
        self._all_results = results
        # Create a frame to hold the Treeview and scrollbars
        self.result_frame = ttk.Frame(self.root)
        self.result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        columns = ['Item', 'Description', 'Lowest Price', 'Quantity', 'Source File']
        style = ttk.Style()
        style.configure('Treeview.Heading', background='#d9ead3', foreground='black', font=('Arial', 10, 'bold'))
        style.map('Treeview.Heading', background=[('active', '#b6d7a8')])
        style.configure('evenrow', background='#f2f2f2')
        style.configure('oddrow', background='#ffffff')
        self.result_tree = ttk.Treeview(self.result_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=150, anchor='center', stretch=True)
        self._populate_result_tree(results)
        self.result_tree.tag_configure('evenrow', background='#f2f2f2')
        self.result_tree.tag_configure('oddrow', background='#ffffff')
        # Add scrollbars
        self.tree_scroll_y = ttk.Scrollbar(self.result_frame, orient='vertical', command=self.result_tree.yview)
        self.tree_scroll_x = ttk.Scrollbar(self.result_frame, orient='horizontal', command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.result_tree.grid(row=0, column=0, sticky='nsew')
        self.tree_scroll_y.grid(row=0, column=1, sticky='ns')
        self.tree_scroll_x.grid(row=1, column=0, sticky='ew')
        self.result_frame.rowconfigure(0, weight=1)
        self.result_frame.columnconfigure(0, weight=1)
        # Make cells editable on double-click
        def on_double_click(event):
            region = self.result_tree.identify('region', event.x, event.y)
            if region != 'cell':
                return
            row_id = self.result_tree.identify_row(event.y)
            col_id = self.result_tree.identify_column(event.x)
            if not row_id or not col_id:
                return
            col_idx = int(col_id.replace('#', '')) - 1
            x, y, width, height = self.result_tree.bbox(row_id, col_id)
            value = self.result_tree.set(row_id, columns[col_idx])
            entry = tk.Entry(self.result_tree)
            entry.place(x=x, y=y, width=width, height=height)
            entry.insert(0, value)
            entry.focus()
            def save_edit(event=None):
                self.result_tree.set(row_id, columns[col_idx], entry.get())
                entry.destroy()
                # If the edited column is 'Quantity', update the total
                if columns[col_idx] == 'Quantity':
                    self._update_total_label()
            entry.bind('<Return>', save_edit)
            entry.bind('<FocusOut>', save_edit)
        self.result_tree.bind('<Double-1>', on_double_click)
        # Add search input and dropdown below the table
        self.search_frame = ttk.Frame(self.root)
        self.search_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Label(self.search_frame, text='Cerca:').pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(self.search_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_column_var = tk.StringVar()
        self.search_column_var.set('Description')
        self.search_column_menu = ttk.Combobox(self.search_frame, textvariable=self.search_column_var, values=['Description', 'Item'], state='readonly', width=12)
        self.search_column_menu.pack(side=tk.LEFT, padx=5)
        def on_search(*args):
            term = self.search_var.get().lower()
            col = self.search_column_var.get().lower()
            # Instead of filtering, just select the first matching row
            for row_id in self.result_tree.get_children():
                values = self.result_tree.item(row_id)['values']
                if col == 'description':
                    cell = str(values[1]).lower()
                else:
                    cell = str(values[0]).lower()
                if term in cell:
                    self.result_tree.selection_set(row_id)
                    self.result_tree.see(row_id)
                    break
            else:
                self.result_tree.selection_remove(self.result_tree.selection())
        self.search_var.trace_add('write', lambda *args: on_search())
        self.search_column_var.trace_add('write', lambda *args: on_search())

        # Add total label below the search frame
        self.total_var = tk.StringVar()
        self.total_label = ttk.Label(self.root, textvariable=self.total_var, font=('Arial', 12, 'bold'))
        self.total_label.pack(pady=(0, 10))
        self._update_total_label()

    def _populate_result_tree(self, results):
        # Helper to clear and repopulate the result_tree
        for row in self.result_tree.get_children():
            self.result_tree.delete(row)
        for idx, r in enumerate(results):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            self.result_tree.insert('', 'end', values=[
                r['item'],
                r['description'],
                r['price'],
                r.get('quantity', ''),
                r['file'].split('/')[-1] if 'file' in r else ''
            ], tags=(tag,))

    def save_results(self):
        import csv
        from tkinter import filedialog
        if not hasattr(self, 'result_tree') or not self.result_tree.get_children():
            messagebox.showwarning('Warning', 'Nessun risultato da salvare.')
            return
        # Let user select destination folder
        folder_path = filedialog.askdirectory(title='Seleziona la cartella di destinazione per i file CSV')
        if not folder_path:
            return
        sep = self.config.get('csv_separator', ',')
        dec_sep = self.config.get('decimal_separator', '.')
        thou_sep = self.config.get('thousands_separator', ',')
        # Gather current data from the Treeview
        tree_data = []
        for row_id in self.result_tree.get_children():
            values = self.result_tree.item(row_id)['values']
            # Ensure we have all columns (Item, Description, Lowest Price, Quantity, Source File)
            if len(values) < 5:
                values += [''] * (5 - len(values))
            # # Format price according to config
            # price = values[2]
            # try:
            #     price_float = float(str(price).replace(thou_sep, '').replace(dec_sep, '.'))
            #     price_str = f"{price_float:,.4f}".replace(',', 'X').replace('.', dec_sep).replace('X', thou_sep)
            # except Exception:
            #     price_str = str(price)
            tree_data.append({
                'item': values[0],
                'description': values[1],
                'price': values[2],
                'quantity': values[3],
                'file': values[4],
            })
        # Group results by source file
        results_by_file = {}
        for r in tree_data:
            source_file = r['file']
            if source_file not in results_by_file:
                results_by_file[source_file] = []
            results_by_file[source_file].append(r)
        try:
            with open(os.path.join(folder_path, 'result_compared.csv'), 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=sep)
                writer.writerow(['Item', 'Description', 'Lowest Price', 'Quantity', 'Source File'])
                for r in tree_data:
                    if str(r['quantity']).strip() and str(r['quantity']).strip() != '0':
                        price = r['price']
                        try:
                            # price_float = float(str(price).replace(thou_sep, '').replace(dec_sep, '.'))
                            price_str = f"{price}".replace(',', 'X').replace('.', dec_sep).replace('X', thou_sep)
                        except Exception:
                            price_str = str(price)
                        writer.writerow([r['item'], r['description'], price_str, r['quantity'], r['file']])
        except Exception as e:
            messagebox.showerror('Error', f'Failed to save combined CSV: {e}')
        # Save one CSV per source Excel file
        for source_file, rows in results_by_file.items():
            base_name = os.path.splitext(os.path.basename(source_file))[0]
            csv_name = f"{base_name}_compared.csv"
            csv_path = os.path.join(folder_path, csv_name)
            try:
                with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter=sep)
                    writer.writerow(['Item', 'Description', 'Lowest Price', 'Quantity'])
                    for r in rows:
                        if str(r['quantity']).strip() and str(r['quantity']).strip() != '0':
                            price = r['price']
                            try:
                                # price_float = float(str(price).replace(thou_sep, '').replace(dec_sep, '.'))
                                price_str = f"{price}".replace(',', 'X').replace('.', dec_sep).replace('X', thou_sep)
                            except Exception:
                                price_str = str(price)
                            writer.writerow([r['item'], r['description'], price_str, r['quantity']])
            except Exception as e:
                messagebox.showerror('Error', f'Failed to save CSV for {source_file}: {e}')
        messagebox.showinfo('Success', f'CSV files saved to {folder_path}')

    def save_temporary_results(self):
        import csv
        from tkinter import filedialog
        if not hasattr(self, 'result_tree') or not self.result_tree.get_children():
            messagebox.showwarning('Warning', 'Nessun risultato da salvare.')
            return
        file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV Files', '*.csv')], title='Salva i risultati temporanei come')
        if not file_path:
            return
        sep = self.config.get('csv_separator', ',')
        dec_sep = self.config.get('decimal_separator', '.')
        thou_sep = self.config.get('thousands_separator', ',')
        # Gather current data from the Treeview
        tree_data = []
        for row_id in self.result_tree.get_children():
            values = self.result_tree.item(row_id)['values']
            # Ensure we have all columns (Item, Description, Lowest Price, Quantity, Source File)
            if len(values) < 5:
                values += [''] * (5 - len(values))
            # Format price according to config
            price = values[2]
            try:
                # price_float = float(str(price).replace(thou_sep, '').replace(dec_sep, '.'))
                price_str = f"{price}".replace(',', 'X').replace('.', dec_sep).replace('X', thou_sep)
            except Exception:
                price_str = str(price)
            tree_data.append({
                'item': values[0],
                'description': values[1],
                'price': price_str,
                'quantity': values[3],
                'file': values[4],
            })
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=sep)
                writer.writerow(['Item', 'Description', 'Lowest Price', 'Quantity', 'Source File'])
                for r in tree_data:
                    writer.writerow([r['item'], r['description'], r['price'], r['quantity'], r['file']])
            messagebox.showinfo('Success', f'Temporary results saved to {file_path}')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to save temporary results: {e}')

    def load_results_from_csv(self):
        import pandas as pd
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(filetypes=[('CSV Files', '*.csv')], title='Carica i risultati da CSV')
        if not file_path:
            return
        sep = self.config.get('csv_separator', ',')
        dec_sep = self.config.get('decimal_separator', '.')
        thou_sep = self.config.get('thousands_separator', ',')
        try:
            df = pd.read_csv(file_path, delimiter=sep, dtype=str)
            # Ensure columns are in the expected order and format
            if 'Item' in df.columns and 'Description' in df.columns and 'Lowest Price' in df.columns:
                # Parse price column to string formatted according to config
                def format_price(val):
                    if pd.isna(val) or str(val).strip() == '':
                        return ''
                    try:
                        return  float(str(val).replace(thou_sep, '').replace(dec_sep, '.'))
                        # return f"{price_float:,.4f}".replace(',', 'X').replace('.', dec_sep).replace('X', thou_sep)
                    except Exception:
                        return str(val)
                df['Lowest Price'] = df['Lowest Price'].apply(format_price)
                # Rename columns to match display_results expectations
                df = df.rename(columns={
                    'Item': 'item',
                    'Description': 'description',
                    'Lowest Price': 'price',
                    'Quantity': 'quantity',
                    'Source File': 'file'
                })
                # Ensure quantity is empty string if null/empty/NaN
                import numpy as np
                df['quantity'] = df['quantity'].apply(lambda x: '' if pd.isna(x) or str(x).strip().lower() in ('', 'nan') else x)
                self.display_results(df.to_dict('records'))
                messagebox.showinfo('Success', f'Results loaded from {file_path}')
                # Enable menu items for saving
                self.file_menu.entryconfig('Salva risultati come CSV', state='normal')
                self.file_menu.entryconfig('Salva risultati temporanei', state='normal')
            else:
                messagebox.showwarning('Warning', 'CSV file does not contain expected columns (Item, Description, Lowest Price).')
        except Exception as e:
            messagebox.showerror('Error', f'Impossibile caricare i risultati da CSV: {e}')

    def open_config_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title('Configurazione')
        dialog.grab_set()
        tk.Label(dialog, text='Separatore colonne CSV:').pack(pady=5)
        sep_var = tk.StringVar(dialog)
        sep_var.set(self.config.get('csv_separator', ','))
        sep_entry = ttk.Entry(dialog, textvariable=sep_var, width=5)
        sep_entry.pack(pady=5)
        tk.Label(dialog, text='Separatore decimali:').pack(pady=5)
        dec_var = tk.StringVar(dialog)
        dec_var.set(self.config.get('decimal_separator', '.'))
        dec_entry = ttk.Entry(dialog, textvariable=dec_var, width=5)
        dec_entry.pack(pady=5)
        tk.Label(dialog, text='Separatore migliaia:').pack(pady=5)
        thou_var = tk.StringVar(dialog)
        thou_var.set(self.config.get('thousands_separator', ','))
        thou_entry = ttk.Entry(dialog, textvariable=thou_var, width=5)
        thou_entry.pack(pady=5)
        def on_save():
            self.config['csv_separator'] = sep_var.get()
            self.config['decimal_separator'] = dec_var.get()
            self.config['thousands_separator'] = thou_var.get()
            self.save_config()
            dialog.destroy()
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text='Salva', command=on_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text='Annulla', command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        dialog.wait_window()

    def _update_total_label(self):
        # Calculate the total (prezzo*quantità) grouped by listino (source file)
        totals = {}
        for row_id in self.result_tree.get_children():
            values = self.result_tree.item(row_id)['values']
            try:
                prezzo = float(str(values[2]).replace(',', '.').replace(' ', ''))
                quantita = float(str(values[3]).replace(',', '.').replace(' ', '')) if str(values[3]).strip() else 0.0
                listino = str(values[4]) if len(values) > 4 else 'Listino sconosciuto'
                if listino not in totals:
                    totals[listino] = 0.0
                totals[listino] += prezzo * quantita
            except Exception:
                continue
        if totals:
            lines = [f"Totale per {listino}: {tot:,.2f}" for listino, tot in totals.items()]
            self.total_var.set("\n".join(lines))
        else:
            self.total_var.set("")

if __name__ == '__main__':
    root = tk.Tk()
    app = PriceCompareApp(root)
    root.mainloop() 