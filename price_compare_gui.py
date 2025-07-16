import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import json
import os

CONFIG_FILE = 'config.json'

class PriceCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Price Compare Application')
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
            messagebox.showerror('Error', f'Failed to save config: {e}')

    def setup_ui(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        # Config button at top right
        self.config_btn = ttk.Button(top_frame, text='Config', command=self.open_config_dialog)
        self.config_btn.pack(side=tk.RIGHT)
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        self.select_btn = ttk.Button(frame, text='Select Excel Files', command=self.select_files)
        self.select_btn.pack(pady=5)

        self.files_label = ttk.Label(frame, text='No files selected.')
        self.files_label.pack(pady=5)

        self.result_box = tk.Text(frame, height=15, width=60)
        self.result_box.pack(pady=5)
        self.result_box.insert(tk.END, 'Results will appear here.')
        self.result_box.config(state=tk.DISABLED)

        self.save_btn = ttk.Button(frame, text='Save Results as CSV', command=self.save_results, state=tk.DISABLED)
        self.save_btn.pack(pady=5)

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx;*.xls')])
        if files:
            self.files = list(files)
            self.files_label.config(text=f'Selected files: {len(self.files)}')
            self.file_column_mappings = []  # List of dicts: {file, item_col, price_col, header_idx}
            for file in self.files:
                try:
                    # Pass a dummy columns list, will be ignored
                    item_col, price_col, description_col, header_idx = self.ask_column_mapping_with_header(file)
                    if item_col is None or price_col is None or description_col is None:
                        messagebox.showwarning('Warning', f'Skipping file: {file}')
                        continue
                    self.file_column_mappings.append({
                        'file': file,
                        'item_col': item_col,
                        'price_col': price_col,
                        'description_col': description_col,
                        'header_idx': header_idx
                    })
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to read {file}: {e}')
            # Placeholder: Next, process files and compare prices
            if self.file_column_mappings:
                self.compare_and_display()
            else:
                self.result_box.config(state=tk.NORMAL)
                self.result_box.delete(1.0, tk.END)
                self.result_box.insert(tk.END, 'No valid files/columns selected.')
                self.result_box.config(state=tk.DISABLED)
                self.save_btn.config(state=tk.DISABLED)
        else:
            self.files_label.config(text='No files selected.')

    def ask_column_mapping_with_header(self, file):
        import pandas as pd
        # Read the first 20 rows for preview
        try:
            preview_df = pd.read_excel(file, nrows=20, header=None)
        except Exception as e:
            messagebox.showerror('Error', f'Failed to preview {file}: {e}')
            return None, None, None, None
        dialog = tk.Toplevel(self.root)
        dialog.title(f'Select header row for {file.split("/")[-1]}')
        dialog.grab_set()
        tk.Label(dialog, text=f'Select the header row for file:\n{file}').pack(pady=5)
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
        dialog2.title(f'Select columns for {file.split("/")[-1]}')
        dialog2.grab_set()
        tk.Label(dialog2, text=f'Select columns for file:\n{file}').pack(pady=5)
        tk.Label(dialog2, text='Item column:').pack()
        item_var = tk.StringVar(dialog2)
        item_var.set(columns[0] if columns else '')
        item_menu = ttk.Combobox(dialog2, textvariable=item_var, values=columns, state='readonly')
        item_menu.pack(pady=2)
        tk.Label(dialog2, text='Price column:').pack()
        price_var = tk.StringVar(dialog2)
        price_var.set(columns[0] if columns else '')
        price_menu = ttk.Combobox(dialog2, textvariable=price_var, values=columns, state='readonly')
        price_menu.pack(pady=2)
        tk.Label(dialog2, text='Description column:').pack()
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
                    original_item = str(row[mapping['item_col']])
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
                messagebox.showerror('Error', f'Failed to process {mapping["file"]}: {e}')
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
        self.save_btn.config(state=tk.NORMAL if results else tk.DISABLED)

    def display_results(self, results):
        # Remove old result widgets if any
        if hasattr(self, 'result_tree') and self.result_tree:
            self.result_tree.destroy()
        if hasattr(self, 'tree_scroll_x') and self.tree_scroll_x:
            self.tree_scroll_x.destroy()
        if hasattr(self, 'tree_scroll_y') and self.tree_scroll_y:
            self.tree_scroll_y.destroy()
        if hasattr(self, 'result_frame') and self.result_frame:
            self.result_frame.destroy()
        # If no results, show a message
        if not results:
            self.result_box.config(state=tk.NORMAL)
            self.result_box.delete(1.0, tk.END)
            self.result_box.insert(tk.END, 'No results to display.')
            self.result_box.config(state=tk.DISABLED)
            return
        # Hide the text box
        self.result_box.config(state=tk.NORMAL)
        self.result_box.delete(1.0, tk.END)
        self.result_box.config(state=tk.DISABLED)
        self.result_box.pack_forget()
        # Sort results by description
        # results_sorted = sorted(results, key=lambda r: str(r['description']).lower() if r['description'] is not None else '')
        results_sorted = results
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
        for idx, r in enumerate(results_sorted):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            self.result_tree.insert('', 'end', values=[
                r['item'],
                r['description'],
                r['price'],
                '',  # Quantity column is empty
                r['file'].split('/')[-1]
            ], tags=(tag,))
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

    def save_results(self):
        import csv
        from tkinter import filedialog
        if not hasattr(self, 'comparison_results') or not self.comparison_results:
            messagebox.showwarning('Warning', 'No results to save.')
            return
        # Let user select destination folder
        folder_path = filedialog.askdirectory(title='Select destination folder for CSV files')
        if not folder_path:
            return
        # Group results by source file
        results_by_file = {}
        for r in self.comparison_results:
            source_file = r['file']
            if source_file not in results_by_file:
                results_by_file[source_file] = []
            results_by_file[source_file].append(r)
        # Save one CSV per source Excel file
        sep = self.config.get('csv_separator', ',')
        dec_sep = self.config.get('decimal_separator', '.')
        thou_sep = self.config.get('thousands_separator', ',')
        # Save combined CSV as well
        combined_csv_path = os.path.join(folder_path, 'result_compared.csv')
        try:
            with open(combined_csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=sep)
                writer.writerow(['Item', 'Description', 'Lowest Price', 'Quantity', 'Source File'])
                for r in self.comparison_results:
                    price = r['price']
                    if isinstance(price, float):
                        price_str = f"{price:,.4f}".replace(',', 'X').replace('.', dec_sep).replace('X', thou_sep)
                    else:
                        price_str = str(price)
                    writer.writerow([r['item'], r['description'], price_str, '', r['file'].split('/')[-1]])
        except Exception as e:
            messagebox.showerror('Error', f'Failed to save combined CSV: {e}')
        for source_file, rows in results_by_file.items():
            base_name = os.path.splitext(os.path.basename(source_file))[0]
            csv_name = f"{base_name}_compared.csv"
            csv_path = os.path.join(folder_path, csv_name)
            try:
                with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter=sep)
                    writer.writerow(['Item', 'Description', 'Lowest Price', 'Quantity'])
                    for r in rows:
                        price = r['price']
                        if isinstance(price, float):
                            price_str = f"{price:,.4f}".replace(',', 'X').replace('.', dec_sep).replace('X', thou_sep)
                        else:
                            price_str = str(price)
                        writer.writerow([r['item'], r['description'], price_str, ''])
            except Exception as e:
                messagebox.showerror('Error', f'Failed to save CSV for {source_file}: {e}')
        messagebox.showinfo('Success', f'CSV files saved to {folder_path}')

    def open_config_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title('Configuration')
        dialog.grab_set()
        tk.Label(dialog, text='CSV Column Separator:').pack(pady=5)
        sep_var = tk.StringVar(dialog)
        sep_var.set(self.config.get('csv_separator', ','))
        sep_entry = ttk.Entry(dialog, textvariable=sep_var, width=5)
        sep_entry.pack(pady=5)
        tk.Label(dialog, text='Decimal Separator:').pack(pady=5)
        dec_var = tk.StringVar(dialog)
        dec_var.set(self.config.get('decimal_separator', '.'))
        dec_entry = ttk.Entry(dialog, textvariable=dec_var, width=5)
        dec_entry.pack(pady=5)
        tk.Label(dialog, text='Thousands Separator:').pack(pady=5)
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
        ttk.Button(btn_frame, text='Save', command=on_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text='Cancel', command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        dialog.wait_window()

if __name__ == '__main__':
    root = tk.Tk()
    app = PriceCompareApp(root)
    root.mainloop() 