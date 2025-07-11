import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class PriceCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Price Compare Application')
        self.files = []
        self.setup_ui()

    def setup_ui(self):
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
                    item_col, price_col, header_idx = self.ask_column_mapping_with_header(file)
                    if item_col is None or price_col is None:
                        messagebox.showwarning('Warning', f'Skipping file: {file}')
                        continue
                    self.file_column_mappings.append({
                        'file': file,
                        'item_col': item_col,
                        'price_col': price_col,
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
            return None, None, None
        dialog = tk.Toplevel(self.root)
        dialog.title(f'Select header row for {file.split("/")[-1]}')
        dialog.grab_set()
        tk.Label(dialog, text=f'Select the header row for file:\n{file}').pack(pady=5)
        listbox = tk.Listbox(dialog, width=80, height=15)
        for idx, row in preview_df.iterrows():
            display = ' | '.join([str(x) for x in row.values])
            listbox.insert(tk.END, f'Row {idx+1}: {display}')
        listbox.pack(pady=5)
        listbox.selection_set(0)
        header_row_idx = [0]
        def on_header_select(event=None):
            sel = listbox.curselection()
            if sel:
                header_row_idx[0] = sel[0]
        listbox.bind('<<ListboxSelect>>', on_header_select)
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
        price_var.set(columns[1] if len(columns) > 1 else (columns[0] if columns else ''))
        price_menu = ttk.Combobox(dialog2, textvariable=price_var, values=columns, state='readonly')
        price_menu.pack(pady=2)
        result = {'item': None, 'price': None}
        def on_ok():
            result['item'] = item_var.get()
            result['price'] = price_var.get()
            dialog2.destroy()
        def on_cancel():
            dialog2.destroy()
        btn_frame2 = ttk.Frame(dialog2)
        btn_frame2.pack(pady=5)
        ttk.Button(btn_frame2, text='OK', command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame2, text='Cancel', command=on_cancel).pack(side=tk.LEFT, padx=5)
        dialog2.wait_window()
        return result['item'], result['price'], header_idx

    def compare_and_display(self):
        # Read all items and prices
        all_items = []  # List of dicts: {item, price, file, original_item}
        for mapping in self.file_column_mappings:
            try:
                df = pd.read_excel(
                    mapping['file'],
                    usecols=[mapping['item_col'], mapping['price_col']],
                    header=mapping['header_idx']
                )
                for _, row in df.iterrows():
                    original_item = str(row[mapping['item_col']])
                    item_key = original_item.strip().lower()
                    price_cell = row[mapping['price_col']]
                    # If price is missing or empty, treat as zero
                    if pd.isna(price_cell) or str(price_cell).strip() == '':
                        price = 0.0
                    else:
                        try:
                            price = float(price_cell)
                        except Exception:
                            price = 0.0
                    all_items.append({'item_key': item_key, 'price': price, 'file': mapping['file'], 'original_item': original_item})
            except Exception as e:
                messagebox.showerror('Error', f'Failed to process {mapping["file"]}: {e}')
        # Find lowest price for each item (case-insensitive, stripped)
        best_prices = {}
        for entry in all_items:
            item_key = entry['item_key']
            if item_key not in best_prices or entry['price'] < best_prices[item_key]['price']:
                best_prices[item_key] = {'price': entry['price'], 'file': entry['file'], 'original_item': entry['original_item']}
        # Prepare results for display
        results = []
        for item_key, data in sorted(best_prices.items()):
            results.append({'item': data['original_item'], 'price': data['price'], 'file': data['file']})
        self.display_results(results)
        self.comparison_results = results  # Save for CSV export
        self.save_btn.config(state=tk.NORMAL if results else tk.DISABLED)

    def display_results(self, results):
        self.result_box.config(state=tk.NORMAL)
        self.result_box.delete(1.0, tk.END)
        if not results:
            self.result_box.insert(tk.END, 'No results to display.')
        else:
            self.result_box.insert(tk.END, f'Item\tLowest Price\tSource File\n')
            for r in results:
                self.result_box.insert(tk.END, f"{r['item']}\t{r['price']}\t{r['file'].split('/')[-1]}\n")
        self.result_box.config(state=tk.DISABLED)

    def save_results(self):
        if not hasattr(self, 'comparison_results') or not self.comparison_results:
            messagebox.showwarning('Warning', 'No results to save.')
            return
        file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV Files', '*.csv')])
        if not file_path:
            return
        import csv
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Item', 'Lowest Price', 'Source File'])
                for r in self.comparison_results:
                    writer.writerow([r['item'], r['price'], r['file'].split('/')[-1]])
            messagebox.showinfo('Success', f'Results saved to {file_path}')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to save CSV: {e}')

if __name__ == '__main__':
    root = tk.Tk()
    app = PriceCompareApp(root)
    root.mainloop() 