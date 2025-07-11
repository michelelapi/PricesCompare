import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import camelot
import pandas as pd
import threading

class PDFtoExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title('PDF to Excel Converter')
        self.tables = []
        self.setup_ui()

    def setup_ui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        self.select_btn = ttk.Button(frame, text='Select PDF File', command=self.select_pdf)
        self.select_btn.pack(pady=5)

        self.progress = ttk.Progressbar(frame, orient='horizontal', length=300, mode='determinate')
        self.progress.pack(pady=5)
        self.progress_label = ttk.Label(frame, text='')
        self.progress_label.pack(pady=2)

        self.tables_label = ttk.Label(frame, text='No PDF loaded.')
        self.tables_label.pack(pady=5)

        self.export_btn = ttk.Button(frame, text='Export All Tables to Excel', command=self.export_excel, state=tk.DISABLED)
        self.export_btn.pack(pady=5)

    def select_pdf(self):
        pdf_path = filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')])
        if not pdf_path:
            return
        # Always use 'stream' flavor and all pages
        import PyPDF2
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            total_pages = len(reader.pages)
        pages = list(range(1, total_pages + 1))
        self.progress['value'] = 0
        self.progress['maximum'] = len(pages)
        self.progress_label.config(text='Starting extraction...')
        self.tables_label.config(text='')
        self.export_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.extract_tables_with_progress, args=(pdf_path, pages), daemon=True).start()

    def extract_tables_with_progress(self, pdf_path, pages):
        all_tables = []
        for idx, page in enumerate(pages):
            try:
                tables = camelot.read_pdf(pdf_path, pages=str(page), flavor='stream')
                all_tables.extend(tables)
            except Exception as e:
                pass  # skip errors for individual pages
            percent = int(((idx + 1) / len(pages)) * 100)
            self.root.after(0, self.update_progress, idx + 1, len(pages), percent)
        self.tables = all_tables
        self.root.after(0, self.show_tables, pdf_path)

    def update_progress(self, current, total, percent):
        self.progress['value'] = current
        self.progress_label.config(text=f'Processed {current}/{total} pages ({percent}%)')

    def show_tables(self, pdf_path):
        if not self.tables or len(self.tables) == 0:
            self.tables_label.config(text='No tables found.')
            self.progress_label.config(text='Done. No tables found.')
            self.export_btn.config(state=tk.DISABLED)
            return
        self.pdf_path = pdf_path
        self.tables_label.config(text=f'{len(self.tables)} tables found in {pdf_path.split("/")[-1]}:')
        self.export_btn.config(state=tk.NORMAL)
        self.progress_label.config(text='Done.')
        self.progress['value'] = self.progress['maximum']

    def export_excel(self):
        if not self.tables or len(self.tables) == 0:
            messagebox.showwarning('No Tables', 'No tables to export.')
            return
        save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])
        if not save_path:
            return
        try:
            # Concatenate all tables, separated by a blank row
            all_dfs = []
            for table in self.tables:
                df = table.df
                all_dfs.append(df)
                # Add a blank row for separation
                all_dfs.append(pd.DataFrame([[''] * df.shape[1]], columns=df.columns))
            if all_dfs:
                combined_df = pd.concat(all_dfs, ignore_index=True)
            else:
                combined_df = pd.DataFrame()
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='Tables', index=False, header=False)
            messagebox.showinfo('Success', f'All tables exported to {save_path}')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to export Excel: {e}')

if __name__ == '__main__':
    root = tk.Tk()
    app = PDFtoExcelApp(root)
    root.mainloop() 