import os
import csv
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

class TextSearchApp:
    def __init__(self, master):
        self.master = master
        master.title("Text Search in Folder (Export CSV/Excel + Match Once Option)")
        master.geometry("930x700")

        # Folder Selection
        tk.Label(master, text="Folder:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.folder_entry = tk.Entry(master, width=60)
        self.folder_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(master, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5)

        # Keyword Entry
        tk.Label(master, text="Keyword:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.keyword_entry = tk.Entry(master, width=60)
        self.keyword_entry.grid(row=1, column=1, padx=5, pady=5)

        # Match Once Option
        self.match_once_var = tk.BooleanVar()
        self.match_once_check = tk.Checkbutton(master, text="Match once per file only", variable=self.match_once_var)
        self.match_once_check.grid(row=2, column=1, sticky='w', padx=5)

        # Buttons
        self.search_btn = tk.Button(master, text="Search", command=self.start_search_thread)
        self.search_btn.grid(row=1, column=2, padx=5)

        self.cancel_btn = tk.Button(master, text="Cancel", command=self.cancel_search, state=tk.DISABLED)
        self.cancel_btn.grid(row=2, column=2, padx=5)

        # Progress and Status
        self.progress = ttk.Progressbar(master, orient="horizontal", length=600, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=2, padx=10, pady=5)

        self.status_label = tk.Label(master, text="Status: Ready")
        self.status_label.grid(row=4, column=0, columnspan=3, sticky="w", padx=10)

        self.export_csv_btn = tk.Button(master, text="Export to CSV", command=self.export_csv, state=tk.DISABLED)
        self.export_csv_btn.grid(row=4, column=2, sticky="e", padx=10)

        self.export_excel_btn = tk.Button(master, text="Export to Excel", command=self.export_excel, state=tk.DISABLED)
        self.export_excel_btn.grid(row=4, column=1, sticky="e", padx=10)

        self.result_area = scrolledtext.ScrolledText(master, width=110, height=30)
        self.result_area.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

        self.stop_flag = threading.Event()
        self.temp_csv_path = None

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)

    def start_search_thread(self):
        self.stop_flag.clear()
        self.search_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.export_csv_btn.config(state=tk.DISABLED)
        self.export_excel_btn.config(state=tk.DISABLED)
        self.result_area.delete(1.0, tk.END)
        self.status_label.config(text="Status: Searching...")
        self.progress["value"] = 0

        thread = threading.Thread(target=self.search_files)
        thread.start()

    def cancel_search(self):
        self.stop_flag.set()
        self.status_label.config(text="Status: Cancelling...")

    def search_files(self):
        folder = self.folder_entry.get()
        keyword = self.keyword_entry.get().strip()
        match_once = self.match_once_var.get()

        if not os.path.isdir(folder) or not keyword:
            messagebox.showerror("Error", "Please select a valid folder and enter a keyword.")
            self.search_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)
            return

        file_list = []
        for root, dirs, files in os.walk(folder):
            for file in files:
                if file.lower().endswith(('.txt', '.log', '.csv', '.json', '.xml')):
                    file_list.append(os.path.join(root, file))

        total_files = len(file_list)
        self.progress["maximum"] = total_files

        # Create temp file to write all results
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".csv", mode='w', encoding='utf-8', newline='')
        self.temp_csv_path = temp_file.name
        writer = csv.writer(temp_file)
        writer.writerow(["Filename", "Line Number", "Line"])

        preview_results = []
        match_count = 0

        for index, file_path in enumerate(file_list, 1):
            if self.stop_flag.is_set():
                self.status_label.config(text="Status: Search cancelled.")
                break

            filename = os.path.basename(file_path)
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    for i, line in enumerate(f, 1):
                        if keyword in line:
                            match_count += 1
                            writer.writerow([filename, i, line.strip()])
                            if len(preview_results) < 1000:
                                preview_results.append((filename, i, line.strip()))
                            if match_once:
                                break
            except Exception as e:
                writer.writerow([filename, 0, f"[Error: {str(e)}]"])
                if len(preview_results) < 1000:
                    preview_results.append((filename, 0, f"[Error: {str(e)}]"))

            if index % 10 == 0:
                self.progress["value"] = index
                self.status_label.config(text=f"Scanning: {index}/{total_files} files")
                self.master.update_idletasks()

        temp_file.close()

        if not self.stop_flag.is_set():
            self.status_label.config(text=f"Completed: {match_count} matches in {total_files} files")

        self.show_preview(preview_results, match_count)
        self.search_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.export_csv_btn.config(state=tk.NORMAL if match_count else tk.DISABLED)
        self.export_excel_btn.config(state=tk.NORMAL if match_count else tk.DISABLED)

    def show_preview(self, preview, total_count):
        self.result_area.delete(1.0, tk.END)
        for file, line_num, line in preview:
            self.result_area.insert(tk.END, f"{file} (Line {line_num}): {line}\n")

        if total_count > len(preview):
            self.result_area.insert(tk.END, f"\n--- Showing first {len(preview)} of {total_count} results ---\n")
            self.result_area.insert(tk.END, f"Please use 'Export to CSV' or 'Export to Excel' to see all matches.\n")

        self.result_area.yview_moveto(0)

    def export_csv(self):
        if not self.temp_csv_path or not os.path.exists(self.temp_csv_path):
            messagebox.showwarning("No Data", "No results to export.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV Files", "*.csv")])
        if path:
            try:
                with open(self.temp_csv_path, 'r', encoding='utf-8') as src, open(path, 'w', encoding='utf-8', newline='') as dst:
                    dst.write(src.read())
                messagebox.showinfo("Success", f"Exported results to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to write CSV:\n{e}")

    def export_excel(self):
        if not self.temp_csv_path or not os.path.exists(self.temp_csv_path):
            messagebox.showwarning("No Data", "No results to export.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel Files", "*.xlsx")])
        if path:
            try:
                import openpyxl
                from openpyxl.styles import Font

                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Search Results"

                with open(self.temp_csv_path, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    for r, row in enumerate(reader, 1):
                        for c, value in enumerate(row, 1):
                            cell = ws.cell(row=r, column=c, value=value)
                            if r == 1:
                                cell.font = Font(bold=True)

                wb.save(path)
                messagebox.showinfo("Success", f"Exported results to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export Excel:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TextSearchApp(root)
    root.mainloop()
