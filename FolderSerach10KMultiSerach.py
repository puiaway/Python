import os
import csv
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


class TextSearchApp:
    def __init__(self, master):
        self.master = master
        master.title("Text Search in Folder (Multi-keyword, Cancel, CSV)")
        master.geometry("820x640")

        # Folder selection
        tk.Label(master, text="Folder:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.folder_entry = tk.Entry(master, width=60)
        self.folder_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(master, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5)

        # Keyword input (supports comma-separated)
        tk.Label(master, text="Keywords (comma-separated):").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.keyword_entry = tk.Entry(master, width=60)
        self.keyword_entry.grid(row=1, column=1, padx=5, pady=5)

        # Buttons
        self.search_btn = tk.Button(master, text="Search", command=self.start_search_thread)
        self.search_btn.grid(row=1, column=2, padx=5)

        self.cancel_btn = tk.Button(master, text="Cancel", command=self.cancel_search, state=tk.DISABLED)
        self.cancel_btn.grid(row=2, column=2, padx=5)

        # Progress bar and status
        self.progress = ttk.Progressbar(master, orient="horizontal", length=600, mode="determinate")
        self.progress.grid(row=2, column=0, columnspan=2, padx=10, pady=5)

        self.status_label = tk.Label(master, text="Status: Ready")
        self.status_label.grid(row=3, column=0, columnspan=3, sticky="w", padx=10)

        # Export button
        self.export_btn = tk.Button(master, text="Export to CSV", command=self.export_csv, state=tk.DISABLED)
        self.export_btn.grid(row=3, column=2, sticky="e", padx=10)

        # Result display
        self.result_area = scrolledtext.ScrolledText(master, width=100, height=30)
        self.result_area.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

        self.results = []
        self.stop_flag = threading.Event()

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)

    def start_search_thread(self):
        self.stop_flag.clear()
        self.search_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        self.result_area.delete(1.0, tk.END)
        self.status_label.config(text="Status: Searching...")
        self.progress["value"] = 0
        self.results.clear()

        thread = threading.Thread(target=self.search_files)
        thread.start()

    def cancel_search(self):
        self.stop_flag.set()
        self.status_label.config(text="Status: Cancelling...")

    def search_files(self):
        folder = self.folder_entry.get()
        raw_keywords = self.keyword_entry.get().strip()

        if not os.path.isdir(folder) or not raw_keywords:
            messagebox.showerror("Error", "Please select a valid folder and enter keywords.")
            self.search_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)
            return

        keywords = [k.strip().lower() for k in raw_keywords.split(",") if k.strip()]
        if not keywords:
            messagebox.showerror("Error", "No valid keywords found.")
            self.search_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)
            return

        file_list = []
        for root, _, files in os.walk(folder):
            for file in files:
                if file.lower().endswith(('.txt', '.log', '.csv', '.json', '.prn', '.jrn')):
                    file_list.append(os.path.join(root, file))

        total_files = len(file_list)
        self.progress["maximum"] = total_files

        for index, file_path in enumerate(file_list, 1):
            if self.stop_flag.is_set():
                self.status_label.config(text="Status: Search cancelled.")
                break

            filename = os.path.basename(file_path)
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    for i, line in enumerate(f, 1):
                        if any(keyword in line.lower() for keyword in keywords):
                            self.results.append((filename, i, line.strip()))
            except Exception as e:
                self.results.append((filename, 0, f"[Error: {str(e)}]"))

            self.progress["value"] = index
            self.status_label.config(text=f"Scanning: {index}/{total_files} files")
            self.master.update_idletasks()

        if not self.stop_flag.is_set():
            self.status_label.config(text=f"Completed: {len(self.results)} matches in {total_files} files")

        self.show_results()
        self.search_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.NORMAL if self.results else tk.DISABLED)

    def show_results(self):
        self.result_area.delete(1.0, tk.END)
        for file, line_num, line in self.results:
            self.result_area.insert(tk.END, f"{file} (Line {line_num}): {line}\n")
        self.result_area.yview_moveto(0)

    def export_csv(self):
        if not self.results:
            messagebox.showwarning("No Data", "No results to export.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV Files", "*.csv")])
        if path:
            try:
                with open(path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Filename", "Line Number", "Line"])
                    writer.writerows(self.results)
                messagebox.showinfo("Success", f"Exported {len(self.results)} rows to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to write CSV:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TextSearchApp(root)
    root.mainloop()
