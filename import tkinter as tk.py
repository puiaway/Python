import os
import tkinter as tk
from tkinter import filedialog, messagebox
import csv

class LogTextSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text and Log File Search with CSV Export")
        self.all_lines = []
        self.matches = []

        self.create_widgets()

    def create_widgets(self):
        tk.Button(self.root, text="Select Folder", command=self.select_folder).pack(pady=5)

        self.search_var = tk.StringVar()
        tk.Entry(self.root, textvariable=self.search_var, width=60).pack(pady=5)

        tk.Button(self.root, text="Search", command=self.perform_search).pack(pady=5)
        tk.Button(self.root, text="Export to CSV", command=self.export_csv).pack(pady=5)

        self.result_listbox = tk.Listbox(self.root, width=100, height=25)
        self.result_listbox.pack(pady=5)

        self.status_label = tk.Label(self.root, text="No folder selected.")
        self.status_label.pack(pady=5)

    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        if not self.folder_path:
            return

        self.load_text_and_log_files()
        self.status_label.config(text=f"Loaded {len(self.all_lines)} lines from {self.folder_path}")

    def load_text_and_log_files(self):
        self.all_lines.clear()
        for root_dir, _, files in os.walk(self.folder_path):
            for file in files:
                if file.lower().endswith(('.txt', '.log')):
                    file_path = os.path.join(root_dir, file)
                    try:
                        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                            for idx, line in enumerate(f, 1):
                                self.all_lines.append((file, idx, line.strip()))
                    except Exception as e:
                        print(f"Error reading {file_path}: {e}")

    def perform_search(self):
        keyword = self.search_var.get().lower()
        self.matches = [
            (fname, line_num, text)
            for fname, line_num, text in self.all_lines
            if keyword in text.lower()
        ]

        self.result_listbox.delete(0, tk.END)
        for fname, line_num, text in self.matches:
            self.result_listbox.insert(tk.END, f"{fname} [Line {line_num}]: {text}")

        self.status_label.config(text=f"Found {len(self.matches)} matches.")

    def export_csv(self):
        if not self.matches:
            messagebox.showinfo("No Matches", "No matches to export.")
            return

        try:
            with open("search_results.csv", "w", newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["File Name", "Line Number", "Text"])
                writer.writerows(self.matches)
            messagebox.showinfo("Export Complete", "Results saved to search_results.csv")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = LogTextSearchApp(root)
    root.mainloop()
