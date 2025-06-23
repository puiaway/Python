import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


def search_text_in_files(folder_path, keyword):
    results = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(('.txt', '.log')):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        for i, line in enumerate(f, 1):
                            if keyword in line:
                                results.append((file, i, line.strip()))
                except Exception as e:
                    print(f"Error reading {file_path}: {e}")
    return results


def export_to_csv(results, save_path):
    with open(save_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(["File", "Line Number", "Matched Line"])
        writer.writerows(results)


class TextSearchApp:
    def __init__(self, master):
        self.master = master
        master.title("Text Search and Export to CSV")

        # Folder selection
        self.folder_label = tk.Label(master, text="Folder:")
        self.folder_label.grid(row=0, column=0, sticky="w")

        self.folder_path = tk.Entry(master, width=50)
        self.folder_path.grid(row=0, column=1)

        self.browse_button = tk.Button(master, text="Browse", command=self.browse_folder)
        self.browse_button.grid(row=0, column=2)

        # Keyword entry
        self.keyword_label = tk.Label(master, text="Keyword:")
        self.keyword_label.grid(row=1, column=0, sticky="w")

        self.keyword_entry = tk.Entry(master, width=50)
        self.keyword_entry.grid(row=1, column=1)

        # Search and Export buttons
        self.search_button = tk.Button(master, text="Search", command=self.run_search)
        self.search_button.grid(row=2, column=1, pady=10, sticky="w")

        self.export_button = tk.Button(master, text="Export to CSV", command=self.save_csv, state=tk.DISABLED)
        self.export_button.grid(row=2, column=1, pady=10, sticky="e")

        # Result area
        self.result_area = scrolledtext.ScrolledText(master, width=80, height=20)
        self.result_area.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

        self.results = []

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.delete(0, tk.END)
            self.folder_path.insert(0, path)

    def run_search(self):
        folder = self.folder_path.get()
        keyword = self.keyword_entry.get()

        if not os.path.isdir(folder) or not keyword:
            messagebox.showerror("Error", "Please specify a valid folder and keyword.")
            return

        self.result_area.delete(1.0, tk.END)
        self.results = search_text_in_files(folder, keyword)

        for file, line_num, line in self.results:
            self.result_area.insert(tk.END, f"{file} (Line {line_num}): {line}\n")

        self.export_button.config(state=tk.NORMAL if self.results else tk.DISABLED)

    def save_csv(self):
        if not self.results:
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")])
        if save_path:
            export_to_csv(self.results, save_path)
            messagebox.showinfo("Success", f"Results exported to {save_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TextSearchApp(root)
    root.mainloop()
