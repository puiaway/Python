
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import chardet

# Helper function to detect encoding
def detect_encoding(file_path, sample_size=10000):
    with open(file_path, 'rb') as f:
        raw_data = f.read(sample_size)
    return chardet.detect(raw_data)['encoding'] or 'utf-8'

# Function to search for keyword in file
def search_in_file(file_path, keyword, case_sensitive):
    matches = []
    try:
        encoding = detect_encoding(file_path)
        with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
            for line_num, line in enumerate(f, 1):
                line_to_check = line if case_sensitive else line.lower()
                if keyword in line_to_check:
                    matches.append((file_path, line_num, line.strip()))
    except Exception as e:
        matches.append((file_path, "ERROR", f"Could not read file: {e}"))
    return matches

# GUI Application
class KeywordSearchApp:
    def __init__(self, master):
        self.master = master
        master.title("Keyword Search in Folder")
        master.geometry("900x600")

        self.keyword = tk.StringVar()
        self.folder_path = tk.StringVar()
        self.case_sensitive = tk.BooleanVar()

        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.master, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Folder:").grid(row=0, column=0, sticky="w")
        folder_entry = ttk.Entry(frame, textvariable=self.folder_path, width=70)
        folder_entry.grid(row=0, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.browse_folder).grid(row=0, column=2)

        ttk.Label(frame, text="Keyword:").grid(row=1, column=0, sticky="w")
        keyword_entry = ttk.Entry(frame, textvariable=self.keyword, width=70)
        keyword_entry.grid(row=1, column=1, padx=5)

        ttk.Checkbutton(frame, text="Case Sensitive", variable=self.case_sensitive).grid(row=1, column=2)

        ttk.Button(frame, text="Search", command=self.search).grid(row=2, column=1, pady=10)

        self.result_area = scrolledtext.ScrolledText(frame, wrap=tk.WORD, width=100, height=25)
        self.result_area.grid(row=3, column=0, columnspan=3, pady=10)

        self.result_area.tag_configure("highlight", background="yellow")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)

    def search(self):
        folder = self.folder_path.get()
        keyword = self.keyword.get()
        case_sensitive = self.case_sensitive.get()

        if not folder or not keyword:
            messagebox.showerror("Error", "Please specify both folder and keyword.")
            return

        self.result_area.delete(1.0, tk.END)
        keyword_to_find = keyword if case_sensitive else keyword.lower()

        for root, _, files in os.walk(folder):
            for file in files:
                file_path = os.path.join(root, file)
                matches = search_in_file(file_path, keyword_to_find, case_sensitive)
                for file_path, line_num, line in matches:
                    start_index = self.result_area.index(tk.END)
                    self.result_area.insert(tk.END, f"{file_path} (Line {line_num}): {line}\n")
                    end_index = self.result_area.index(tk.END)
                    if line_num != "ERROR":
                        self.highlight_keyword(start_index, end_index, keyword)

    def highlight_keyword(self, start, end, keyword):
        idx = self.result_area.search(keyword, start, stopindex=end, nocase=not self.case_sensitive.get())
        while idx:
            lastidx = f"{idx}+{len(keyword)}c"
            self.result_area.tag_add("highlight", idx, lastidx)
            idx = self.result_area.search(keyword, lastidx, stopindex=end, nocase=not self.case_sensitive.get())

if __name__ == "__main__":
    root = tk.Tk()
    app = KeywordSearchApp(root)
    root.mainloop()
