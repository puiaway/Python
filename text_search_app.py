import json
import logging
import threading
import platform
from pathlib import Path
import tempfile
import csv
import chardet
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# Configure logging
tlogging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s')

# History file for storing recent keywords
HISTORY_FILE = Path.home() / ".text_search_keywords.json"


def safe_path(path: Path) -> str:
    """
    Convert Path to a Windows long path if necessary, otherwise return the string path.
    """
    if platform.system() == 'Windows':
        resolved = path.resolve()
        pstr = str(resolved)
        if not pstr.startswith('\\\\?\\'):
            return '\\\\?\\' + pstr
        return pstr
    return str(path)


def load_history() -> list:
    """
    Load keyword history from JSON, back up if corrupted.
    """
    if HISTORY_FILE.exists():
        try:
            return json.loads(HISTORY_FILE.read_text(encoding='utf-8'))
        except json.JSONDecodeError:
            backup = HISTORY_FILE.with_suffix('.bak')
            HISTORY_FILE.rename(backup)
            logging.warning(
                "Corrupted history file. Backed up to %s",
                backup
            )
    return []


def save_history(history: list):
    """
    Save keyword history to JSON.
    """
    try:
        HISTORY_FILE.write_text(
            json.dumps(history, ensure_ascii=False, indent=2),
            encoding='utf-8'
        )
    except Exception as e:
        logging.error("Failed to save history: %s", e)


def detect_encoding(file_path: Path, cache: dict, sample_size=8192) -> str:
    """
    Detect encoding using chardet and cache by file extension.
    """
    ext = file_path.suffix.lower()
    if ext in cache:
        return cache[ext]
    try:
        with open(file_path, 'rb') as f:
            raw = f.read(sample_size)
        result = chardet.detect(raw)
        encoding = result.get('encoding') or 'utf-8'
    except Exception as e:
        logging.warning("Encoding detection failed for %s: %s", file_path, e)
        encoding = 'utf-8'
    cache[ext] = encoding
    return encoding


class TextSearchApp:
    def __init__(self, master):
        self.master = master
        master.title("Text Search in Folder (Export CSV/Excel + Match Options)")
        master.geometry("930x700")

        # Keyword history
        self.keyword_history = load_history()
        self.encoding_cache = {}

        # Folder selection
        tk.Label(master, text="Folder:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.folder_entry = tk.Entry(master, width=60)
        self.folder_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(master, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5)

        # Keyword input
        tk.Label(master, text="Keyword:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.keyword_entry = ttk.Combobox(master, width=58, values=self.keyword_history)
        self.keyword_entry.grid(row=1, column=1, padx=5, pady=5)

        # Match options
        self.match_once_var = tk.BooleanVar()
        self.match_once_check = tk.Checkbutton(
            master,
            text="Match once per file only",
            variable=self.match_once_var,
            command=self.on_match_once_toggled
        )
        self.match_once_check.grid(row=2, column=1, sticky='w', padx=5)

        self.last_match_var = tk.BooleanVar()
        self.last_match_check = tk.Checkbutton(
            master,
            text="Use last match per file",
            variable=self.last_match_var,
            command=self.on_last_match_toggled
        )
        self.last_match_check.grid(row=2, column=1, sticky='e', padx=5)

        self.case_sensitive_var = tk.BooleanVar()
        tk.Checkbutton(
            master,
            text="Case sensitive",
            variable=self.case_sensitive_var
        ).grid(row=3, column=1, sticky='w', padx=5, pady=2)

        # Buttons
        self.search_btn = tk.Button(master, text="Search", command=self.start_search_thread)
        self.search_btn.grid(row=1, column=2, padx=5)
        self.cancel_btn = tk.Button(master, text="Cancel", command=self.cancel_search, state=tk.DISABLED)
        self.cancel_btn.grid(row=2, column=2, padx=5)

        # Progress and status
        self.progress = ttk.Progressbar(master, orient="horizontal", length=600, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=2, padx=10, pady=5)

        self.status_label = tk.Label(master, text="Status: Ready")
        self.status_label.grid(row=5, column=0, columnspan=3, sticky="w", padx=10)

        # Export buttons
        self.export_csv_btn = tk.Button(
            master, text="Export to CSV", command=self.export_csv, state=tk.DISABLED
        )
        self.export_csv_btn.grid(row=5, column=2, sticky="e", padx=10)
        self.export_excel_btn = tk.Button(
            master, text="Export to Excel", command=self.export_excel, state=tk.DISABLED
        )
        self.export_excel_btn.grid(row=5, column=1, sticky="e", padx=10)

        # Results area
        self.result_area = scrolledtext.ScrolledText(master, width=110, height=30)
        self.result_area.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

        self.stop_flag = threading.Event()
        self.temp_csv_path = None

    def on_match_once_toggled(self):
        # Disable conflicting option
        if self.match_once_var.get():
            self.last_match_check.config(state=tk.DISABLED)
        else:
            self.last_match_check.config(state=tk.NORMAL)

    def on_last_match_toggled(self):
        # Disable conflicting option
        if self.last_match_var.get():
            self.match_once_check.config(state=tk.DISABLED)
        else:
            self.match_once_check.config(state=tk.NORMAL)

    def reset_ui(self):
        self.search_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.update_status("Status: Ready")

    def update_status(self, text: str):
        self.master.after(0, self.status_label.config, {'text': text})

    def update_progress(self, value: int):
        self.master.after(0, self.progress.config, {'value': value})

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)

    def start_search_thread(self):
        keyword = self.keyword_entry.get().strip()
        # Update history
        if keyword and keyword not in self.keyword_history:
            self.keyword_history.insert(0, keyword)
            self.keyword_history = self.keyword_history[:20]
            self.keyword_entry['values'] = self.keyword_history
            save_history(self.keyword_history)

        # Reset flags and UI
        self.stop_flag.clear()
        self.search_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.export_csv_btn.config(state=tk.DISABLED)
        self.export_excel_btn.config(state=tk.DISABLED)
        self.result_area.delete(1.0, tk.END)
        self.update_status("Status: Searching...")
        self.update_progress(0)

        threading.Thread(target=self.search_files, daemon=True).start()

    def cancel_search(self):
        self.stop_flag.set()
        self.update_status("Status: Cancelling...")

    def search_files(self):
        folder = Path(self.folder_entry.get())
        keyword = self.keyword_entry.get().strip()
        match_once = self.match_once_var.get()
        use_last_match = self.last_match_var.get()
        case_sensitive = self.case_sensitive_var.get()

        # Validate inputs
        if not folder.is_dir():
            messagebox.showerror("Error", "Please select a valid folder.")
            self.reset_ui()
            return
        if not keyword:
            messagebox.showerror("Error", "Please enter a keyword.")
            self.reset_ui()
            return
        if len(keyword) > 100:
            messagebox.showerror(
                "Error", "Keyword is too long. Please limit to 100 characters."
            )
            self.reset_ui()
            return

        keyword_cmp = keyword if case_sensitive else keyword.lower()

        # Gather files
        file_list = [
            Path(root) / file
            for root, _, files in os.walk(folder)
            for file in files
            if file.lower().endswith(('.txt', '.log', '.csv', '.json', '.xml'))
        ]

        total = len(file_list)
        self.master.after(0, self.progress.config, {'maximum': total})

        # Prepare CSV
        temp = tempfile.NamedTemporaryFile(
            delete=False, suffix=".csv", mode='w', encoding='utf-8', newline=''
        )
        self.temp_csv_path = Path(temp.name)
        writer = csv.writer(temp)
        writer.writerow(["Filename", "Line Number", "Line"])

        preview, count = [], 0
        for idx, path in enumerate(file_list, start=1):
            if self.stop_flag.is_set():
                self.update_status("Status: Search cancelled.")
                break

            fname = path.name
            try:
                enc = detect_encoding(path, self.encoding_cache)
                with open(safe_path(path), 'r', encoding=enc, errors='ignore', buffering=64*1024) as f:
                    last = None
                    for lineno, line in enumerate(f, 1):
                        txt = line if case_sensitive else line.lower()
                        if keyword_cmp in txt:
                            if match_once:
                                count += 1
                                writer.writerow([fname, lineno, line.strip()])
                                preview.append((fname, lineno, line.strip()))
                                break
                            if use_last_match:
                                last = (lineno, line.strip())
                            else:
                                count += 1
                                writer.writerow([fname, lineno, line.strip()])
                                preview.append((fname, lineno, line.strip()))
                    if use_last_match and last:
                        count += 1
                        writer.writerow([fname, last[0], last[1]])
                        preview.append((fname, last[0], last[1]))
            except (IOError, UnicodeDecodeError) as e:
                logging.warning("Error reading %s: %s", path, e)
                writer.writerow([fname, 0, f"[Error: {e}]"])
                preview.append((fname, 0, f"[Error: {e}]") )

            if idx % 10 == 0 or idx == total:
                self.update_progress(idx)
                self.update_status(f"Scanning: {idx}/{total} files")

        temp.close()
        if not self.stop_flag.is_set():
            self.update_status(f"Completed: {count} matches in {total} files")

        # Show preview and enable exports
        self.show_preview(preview, count)
        self.master.after(0, lambda: self.export_csv_btn.config(state=tk.NORMAL if count else tk.DISABLED))
        self.master.after(0, lambda: self.export_excel_btn.config(state=tk.NORMAL if count else tk.DISABLED))
        self.reset_ui()

    def show_preview(self, preview, total):
        self.result_area.delete(1.0, tk.END)
        for fname, lineno, text in preview[:1000]:
            self.result_area.insert(
                tk.END,
                f"{fname} (Line {lineno}): {text}\n"
            )
        if total > len(preview):
            self.result_area.insert(
                tk.END,
                f"\n--- Showing first {len(preview)} of {total} results ---\n"
            )
            self.result_area.insert(
                tk.END,
                "Please use 'Export to CSV' or 'Export to Excel' to see all matches.\n"
            )
        self.result_area.yview_moveto(0)

    def export_csv(self):
        if not self.temp_csv_path or not self.temp_csv_path.exists():
            messagebox.showwarning("No Data", "No results to export.")
            return
        dest = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )
        if dest:
            try:
                self.temp_csv_path.write_bytes(
                    Path(self.temp_csv_path).read_bytes()
                )
                messagebox.showinfo("Success", f"Exported results to {dest}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to write CSV:\n{e}")

    def export_excel(self):
        if not self.temp_csv_path or not self.temp_csv_path.exists():
            messagebox.showwarning("No Data", "No results to export.")
            return
        dest = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if dest:
            try:
                import openpyxl
                from openpyxl.styles import Font

                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Search Results"
                with open(self.temp_csv_path, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    for r, row in enumerate(reader, 1):
                        for c, v in enumerate(row, 1):
                            cell = ws.cell(row=r, column=c, value=v)
                            if r == 1:
                                cell.font = Font(bold=True)
                wb.save(dest)
                messagebox.showinfo("Success", f"Exported results to {dest}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export Excel:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TextSearchApp(root)
    root.mainloop()
