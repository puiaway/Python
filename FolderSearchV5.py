import os
import csv
import tempfile
import threading
import platform
import chardet
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import json

# --- Dependency Check ---
try:
    import openpyxl
    from openpyxl.styles import Font
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# --- Constants ---
HISTORY_FILE = os.path.expanduser("~/.text_search_app_history.json")
DEFAULT_EXTENSIONS = ".txt, .log, .csv, .json, .xml, .md, .py"

# --- Helper Functions ---
def safe_path(path):
    """Prepares a path for long path support on Windows."""
    if platform.system() == 'Windows':
        path = os.path.abspath(path)
        if not path.startswith('\\\\?\\'):
            path = '\\\\?\\' + path
    return path

def detect_encoding(file_path, sample_size=10000):
    """Detects file encoding using a sample of the file."""
    with open(file_path, 'rb') as f:
        raw_data = f.read(sample_size)
    # Fallback to utf-8 if detection fails
    return chardet.detect(raw_data)['encoding'] or 'utf-8'


class TextSearchApp:
    def __init__(self, master):
        self.master = master
        master.title("Text Search Pro")
        master.geometry("930x700")
        master.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.stop_flag = threading.Event()
        self.temp_csv_path = None

        self._create_widgets()
        self.keyword_history = self._load_keyword_history()
        self.keyword_entry['values'] = self.keyword_history

    def _create_widgets(self):
        """Creates and places all GUI widgets."""
        # --- Main Frame ---
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.master.rowconfigure(0, weight=1)
        self.master.columnconfigure(0, weight=1)

        # --- Input Widgets ---
        ttk.Label(main_frame, text="Folder:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.folder_entry = ttk.Entry(main_frame, width=80)
        self.folder_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=2)
        ttk.Button(main_frame, text="Browse...", command=self.browse_folder).grid(row=0, column=3, padx=5)

        ttk.Label(main_frame, text="Keyword:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.keyword_entry = ttk.Combobox(main_frame, width=78)
        self.keyword_entry.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=2)

        ttk.Label(main_frame, text="File Types:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.extensions_entry = ttk.Entry(main_frame, width=78)
        self.extensions_entry.grid(row=2, column=1, columnspan=2, sticky="ew", padx=5, pady=2)
        self.extensions_entry.insert(0, DEFAULT_EXTENSIONS)

        self.search_btn = ttk.Button(main_frame, text="Search", command=self.start_search_thread)
        self.search_btn.grid(row=1, column=3, padx=5)

        # --- Options Frame ---
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.grid(row=3, column=0, columnspan=4, sticky='ew', padx=5, pady=5)

        self.match_once_var = tk.BooleanVar()
        self.last_match_var = tk.BooleanVar()
        self.case_sensitive_var = tk.BooleanVar()
        self.show_all_var = tk.BooleanVar()
        self.include_nomatch_var = tk.BooleanVar(value=True)

        ttk.Checkbutton(options_frame, text="Match once per file", variable=self.match_once_var).grid(row=0, column=0, sticky='w', padx=5)
        ttk.Checkbutton(options_frame, text="Use last match", variable=self.last_match_var).grid(row=0, column=1, sticky='w', padx=5)
        ttk.Checkbutton(options_frame, text="Case sensitive", variable=self.case_sensitive_var).grid(row=0, column=2, sticky='w', padx=5)
        ttk.Checkbutton(options_frame, text="Include non-matching files", variable=self.include_nomatch_var).grid(row=0, column=3, sticky='w', padx=5)
        ttk.Checkbutton(options_frame, text="Show all in preview", variable=self.show_all_var).grid(row=0, column=4, sticky='w', padx=5)

        # --- Progress & Status ---
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=4, sticky="ew", padx=5, pady=5)
        self.status_label = ttk.Label(main_frame, text="Status: Ready")
        self.status_label.grid(row=5, column=0, columnspan=2, sticky="w", padx=5)

        # --- Action Buttons ---
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=5, column=2, columnspan=2, sticky="e", padx=5)
        self.cancel_btn = ttk.Button(action_frame, text="Cancel", command=self.cancel_search, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=5)
        self.export_csv_btn = ttk.Button(action_frame, text="Export CSV", command=self.export_csv, state=tk.DISABLED)
        self.export_csv_btn.pack(side=tk.LEFT, padx=5)
        self.export_excel_btn = ttk.Button(action_frame, text="Export Excel", command=self.export_excel, state=tk.DISABLED)
        self.export_excel_btn.pack(side=tk.LEFT, padx=5)
        
        if not OPENPYXL_AVAILABLE:
            self.export_excel_btn.config(text="Excel (install openpyxl)")

        # --- Result Area ---
        self.result_area = scrolledtext.ScrolledText(main_frame, width=110, height=25, wrap=tk.WORD)
        self.result_area.grid(row=6, column=0, columnspan=4, sticky="nsew", padx=5, pady=5)

        main_frame.rowconfigure(6, weight=1)
        main_frame.columnconfigure(1, weight=1)

    def _load_keyword_history(self):
        if not os.path.exists(HISTORY_FILE):
            return []
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return []

    def _save_keyword_history(self):
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.keyword_history, f, indent=4)
        except IOError:
            # Silently fail if history cannot be saved
            pass

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)

    def _set_ui_state(self, searching=False):
        """Helper to enable/disable UI elements."""
        state = tk.DISABLED if searching else tk.NORMAL
        self.search_btn.config(state=state)
        self.folder_entry.config(state=state)
        self.keyword_entry.config(state=state)
        self.extensions_entry.config(state=state)

        self.cancel_btn.config(state=tk.NORMAL if searching else tk.DISABLED)
        if not searching:
            self.export_csv_btn.config(state=tk.NORMAL)
            if OPENPYXL_AVAILABLE:
                self.export_excel_btn.config(state=tk.NORMAL)
        else:
            self.export_csv_btn.config(state=tk.DISABLED)
            self.export_excel_btn.config(state=tk.DISABLED)

    def start_search_thread(self):
        """Validates input and starts the search in a new thread."""
        folder = self.folder_entry.get()
        keyword = self.keyword_entry.get().strip()
        extensions = self.extensions_entry.get().strip()

        if not os.path.isdir(folder) or not keyword or not extensions:
            messagebox.showerror("Error", "Please provide a valid folder, keyword, and file types.")
            return

        # Update and save keyword history
        if keyword not in self.keyword_history:
            self.keyword_history.insert(0, keyword)
            self.keyword_history = self.keyword_history[:20]
            self.keyword_entry['values'] = self.keyword_history
            self._save_keyword_history()

        self.stop_flag.clear()
        self._set_ui_state(searching=True)
        self.result_area.delete(1.0, tk.END)
        self.status_label.config(text="Status: Preparing to search...")
        self.progress["value"] = 0

        # Start search in a separate thread
        search_thread = threading.Thread(
            target=self.search_files,
            args=(folder, keyword, extensions),
            daemon=True
        )
        search_thread.start()

    def cancel_search(self):
        self.status_label.config(text="Status: Cancelling...")
        self.stop_flag.set()

    def search_files(self, folder, keyword, extensions_str):
        """The core search logic that runs in a background thread."""
        # --- Prepare for search ---
        search_params = {
            "match_once": self.match_once_var.get(),
            "use_last": self.last_match_var.get(),
            "case_sens": self.case_sensitive_var.get(),
            "show_all": self.show_all_var.get(),
            "inc_nomatch": self.include_nomatch_var.get()
        }
        keyword_to_find = keyword if search_params["case_sens"] else keyword.lower()
        file_exts = tuple(ext.strip() for ext in extensions_str.split(',') if ext.strip())
        
        file_list = [os.path.join(root, file)
                     for root, _, files in os.walk(folder)
                     for file in files if file.lower().endswith(file_exts)]
        
        total_files = len(file_list)
        if total_files == 0:
            self.status_label.config(text="Status: No matching file types found.")
            self._set_ui_state(searching=False)
            return
            
        self.progress["maximum"] = total_files

        # --- Create temporary file for results ---
        try:
            temp_file = tempfile.NamedTemporaryFile(delete=False, mode='w', encoding='utf-8', newline='')
            self.temp_csv_path = temp_file.name
            writer = csv.writer(temp_file)
            writer.writerow(["Filename", "Line Number", "Line Content"])
        except IOError as e:
            messagebox.showerror("File Error", f"Could not create a temporary file: {e}")
            self._set_ui_state(searching=False)
            return

        # --- Main search loop ---
        preview_results, total_matches = self._process_files(file_list, writer, keyword_to_find, search_params)

        # --- Finalize ---
        temp_file.close()

        if self.stop_flag.is_set():
            self.status_label.config(text=f"Status: Search cancelled. Processed {self.progress['value']} files.")
        else:
            self.status_label.config(text=f"Status: Complete. Found {total_matches} matches in {total_files} files.")

        self.show_preview(preview_results, total_matches)
        self._set_ui_state(searching=False)

    def _process_files(self, file_list, writer, keyword, params):
        """Iterates through files and performs the search."""
        preview, match_count = [], 0
        total_files = len(file_list)

        for i, file_path in enumerate(file_list, 1):
            if self.stop_flag.is_set():
                break

            filename = os.path.basename(file_path)
            matches_in_file = []
            found_in_file = False
            
            try:
                encoding = detect_encoding(file_path)
                with open(safe_path(file_path), 'r', encoding=encoding, errors='ignore') as f:
                    last_match_data = None
                    for line_num, line in enumerate(f, 1):
                        line_to_check = line if params["case_sens"] else line.lower()
                        if keyword in line_to_check:
                            found_in_file = True
                            match_data = (filename, line_num, line.strip())
                            if params["match_once"]:
                                matches_in_file = [match_data]
                                break
                            elif params["use_last"]:
                                last_match_data = match_data
                            else:
                                matches_in_file.append(match_data)
                    if params["use_last"] and last_match_data:
                        matches_in_file = [last_match_data]
            except Exception as e:
                matches_in_file.append((filename, "ERROR", f"Could not read file: {e}"))

            # --- Write results for the current file ---
            if found_in_file or (params["inc_nomatch"] and not matches_in_file):
                if not matches_in_file and params["inc_nomatch"]:
                     matches_in_file.append((filename, "-", "[No match found]"))
                
                for match in matches_in_file:
                    writer.writerow(match)
                    if params["show_all"] or len(preview) < 1000:
                        preview.append(match)
                match_count += len(matches_in_file) if found_in_file else 0

            # --- Update progress ---
            if i % 10 == 0 or i == total_files:
                self.progress["value"] = i
                self.status_label.config(text=f"Status: Scanning {i}/{total_files} files...")
                self.master.update_idletasks()
        
        return preview, match_count
        
    def show_preview(self, preview, total_count):
        self.result_area.delete(1.0, tk.END)
        if not preview and not self.include_nomatch_var.get():
             self.result_area.insert(tk.END, "No matches found.")
             return

        for file, line_num, line in preview:
            self.result_area.insert(tk.END, f"{file} (Line {line_num}): {line}\n")
        
        if not self.show_all_var.get() and total_count > len(preview):
            self.result_area.insert(tk.END, f"\n--- Showing first {len(preview)} of {total_count} total matches ---\n")
            self.result_area.insert(tk.END, "Use 'Export to CSV' or 'Export to Excel' to view all results.")
        self.result_area.yview_moveto(0)

    def export_csv(self):
        if not self.temp_csv_path or not os.path.exists(self.temp_csv_path):
            messagebox.showwarning("No Data", "No results available to export.")
            return
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile="search_results.csv"
        )
        if save_path:
            try:
                # Simple file copy is fast and efficient
                import shutil
                shutil.copy(self.temp_csv_path, save_path)
                messagebox.showinfo("Success", f"Results exported to\n{save_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export CSV:\n{e}")

    def export_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Dependency Missing", "Please install the 'openpyxl' library to use this feature.\n\nCommand: pip install openpyxl")
            return
            
        if not self.temp_csv_path or not os.path.exists(self.temp_csv_path):
            messagebox.showwarning("No Data", "No results available to export.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="search_results.xlsx"
        )
        if save_path:
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Search Results"
                
                with open(self.temp_csv_path, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    for r_idx, row_data in enumerate(reader, 1):
                        for c_idx, cell_value in enumerate(row_data, 1):
                            cell = ws.cell(row=r_idx, column=c_idx, value=cell_value)
                            if r_idx == 1: # Bold header
                                cell.font = Font(bold=True)
                
                # Auto-fit column widths
                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    ws.column_dimensions[column].width = min(adjusted_width, 70)

                wb.save(save_path)
                messagebox.showinfo("Success", f"Results exported to\n{save_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export to Excel:\n{e}")

    def on_closing(self):
        """Handles window close event to clean up temporary files."""
        if self.temp_csv_path and os.path.exists(self.temp_csv_path):
            try:
                os.remove(self.temp_csv_path)
            except OSError:
                pass # Fail silently
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = TextSearchApp(root)
    root.mainloop()