import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import shutil
import tempfile
import re
from datetime import datetime
import unicodedata  # for Thai/locale digits
import json

# --- Optional 7z support ---
try:
    import py7zr  # pip install py7zr
    PY7ZR_AVAILABLE = True
except Exception:
    PY7ZR_AVAILABLE = False

INVALID_WIN_CHARS = r'[<>:"/\\|?*\x00-\x1F]'
DATE_REGEX = re.compile(r"\d{4}-\d{2}-\d{2}")  # YYYY-MM-DD
PREFS_FILE = os.path.join(os.path.expanduser("~"), ".zip_rename_extractor_prefs.json")


def sanitize_filename(name: str) -> str:
    name = re.sub(INVALID_WIN_CHARS, "_", name).strip().rstrip(". ")
    return name or "unnamed"


def unique_path(base_path: str) -> str:
    if not os.path.exists(base_path):
        return base_path
    root, ext = os.path.splitext(base_path)
    i = 1
    while True:
        candidate = f"{root} ({i}){ext}"
        if not os.path.exists(candidate):
            return candidate
        i += 1


def extract_date_from_name(name: str):
    m = DATE_REGEX.search(name)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(), "%Y-%m-%d").date()
    except Exception:
        return None


class ZipRenameExtractor:
    def __init__(self, master):
        self.master = master
        master.title("Extract Archives (Prefix) — Recent or Date Range Mode")
        master.geometry("880x460")
        master.protocol("WM_DELETE_WINDOW", self.on_close)

        # Row 0: Archive folder
        tk.Label(master, text="Archive Folder (.zip / .7z):").grid(row=0, column=0, padx=10, pady=8, sticky='w')
        self.zip_folder_entry = tk.Entry(master, width=74)
        self.zip_folder_entry.grid(row=0, column=1)
        tk.Button(master, text="Browse", command=self.browse_zip_folder).grid(row=0, column=2, padx=6)

        # Row 1: Output folder
        tk.Label(master, text="Output Folder:").grid(row=1, column=0, padx=10, pady=8, sticky='w')
        self.output_folder_entry = tk.Entry(master, width=74)
        self.output_folder_entry.grid(row=1, column=1)
        tk.Button(master, text="Browse", command=self.browse_output_folder).grid(row=1, column=2, padx=6)

        # Row 2: Mode selector
        mode_frame = tk.Frame(master)
        mode_frame.grid(row=2, column=0, columnspan=3, sticky="w", padx=10, pady=4)
        tk.Label(mode_frame, text="Mode:").pack(side="left")
        self.mode_var = tk.StringVar(value="recent")  # "recent" or "date"
        self.rb_recent = tk.Radiobutton(mode_frame, text="Use Max Recent Files", variable=self.mode_var,
                                        value="recent", command=self._on_mode_change)
        self.rb_recent.pack(side="left", padx=8)
        self.rb_date = tk.Radiobutton(mode_frame, text="Use Start/End Date", variable=self.mode_var,
                                      value="date", command=self._on_mode_change)
        self.rb_date.pack(side="left", padx=8)

        # Row 3: Recent limit
        tk.Label(master, text="Max recent files per archive (0 or blank = all):").grid(row=3, column=0, padx=10, pady=4, sticky='w')
        self.limit_entry = tk.Entry(master, width=12)
        self.limit_entry.grid(row=3, column=1, sticky='w', padx=6)
        self.limit_entry.insert(0, "60")
        self.limit_entry.bind("<FocusOut>", lambda e: self.save_prefs())

        # Row 4/5: Start/End date
        tk.Label(master, text="Start Date (YYYY-MM-DD):").grid(row=4, column=0, padx=10, pady=4, sticky='w')
        self.start_entry = tk.Entry(master, width=18)
        self.start_entry.grid(row=4, column=1, sticky='w', padx=6)
        self.start_entry.bind("<FocusOut>", lambda e: self.save_prefs())

        tk.Label(master, text="End Date (YYYY-MM-DD):").grid(row=5, column=0, padx=10, pady=4, sticky='w')
        self.end_entry = tk.Entry(master, width=18)
        self.end_entry.grid(row=5, column=1, sticky='w', padx=6)
        self.end_entry.bind("<FocusOut>", lambda e: self.save_prefs())

        # Row 6: Progress
        self.progress = ttk.Progressbar(master, orient="horizontal", length=760, mode="determinate")
        self.progress.grid(row=6, column=0, columnspan=3, padx=10, pady=12)

        # Row 7: Buttons
        btns = tk.Frame(master)
        btns.grid(row=7, column=0, columnspan=3)
        self.extract_btn = tk.Button(btns, text="Extract & Rename", command=self.start_extract_thread)
        self.extract_btn.pack(side="left", padx=6)
        self.cancel_btn = tk.Button(btns, text="Cancel", command=self.request_cancel, state="disabled")
        self.cancel_btn.pack(side="left", padx=6)

        # Row 8: Status
        self.status_label = tk.Label(
            master,
            text="Status: Ready" + ("" if PY7ZR_AVAILABLE else "  (Tip: pip install py7zr for .7z)")
        )
        self.status_label.grid(row=8, column=0, columnspan=3, padx=10, sticky="w")

        self._cancel = False
        self._worker = None

        # Load and apply last selections
        self.load_prefs()
        self._toggle_inputs()

    # --- preferences ---
    def current_prefs(self):
        return {
            "zip_folder": self.zip_folder_entry.get().strip(),
            "output_folder": self.output_folder_entry.get().strip(),
            "mode": self.mode_var.get(),
            "limit": self.limit_entry.get().strip(),
            "start_date": self.start_entry.get().strip(),
            "end_date": self.end_entry.get().strip(),
            "geometry": self.master.winfo_geometry(),
        }

    def save_prefs(self):
        try:
            with open(PREFS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.current_prefs(), f, ensure_ascii=False, indent=2)
        except Exception:
            # Don't crash on prefs save errors
            pass

    def load_prefs(self):
        try:
            if os.path.isfile(PREFS_FILE):
                with open(PREFS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # Populate fields if present
                if isinstance(data, dict):
                    if data.get("zip_folder"):
                        self.zip_folder_entry.delete(0, tk.END)
                        self.zip_folder_entry.insert(0, data["zip_folder"])
                    if data.get("output_folder"):
                        self.output_folder_entry.delete(0, tk.END)
                        self.output_folder_entry.insert(0, data["output_folder"])
                    if data.get("mode") in ("recent", "date"):
                        self.mode_var.set(data["mode"])
                    if "limit" in data and str(data["limit"]).strip():
                        self.limit_entry.delete(0, tk.END)
                        self.limit_entry.insert(0, str(data["limit"]))
                    if data.get("start_date"):
                        self.start_entry.delete(0, tk.END)
                        self.start_entry.insert(0, data["start_date"])
                    if data.get("end_date"):
                        self.end_entry.delete(0, tk.END)
                        self.end_entry.insert(0, data["end_date"])
                    if data.get("geometry"):
                        # Geometry validity can vary; wrap to avoid errors
                        try:
                            self.master.geometry(data["geometry"])
                        except Exception:
                            pass
        except Exception:
            # Ignore malformed JSON or read errors
            pass

    def on_close(self):
        self.save_prefs()
        self.master.destroy()

    # --- UI helpers ---
    def _on_mode_change(self):
        self._toggle_inputs()
        self.save_prefs()

    def _toggle_inputs(self):
        use_recent = self.mode_var.get() == "recent"
        self.limit_entry.config(state=("normal" if use_recent else "disabled"))
        self.start_entry.config(state=("disabled" if use_recent else "normal"))
        self.end_entry.config(state=("disabled" if use_recent else "normal"))

    def ui(self, fn, *args, **kwargs):
        self.master.after(0, fn, *args, **kwargs)

    def set_status(self, text):
        self.status_label.config(text=text)

    def set_buttons_running(self, running: bool):
        self.extract_btn.config(state="disabled" if running else "normal")
        self.cancel_btn.config(state="normal" if running else "disabled")

    # --- browse ---
    def browse_zip_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.zip_folder_entry.delete(0, tk.END)
            self.zip_folder_entry.insert(0, folder)
            self.save_prefs()

    def browse_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, folder)
            self.save_prefs()

    # --- run/cancel ---
    def start_extract_thread(self):
        if self._worker and self._worker.is_alive():
            return
        self._cancel = False
        self.set_buttons_running(True)
        self.ui(self.set_status, "Status: Working…")
        self.ui(lambda: self.progress.config(value=0))
        self.save_prefs()  # persist current selections before running
        self._worker = threading.Thread(target=self.extract_and_rename, daemon=True)
        self._worker.start()

    def request_cancel(self):
        self._cancel = True
        self.ui(self.set_status, "Status: Cancelling…")

    # --- parsing helpers ---
    def parse_limit(self):
        txt = (self.limit_entry.get() or "").strip()
        if not txt:
            return None  # no limit
        try:
            # normalize locale digits (Thai "๖๐" etc.) to ASCII
            norm = "".join(str(unicodedata.digit(c)) if c.isdigit() else c for c in txt)
            n = int(norm)
            return None if n <= 0 else n
        except Exception:
            return "invalid"

    def parse_date_range(self):
        s = self.start_entry.get().strip()
        e = self.end_entry.get().strip()
        start_date = end_date = None
        try:
            if s:
                start_date = datetime.strptime(s, "%Y-%m-%d").date()
            if e:
                end_date = datetime.strptime(e, "%Y-%m-%d").date()
        except ValueError:
            return "invalid"
        if not s and not e:
            return "empty"  # in date mode, require at least one bound
        return (start_date, end_date)

    def is_in_range(self, fname, start_date, end_date):
        """True if filename contains a date (YYYY-MM-DD) within [start, end]."""
        d = extract_date_from_name(fname)
        if not d:
            return False
        if start_date and d < start_date:
            return False
        if end_date and d > end_date:
            return False
        return True

    # --- core ---
    def extract_and_rename(self):
        zip_folder = self.zip_folder_entry.get()
        output_folder = self.output_folder_entry.get()
        mode = self.mode_var.get()

        # Validate inputs per mode
        if mode == "recent":
            limit = self.parse_limit()
            if limit == "invalid":
                return self.ui(messagebox.showerror, "Error", "Invalid max recent files number.")
            date_range = None
        else:  # date mode
            date_range = self.parse_date_range()
            if date_range in ("invalid", "empty"):
                msg = "Invalid date format (use YYYY-MM-DD)." if date_range == "invalid" else "Please set Start and/or End date."
                return self.ui(messagebox.showerror, "Error", msg)
            limit = None  # ignored in date mode

        if not os.path.isdir(zip_folder) or not os.path.isdir(output_folder):
            return self.ui(messagebox.showerror, "Error", "Please select valid folders.")

        archives = sorted(f for f in os.listdir(zip_folder) if f.lower().endswith((".zip", ".7z")))
        if not archives:
            return self.ui(messagebox.showinfo, "Info", "No .zip or .7z files found.")

        if any(f.lower().endswith(".7z") for f in archives) and not PY7ZR_AVAILABLE:
            self.ui(
                messagebox.showwarning,
                "7z Not Available",
                "'.7z' files detected but py7zr is not installed.\nInstall with: pip install py7zr"
            )
            archives = [f for f in archives if f.lower().endswith(".zip")]

        self.ui(lambda m=len(archives): self.progress.config(maximum=m))
        self.ui(lambda: self.progress.config(value=0))

        errors = []

        for i, arc_name in enumerate(archives, 1):
            if self._cancel:
                break

            arc_path = os.path.join(zip_folder, arc_name)
            arc_prefix = sanitize_filename(os.path.splitext(arc_name)[0])

            try:
                if arc_name.lower().endswith(".zip"):
                    if mode == "recent":
                        self._zip_extract_recent(arc_path, output_folder, arc_prefix, limit)
                    else:
                        start_date, end_date = date_range
                        self._zip_extract_by_date(arc_path, output_folder, arc_prefix, start_date, end_date)
                else:
                    if mode == "recent":
                        self._7z_extract_recent(arc_path, output_folder, arc_prefix, limit)
                    else:
                        start_date, end_date = date_range
                        self._7z_extract_by_date(arc_path, output_folder, arc_prefix, start_date, end_date)

                self.ui(self.set_status, f"Extracted: {arc_name}")

            except Exception as e:
                errors.append(f"{arc_name}: {e}")
                self.ui(self.set_status, f"Error: {arc_name} — {e}")

            self.ui(lambda v=i: self.progress.config(value=v))

        def finish():
            self.set_buttons_running(False)
            if self._cancel:
                self.set_status("Status: Cancelled")
                messagebox.showinfo("Cancelled", "Operation was cancelled.")
            elif errors:
                self.set_status("Status: Done with errors")
                messagebox.showwarning("Completed with errors", "\n".join(errors[:10]) + ("\n..." if len(errors) > 10 else ""))
            else:
                self.set_status("Status: Done")
                messagebox.showinfo("Done", "All archives processed.")

        self.ui(finish)

    # --- ZIP strategies ---
    def _zip_extract_recent(self, arc_path, output_folder, arc_prefix, limit):
        with zipfile.ZipFile(arc_path, 'r', allowZip64=True) as zf:
            members = [zi for zi in zf.infolist() if not zi.filename.endswith(('/', '\\'))]
            # newest first by ZIP internal timestamp
            def _mtime(zi):
                try:
                    y, mo, d, hh, mm, ss = zi.date_time
                    return datetime(y, mo, d, hh, mm, ss)
                except Exception:
                    return datetime.min
            members.sort(key=_mtime, reverse=True)
            selected = members[:limit] if isinstance(limit, int) else members
            self.ui(self.set_status, f"{os.path.basename(arc_path)}: selecting {len(selected)}/{len(members)} newest files")

            for zi in selected:
                if self._cancel:
                    break
                base = os.path.basename(zi.filename)
                if not base:
                    continue
                self._zip_copy_member(zf, zi, output_folder, arc_prefix, base)

    def _zip_extract_by_date(self, arc_path, output_folder, arc_prefix, start_date, end_date):
        with zipfile.ZipFile(arc_path, 'r', allowZip64=True) as zf:
            members = [zi for zi in zf.infolist() if not zi.filename.endswith(('/', '\\'))]
            # filter by date in filename only
            members = [zi for zi in members if self.is_in_range(zi.filename, start_date, end_date)]
            self.ui(self.set_status, f"{os.path.basename(arc_path)}: selecting {len(members)} by date")
            for zi in members:
                if self._cancel:
                    break
                base = os.path.basename(zi.filename)
                if not base:
                    continue
                self._zip_copy_member(zf, zi, output_folder, arc_prefix, base)

    def _zip_copy_member(self, zf, zi, output_folder, arc_prefix, base):
        new_name = sanitize_filename(f"{arc_prefix}_{sanitize_filename(base)}")
        dest_path = unique_path(os.path.join(output_folder, new_name))
        fd, tmp_path = tempfile.mkstemp(prefix="arcx_", dir=output_folder)
        try:
            with os.fdopen(fd, 'wb') as tmp, zf.open(zi, 'r') as src:
                shutil.copyfileobj(src, tmp, length=1024 * 1024)
            os.replace(tmp_path, dest_path)
        except Exception:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            finally:
                raise

    # --- 7Z strategies ---
    def _7z_extract_recent(self, arc_path, output_folder, arc_prefix, limit):
        if not PY7ZR_AVAILABLE:
            raise RuntimeError("py7zr not installed")
        with py7zr.SevenZipFile(arc_path, mode='r') as zf, tempfile.TemporaryDirectory() as tmpdir:
            zf.extractall(path=tmpdir)
            collected = []
            for root, _, files in os.walk(tmpdir):
                for fname in files:
                    p = os.path.join(root, fname)
                    try:
                        mt = os.path.getmtime(p)
                    except Exception:
                        mt = 0
                    collected.append((p, fname, mt))

            collected.sort(key=lambda t: t[2], reverse=True)
            selected = collected[:limit] if isinstance(limit, int) else collected
            self.ui(self.set_status, f"{os.path.basename(arc_path)}: selecting {len(selected)}/{len(collected)} newest files")

            for src_path, fname, _ in selected:
                if self._cancel:
                    break
                self._file_copy_to_output(src_path, output_folder, arc_prefix, fname)

    def _7z_extract_by_date(self, arc_path, output_folder, arc_prefix, start_date, end_date):
        if not PY7ZR_AVAILABLE:
            raise RuntimeError("py7zr not installed")
        with py7zr.SevenZipFile(arc_path, mode='r') as zf, tempfile.TemporaryDirectory() as tmpdir:
            zf.extractall(path=tmpdir)
            collected = []
            for root, _, files in os.walk(tmpdir):
                for fname in files:
                    if not self.is_in_range(fname, start_date, end_date):
                        continue
                    p = os.path.join(root, fname)
                    collected.append((p, fname))

            self.ui(self.set_status, f"{os.path.basename(arc_path)}: selecting {len(collected)} by date")
            for src_path, fname in collected:
                if self._cancel:
                    break
                self._file_copy_to_output(src_path, output_folder, arc_prefix, fname)

    def _file_copy_to_output(self, src_path, output_folder, arc_prefix, fname):
        new_name = sanitize_filename(f"{arc_prefix}_{sanitize_filename(fname)}")
        dest_path = unique_path(os.path.join(output_folder, new_name))
        fd, tmp_path = tempfile.mkstemp(prefix="arcx_", dir=output_folder)
        try:
            with os.fdopen(fd, 'wb') as tmp, open(src_path, 'rb') as src:
                shutil.copyfileobj(src, tmp, length=1024 * 1024)
            os.replace(tmp_path, dest_path)
        except Exception:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            finally:
                raise


if __name__ == "__main__":
    root = tk.Tk()
    app = ZipRenameExtractor(root)
    root.mainloop()
