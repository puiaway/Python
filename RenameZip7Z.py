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

# --- Optional 7z support ---
try:
    import py7zr  # pip install py7zr
    PY7ZR_AVAILABLE = True
except Exception:
    PY7ZR_AVAILABLE = False

INVALID_WIN_CHARS = r'[<>:"/\\|?*\x00-\x1F]'

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

class ZipRenameExtractor:
    def __init__(self, master):
        self.master = master
        master.title("Extract Archives and Rename Files with Archive Prefix")
        master.geometry("760x340")

        # Row 0: Archive folder
        tk.Label(master, text="Archive Folder (.zip / .7z):").grid(row=0, column=0, padx=10, pady=8, sticky='w')
        self.zip_folder_entry = tk.Entry(master, width=68)
        self.zip_folder_entry.grid(row=0, column=1)
        tk.Button(master, text="Browse", command=self.browse_zip_folder).grid(row=0, column=2, padx=6)

        # Row 1: Output folder
        tk.Label(master, text="Output Folder:").grid(row=1, column=0, padx=10, pady=8, sticky='w')
        self.output_folder_entry = tk.Entry(master, width=68)
        self.output_folder_entry.grid(row=1, column=1)
        tk.Button(master, text="Browse", command=self.browse_output_folder).grid(row=1, column=2, padx=6)

        # Row 2: Limit input
        limit_frame = tk.Frame(master)
        limit_frame.grid(row=2, column=0, columnspan=3, sticky="w", padx=10, pady=4)
        tk.Label(limit_frame, text="Max recent files per archive (0 or blank = all):").pack(side="left")
        self.limit_entry = tk.Entry(limit_frame, width=10)
        self.limit_entry.pack(side="left", padx=8)
        self.limit_entry.insert(0, "60")  # default

        # Row 3: Progress
        self.progress = ttk.Progressbar(master, orient="horizontal", length=680, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=3, padx=10, pady=12)

        # Row 4: Buttons
        btns = tk.Frame(master)
        btns.grid(row=4, column=0, columnspan=3)
        self.extract_btn = tk.Button(btns, text="Extract & Rename", command=self.start_extract_thread)
        self.extract_btn.pack(side="left", padx=6)
        self.cancel_btn = tk.Button(btns, text="Cancel", command=self.request_cancel, state="disabled")
        self.cancel_btn.pack(side="left", padx=6)

        # Row 5: Status
        self.status_label = tk.Label(
            master,
            text="Status: Ready" + ("" if PY7ZR_AVAILABLE else "  (Tip: pip install py7zr for .7z)")
        )
        self.status_label.grid(row=5, column=0, columnspan=3, padx=10, sticky="w")

        self._cancel = False
        self._worker = None

    # --- thread-safe UI calls ---
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

    def browse_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, folder)

    # --- run/cancel ---
    def start_extract_thread(self):
        if self._worker and self._worker.is_alive():
            return
        self._cancel = False
        self.set_buttons_running(True)
        self.ui(self.set_status, "Status: Working…")
        self.ui(lambda: self.progress.config(value=0))
        self._worker = threading.Thread(target=self.extract_and_rename, daemon=True)
        self._worker.start()

    def request_cancel(self):
        self._cancel = True
        self.ui(self.set_status, "Status: Cancelling…")

    # --- helpers ---
    def parse_limit(self):
        txt = (self.limit_entry.get() or "").strip()
        if not txt:
            return None  # blank => no limit
        try:
            # normalize locale digits (e.g., Thai "๖๐") to ASCII
            norm = "".join(str(unicodedata.digit(c)) if c.isdigit() else c for c in txt)
            n = int(norm)
            return None if n <= 0 else n
        except Exception:
            return "invalid"

    # --- core ---
    def extract_and_rename(self):
        zip_folder = self.zip_folder_entry.get()
        output_folder = self.output_folder_entry.get()

        # Validate limit
        limit = self.parse_limit()
        if limit == "invalid":
            return self.ui(messagebox.showerror, "Error", "Max recent files must be an integer (e.g., 60 or 0 for all).")

        if not os.path.isdir(zip_folder) or not os.path.isdir(output_folder):
            return self.ui(messagebox.showerror, "Error", "Please select valid folders.")

        all_names = sorted(os.listdir(zip_folder))
        archives = [f for f in all_names if f.lower().endswith((".zip", ".7z"))]

        if not archives:
            return self.ui(messagebox.showinfo, "Info", "No .zip or .7z files found.")

        if any(f.lower().endswith(".7z") for f in archives) and not PY7ZR_AVAILABLE:
            self.ui(messagebox.showwarning, "7z Not Available",
                    "'.7z' files detected but the 'py7zr' package is not installed.\n\n"
                    "Install with: pip install py7zr\n\nContinuing with .zip files only.")
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
                    zf = None
                    try:
                        zf = zipfile.ZipFile(arc_path, 'r', allowZip64=True)
                        # Filter to files only (dirs can be ambiguous; check by trailing slash)
                        members = [zi for zi in zf.infolist() if not zi.filename.endswith(('/', '\\'))]

                        # Sort by ZIP internal datetime (newest first)
                        def _zip_mtime(zi):
                            try:
                                y, mo, d, hh, mm, ss = zi.date_time
                                return datetime(y, mo, d, hh, mm, ss)
                            except Exception:
                                return datetime.min

                        members.sort(key=_zip_mtime, reverse=True)
                        selected = members[:limit] if isinstance(limit, int) else members

                        # Optional feedback
                        self.ui(self.set_status, f"{arc_name}: selecting {len(selected)}/{len(members)} newest files")

                        for zi in selected:
                            if self._cancel:
                                break
                            base = os.path.basename(zi.filename)
                            if not base:
                                continue

                            original_name = sanitize_filename(base)
                            new_name = sanitize_filename(f"{arc_prefix}_{original_name}")
                            dest_path = unique_path(os.path.join(output_folder, new_name))

                            try:
                                with zf.open(zi, 'r') as source:
                                    fd, tmp_path = tempfile.mkstemp(prefix="arcx_", dir=output_folder)
                                    with os.fdopen(fd, 'wb') as tmp:
                                        shutil.copyfileobj(source, tmp, length=1024 * 1024)
                                os.replace(tmp_path, dest_path)
                            except Exception as inner_e:
                                try:
                                    if 'tmp_path' in locals() and os.path.exists(tmp_path):
                                        os.remove(tmp_path)
                                finally:
                                    raise inner_e
                    finally:
                        if zf:
                            zf.close()

                else:
                    # --- .7z handling: extract to temp, select newest N by mtime, prefix, atomic move ---
                    if not PY7ZR_AVAILABLE:
                        raise RuntimeError("py7zr not available to open .7z")

                    zf = None
                    try:
                        zf = py7zr.SevenZipFile(arc_path, mode='r')  # type: ignore
                        with tempfile.TemporaryDirectory() as tmpdir:
                            zf.extractall(path=tmpdir)

                            collected = []
                            for root, _, files in os.walk(tmpdir):
                                for fname in files:
                                    if self._cancel:
                                        break
                                    p = os.path.join(root, fname)
                                    try:
                                        mt = os.path.getmtime(p)
                                    except Exception:
                                        mt = 0
                                    collected.append((p, fname, mt))

                            # Newest first by filesystem mtime
                            collected.sort(key=lambda t: t[2], reverse=True)
                            selected = collected[:limit] if isinstance(limit, int) else collected

                            self.ui(self.set_status, f"{arc_name}: selecting {len(selected)}/{len(collected)} newest files")

                            for src_path, fname, _mt in selected:
                                if self._cancel:
                                    break
                                original_name = sanitize_filename(fname)
                                new_name = sanitize_filename(f"{arc_prefix}_{original_name}")
                                dest_path = unique_path(os.path.join(output_folder, new_name))

                                # atomic write into output folder
                                try:
                                    fd, tmp_path = tempfile.mkstemp(prefix="arcx_", dir=output_folder)
                                    with os.fdopen(fd, 'wb') as tmp, open(src_path, 'rb') as src:
                                        shutil.copyfileobj(src, tmp, length=1024 * 1024)
                                    os.replace(tmp_path, dest_path)
                                except Exception as inner_e:
                                    try:
                                        if 'tmp_path' in locals() and os.path.exists(tmp_path):
                                            os.remove(tmp_path)
                                    finally:
                                        raise inner_e
                    finally:
                        if zf:
                            zf.close()

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
                messagebox.showwarning(
                    "Completed with errors",
                    "Some archives failed:\n\n" + "\n".join(errors[:10]) + ("\n..." if len(errors) > 10 else "")
                )
            else:
                self.set_status("Status: Done")
                messagebox.showinfo("Done", "All archives extracted and renamed.")

        self.ui(finish)

if __name__ == "__main__":
    root = tk.Tk()
    app = ZipRenameExtractor(root)
    root.mainloop()
