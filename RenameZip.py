import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import shutil

class ZipRenameExtractor:
    def __init__(self, master):
        self.master = master
        master.title("Extract ZIPs and Rename Files with ZIP Prefix")
        master.geometry("600x240")

        tk.Label(master, text="ZIP Folder:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.zip_folder_entry = tk.Entry(master, width=50)
        self.zip_folder_entry.grid(row=0, column=1)
        tk.Button(master, text="Browse", command=self.browse_zip_folder).grid(row=0, column=2)

        tk.Label(master, text="Output Folder:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        self.output_folder_entry = tk.Entry(master, width=50)
        self.output_folder_entry.grid(row=1, column=1)
        tk.Button(master, text="Browse", command=self.browse_output_folder).grid(row=1, column=2)

        self.progress = ttk.Progressbar(master, orient="horizontal", length=460, mode="determinate")
        self.progress.grid(row=2, column=0, columnspan=3, padx=10, pady=15)

        self.extract_btn = tk.Button(master, text="Extract & Rename", command=self.start_extract_thread)
        self.extract_btn.grid(row=3, column=1, pady=5)

        self.status_label = tk.Label(master, text="Status: Ready")
        self.status_label.grid(row=4, column=0, columnspan=3, padx=10, sticky="w")

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

    def start_extract_thread(self):
        thread = threading.Thread(target=self.extract_and_rename)
        thread.start()

    def extract_and_rename(self):
        zip_folder = self.zip_folder_entry.get()
        output_folder = self.output_folder_entry.get()

        if not os.path.isdir(zip_folder) or not os.path.isdir(output_folder):
            messagebox.showerror("Error", "Please select valid folders.")
            return

        zip_files = [f for f in os.listdir(zip_folder) if f.lower().endswith('.zip')]
        if not zip_files:
            messagebox.showinfo("Info", "No ZIP files found.")
            return

        self.progress["maximum"] = len(zip_files)
        self.progress["value"] = 0

        for i, zip_name in enumerate(zip_files, 1):
            zip_path = os.path.join(zip_folder, zip_name)
            zip_prefix = os.path.splitext(zip_name)[0]

            try:
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    for file_info in zip_ref.infolist():
                        if file_info.is_dir():
                            continue
                        original_name = os.path.basename(file_info.filename)
                        if not original_name:
                            continue

                        new_name = f"{zip_prefix}_{original_name}"
                        dest_path = os.path.join(output_folder, new_name)

                        # Create subfolder structure if needed
                        with zip_ref.open(file_info, 'r') as source, open(dest_path, 'wb') as target:
                            shutil.copyfileobj(source, target, length=1024 * 1024)  # Copy in 1MB chunks

                self.status_label.config(text=f"Extracted: {zip_name}")
            except Exception as e:
                self.status_label.config(text=f"Error: {zip_name} - {e}")

            self.progress["value"] = i
            self.master.update_idletasks()

        self.status_label.config(text="Status: Done")
        messagebox.showinfo("Done", "All ZIP files extracted and renamed.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ZipRenameExtractor(root)
    root.mainloop()
    