import os
import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import font as tkfont
import platform

EXE_PATH = os.environ.get("VSS_EXE_PATH", r"C:\VCP-Lite\Base\FELINK\Security\Lockdown_active.exe")
APP_VERSION = "Version 2.0 VCP Lite7"

def on_lockdown():
    lock_btn.config(state=tk.DISABLED)
    if not os.path.exists(EXE_PATH):
        messagebox.showerror("Error", f"File not found:\n{EXE_PATH}")
        lock_btn.config(state=tk.NORMAL)
        return

    try:
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if platform.system() == "Windows" else 0
        subprocess.Popen(
            [EXE_PATH],
            shell=False,
            creationflags=creation_flags
        )
        messagebox.showinfo("Launched", "Lockdown process started.")
        root.destroy()
    except Exception as e:
        lock_btn.config(state=tk.NORMAL)
        messagebox.showerror("Error", f"Cannot start:\n{EXE_PATH}\n\n{e}")

root = tk.Tk()
root.title("Vynamic Security Check")
root.resizable(False, False)

# Window size
w, h = 520, 280

# Center window
root.update_idletasks()
x = (root.winfo_screenwidth() - w) // 2
y = (root.winfo_screenheight() - h) // 2
root.geometry(f"{w}x{h}+{x}+{y}")

# Fonts
title_font = tkfont.Font(family="Arial", size=26, weight="bold")
msg_font   = tkfont.Font(family="Arial", size=16)
ver_font   = tkfont.Font(family="Arial", size=9)

# Main content frame
frame = tk.Frame(root)
frame.pack(expand=True)

tk.Label(frame, text="VSS Policy not Activate!", fg="red", font=title_font).pack(pady=(10, 10))
tk.Label(frame, text="Please Lockdown Policy", fg="blue", font=msg_font).pack(pady=(0, 25))
lock_btn = tk.Button(frame, text="LOCKDOWN", width=14, height=2, command=on_lockdown)
lock_btn.pack()

# ---- Version label (bottom-right) ----
version_label = tk.Label(
    root,
    text=APP_VERSION,
    font=ver_font,
    fg="gray"
)
version_label.place(relx=1.0, rely=1.0, anchor="se", x=-8, y=-5)

root.mainloop()
