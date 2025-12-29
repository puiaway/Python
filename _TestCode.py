import tkinter as tk
from tkinter import ttk
import time
from pathlib import Path
import mss
import keyboard
import threading

# ========== Screenshot Function ==========
def take_screenshot(monitor_index):
    try:
        output_dir = Path("D:/MyShots")  # Change save path here
        output_dir.mkdir(parents=True, exist_ok=True)

        timestamp = time.strftime("%Y%m%d-%H%M%S")
        file_path = output_dir / f"screenshot_monitor{monitor_index}_{timestamp}.png"

        with mss.mss() as sct:
            monitor = sct.monitors[monitor_index + 1]  # sct.monitors[1] = monitor 0
            img = sct.grab(monitor)
            mss.tools.to_png(img.rgb, img.size, output=str(file_path))
    except Exception as e:
        print(f"Error taking screenshot: {e}")

# ========== GUI + Hotkey Setup ==========
def setup_gui():
    def on_select(event=None):
        idx = int(monitor_var.get())
        info_label.config(text=monitor_list[idx])

    def gui_screenshot():
        idx = int(monitor_var.get())
        root.withdraw()
        time.sleep(0.5)
        take_screenshot(idx)
        root.deiconify()

    def listen_hotkey():
        while True:
            keyboard.wait('F9')
            idx = int(monitor_var.get())
            take_screenshot(idx)

    # Get all monitors
    with mss.mss() as sct:
        monitors = sct.monitors[1:]  # skip [0], which is all combined
        monitor_count = len(monitors)

    # Setup GUI
    global root
    root = tk.Tk()
    root.title("Single Monitor Screenshot (F9 Hotkey)")
    root.geometry("400x230")
    root.resizable(False, False)

    label = tk.Label(root, text="Select monitor to capture:", font=("Arial", 12))
    label.pack(pady=10)

    monitor_var = tk.StringVar(value="0")
    monitor_indices = [str(i) for i in range(monitor_count)]
    monitor_list = [f"{i}: {m['width']}x{m['height']} @ ({m['left']},{m['top']})" for i, m in enumerate(monitors)]

    combo = ttk.Combobox(root, values=monitor_indices, textvariable=monitor_var, font=("Arial", 10), state="readonly")
    combo.pack(pady=5)
    combo.current(0)

    info_label = tk.Label(root, text=monitor_list[0], font=("Arial", 10))
    info_label.pack(pady=5)
    combo.bind("<<ComboboxSelected>>", on_select)

    btn = tk.Button(root, text="Take Screenshot", font=("Arial", 12), command=gui_screenshot)
    btn.pack(pady=15)

    hotkey_label = tk.Label(root, text="Or press F9 to take screenshot (D:\MyShot)", font=("Arial", 10), fg="gray")
    hotkey_label.pack(pady=5)

    # Start hotkey listener thread
    threading.Thread(target=listen_hotkey, daemon=True).start()

    root.mainloop()

setup_gui()
