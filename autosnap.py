import tkinter as tk
import time
from pathlib import Path
import pyautogui

def take_screenshot():
    # Hide the window
    root.withdraw()
    time.sleep(0.5)  # Short delay to allow the window to fully hide

    # Save path
    output_dir = Path("D:/MyShots")  # Change path as needed
    output_dir.mkdir(parents=True, exist_ok=True)

    # File name with timestamp
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    file_path = output_dir / f"screenshot_{timestamp}.png"

    # Capture full screen
    screenshot = pyautogui.screenshot()
    screenshot.save(file_path)

    # Re-show the window (optional)
    root.deiconify()

# GUI setup
root = tk.Tk()
root.title("Screenshot App")
root.geometry("300x150")
root.resizable(False, False)

label = tk.Label(root, text="Click to take full screen snapshot", font=("Arial", 12))
label.pack(pady=20)

btn = tk.Button(root, text="Take Screenshot", font=("Arial", 12), command=take_screenshot)
btn.pack()

root.mainloop()
