import tkinter as tk
from tkinter import filedialog, messagebox

def count_text():
    filepath = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if not filepath:
        return

    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            text = file.read()

        word_count = len(text.split())
        char_count = len(text)

        result_label.config(text=f"File: {filepath}\nWords: {word_count}\nCharacters: {char_count}")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")

# Create GUI
root = tk.Tk()
root.title("Text Counter from File")
root.geometry("500x200")

select_button = tk.Button(root, text="Select Text File", command=count_text)
select_button.pack(pady=20)

result_label = tk.Label(root, text="Select a file to count text.", wraplength=480, justify="left")
result_label.pack(padx=10)

root.mainloop()
