import tkinter as tk
from tkinter import filedialog, messagebox
import csv

def browse_file():
    file_path.set(filedialog.askopenfilename(filetypes=[("Text files", "*.txt *.log")]))


def search():
    filepath = file_path.get()
    word1 = entry_word1.get()
    word2 = entry_word2.get()

    if not filepath or not word1 or not word2:
        messagebox.showerror("Error", "Please select file and input both words.")
        return

    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as file:
            lines = file.readlines()

        results.clear()
        text_output.delete("1.0", tk.END)

        for i, line in enumerate(lines):
            if word1 in line and word2 in line:
                start = max(0, i - 5)
                context = lines[start:i+1]
                result_block = f"\n--- Found at line {i+1} ---\n" + ''.join(context)
                text_output.insert(tk.END, result_block)
                results.append((i+1, context))

    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")


def export_csv():
    if not results:
        messagebox.showinfo("No Data", "No results to export.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV files", "*.csv")])
    if not save_path:
        return

    try:
        with open(save_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["Match Line Number", "Context (Previous 5 lines + match)"])
            for lineno, context in results:
                context_text = ''.join(context).replace('\n', '\\n')
                writer.writerow([lineno, context_text])
        messagebox.showinfo("Success", f"Results exported to {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export CSV:\n{str(e)}")


# GUI Layout
root = tk.Tk()
root.title("Text Search (2 Keywords + Backward 5 Lines)")
root.geometry("700x500")

file_path = tk.StringVar()
results = []

tk.Label(root, text="Select Log File:").pack(anchor='w', padx=10, pady=(10,0))
tk.Entry(root, textvariable=file_path, width=60).pack(side='left', padx=10)
tk.Button(root, text="Browse", command=browse_file).pack(side='left')

tk.Label(root, text="Keyword 1:").pack(anchor='w', padx=10, pady=(10,0))
entry_word1 = tk.Entry(root, width=30)
entry_word1.pack(anchor='w', padx=10)

tk.Label(root, text="Keyword 2:").pack(anchor='w', padx=10, pady=(10,0))
entry_word2 = tk.Entry(root, width=30)
entry_word2.pack(anchor='w', padx=10)

tk.Button(root, text="Search", command=search).pack(pady=10)
tk.Button(root, text="Export CSV", command=export_csv).pack()

text_output = tk.Text(root, wrap='none', height=20)
text_output.pack(fill='both', expand=True, padx=10, pady=10)

root.mainloop()
