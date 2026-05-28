import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from openpyxl import Workbook
import math
import random

def update_item_count():
    """Update counter jumlah item secara real-time"""
    try:
        raw_text = text_input.get("1.0", tk.END).strip()
        items = [line.strip() for line in raw_text.split('\n') if line.strip()]
        count_label.config(text=f"Total Item: {len(items)}")
        
        try:
            num = int(entry_columns.get())
            if num > 0:
                if items_per_column_mode.get():
                    items_per_column = num
                    num_columns = math.ceil(len(items) / items_per_column) if len(items) > 0 else 0
                    column_info_label.config(text=f"{num_columns} kolom" if num_columns > 0 else "")
                else:
                    num_columns = num
                    items_per_column = math.ceil(len(items) / num_columns) if len(items) > 0 else 0
                    column_info_label.config(text=f"{items_per_column}/kolom" if items_per_column > 0 else "")
            else:
                column_info_label.config(text="")
        except ValueError:
            column_info_label.config(text="")
    except Exception as e:
        print(f"Error in update_item_count: {str(e)}")

def process_list():
    try:
        raw_text = text_input.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showerror("Error", "Silakan paste list Anda terlebih dahulu!")
            return
        
        items = [line.strip() for line in raw_text.split('\n') if line.strip()]
        if not items:
            messagebox.showerror("Error", "Tidak ada item yang valid untuk diproses!")
            return
        
        if shuffle_var.get():
            try:
                random.shuffle(items)
                status_label.config(text="List telah diacak", foreground="#4CAF50")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal mengacak list:\n{str(e)}")
                return
        
        try:
            num = int(entry_columns.get())
            if num < 1:
                messagebox.showerror("Error", "Nilai harus lebih dari 0!")
                return
        except ValueError:
            messagebox.showerror("Error", "Harap masukkan angka yang valid!")
            return
        
        if items_per_column_mode.get():
            items_per_column = num
            num_columns = math.ceil(len(items) / items_per_column)
        else:
            num_columns = num
            items_per_column = math.ceil(len(items) / num_columns)
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Data Terbagi"
        
        for col in range(num_columns):
            start_idx = col * items_per_column
            end_idx = start_idx + items_per_column
            for row, item in enumerate(items[start_idx:end_idx], 1):
                ws.cell(row=row, column=col+1, value=item)
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialfile="List_Teracak.xlsx" if shuffle_var.get() else "List_Original.xlsx"
        )
        
        if output_file:
            try:
                wb.save(output_file)
                show_success_message(output_file, len(items), num_columns, items_per_column)
            except Exception as e:
                messagebox.showerror("Error", f"Gagal menyimpan file:\n{str(e)}")
            
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan sistem:\n{str(e)}")

def show_success_message(filename, item_count, column_count, items_per_column):
    success_msg = f"""
    File Excel berhasil dibuat!
    
    Lokasi: {filename}
    Total Item: {item_count}
    Kolom dibuat: {column_count}
    Item per kolom: {items_per_column}
    Status: {'TERACAK' if shuffle_var.get() else 'ORIGINAL'}
    Mode: {'ITEM PER KOLOM' if items_per_column_mode.get() else 'JUMLAH KOLOM'}
    """
    messagebox.showinfo("Sukses", success_msg.strip())

def clear_input():
    text_input.delete("1.0", tk.END)
    status_label.config(text="")
    count_label.config(text="Total Item: 0")
    column_info_label.config(text="")

def toggle_mode():
    if items_per_column_mode.get():
        mode_label.config(text="Item per Kolom:")
        column_mode_label.config(text="Mode: Item per Kolom", foreground="#2196F3")
    else:
        mode_label.config(text="Jumlah Kolom:")
        column_mode_label.config(text="Mode: Jumlah Kolom", foreground="#4CAF50")
    update_item_count()

# GUI Setup
root = tk.Tk()
root.title("List Shuffler Pro")
root.geometry("350x600")  # Lebih sempit dan lebih tinggi
root.resizable(False, False)

# Style Configuration
style = ttk.Style()
style.theme_use('clam')  # Clean, minimalist theme

# Color scheme
bg_color = "#f5f5f5"
primary_color = "#2196F3"
success_color = "#4CAF50"
warning_color = "#FF9800"
text_color = "#333333"

style.configure(".", background=bg_color, foreground=text_color)
style.configure("TFrame", background=bg_color)
style.configure("TLabel", background=bg_color, font=('Segoe UI', 9))
style.configure("TButton", font=('Segoe UI', 9), padding=5, borderwidth=1)
style.configure("Bold.TButton", font=('Segoe UI', 9, 'bold'))
style.map("Bold.TButton",
          foreground=[('pressed', 'white'), ('active', 'white')],
          background=[('pressed', '#45a049'), ('active', success_color)])
style.configure("TEntry", padding=5, relief="flat")
style.configure("TCheckbutton", background=bg_color)
style.configure("Vertical.TScrollbar", arrowsize=12)

# Main Frame
main_frame = ttk.Frame(root, padding=10)
main_frame.pack(expand=True, fill=tk.BOTH)

# Text Input Area
input_frame = ttk.Frame(main_frame)
input_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

input_header = ttk.Frame(input_frame)
input_header.pack(fill=tk.X, pady=(0, 5))
ttk.Label(input_header, text="Paste list Anda (1 item per baris):").pack(side=tk.LEFT)
count_label = ttk.Label(input_header, text="Total Item: 0", foreground=primary_color)
count_label.pack(side=tk.RIGHT)

text_input = scrolledtext.ScrolledText(
    input_frame, 
    height=10,
    wrap=tk.WORD,
    font=('Segoe UI', 9),
    padx=8,
    pady=8,
    bd=1,
    relief=tk.SOLID,
    highlightthickness=1,
    highlightbackground="#e0e0e0"
)
text_input.pack(expand=True, fill=tk.BOTH)
text_input.bind("<KeyRelease>", lambda e: update_item_count())

# Control Panel - Vertical Layout
control_panel = ttk.Frame(main_frame)
control_panel.pack(fill=tk.X, pady=(5, 10))

# Mode toggle
mode_row = ttk.Frame(control_panel)
mode_row.pack(fill=tk.X, pady=2)
items_per_column_mode = tk.BooleanVar(value=False)
mode_check = ttk.Checkbutton(
    mode_row,
    text="Mode Item/Kolom",
    variable=items_per_column_mode,
    command=toggle_mode
)
mode_check.pack(side=tk.LEFT)

# Column input
input_row = ttk.Frame(control_panel)
input_row.pack(fill=tk.X, pady=2)
mode_label = ttk.Label(input_row, text="Jumlah Kolom:")
mode_label.pack(side=tk.LEFT)

entry_columns = ttk.Entry(input_row, width=6)
entry_columns.pack(side=tk.LEFT, padx=5)
entry_columns.insert(0, "3")

column_info_label = ttk.Label(input_row, text="", foreground="#757575", font=('Segoe UI', 8))
column_info_label.pack(side=tk.LEFT)

# Shuffle option
shuffle_row = ttk.Frame(control_panel)
shuffle_row.pack(fill=tk.X, pady=2)
shuffle_var = tk.BooleanVar(value=False)
shuffle_check = ttk.Checkbutton(
    shuffle_row,
    text="Acak List",
    variable=shuffle_var,
    command=lambda: status_label.config(
        text="Mode pengacakan AKTIF" if shuffle_var.get() else "",
        foreground=success_color if shuffle_var.get() else bg_color
    )
)
shuffle_check.pack(side=tk.LEFT)

# Button row
button_row = ttk.Frame(control_panel)
button_row.pack(fill=tk.X, pady=(10, 5))
ttk.Button(button_row, text="Clear", command=clear_input, width=10).pack(side=tk.LEFT, padx=5)
ttk.Button(button_row, text="Copy", command=lambda: [root.clipboard_clear(), 
          root.clipboard_append(text_input.get("1.0", tk.END)), 
          status_label.config(text="Teks disalin!", foreground=primary_color)], width=10).pack(side=tk.LEFT)

# Mode indicator
column_mode_label = ttk.Label(
    main_frame,
    text="Mode: Jumlah Kolom",
    foreground=success_color,
    font=('Segoe UI', 8, 'italic')
)
column_mode_label.pack()

# Process Button
process_btn = ttk.Button(
    main_frame,
    text="GENERATE EXCEL",
    command=process_list,
    style="Bold.TButton",
    width=20
)
process_btn.pack(pady=(10, 5))

# Status Bar
status_label = ttk.Label(
    main_frame,
    text="",
    foreground=success_color,
    font=('Segoe UI', 8),
    anchor=tk.CENTER
)
status_label.pack(fill=tk.X)

# Bind entry_columns to update item count when changed
entry_columns.bind("<KeyRelease>", lambda e: update_item_count())

# Run the application
try:
    root.mainloop()
except Exception as e:
    messagebox.showerror("Fatal Error", f"Aplikasi tidak dapat dijalankan:\n{str(e)}")