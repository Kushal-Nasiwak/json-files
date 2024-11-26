import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd  


Ankens_bango_builder = {}
# Function to clear placeholder text when user focuses on the entry widget
def clear_placeholder(event, entry, placeholder):
    if entry.get() == placeholder:
        entry.delete(0, tk.END)
        entry.config(foreground="black")


# Function to restore placeholder text if entry is empty
def restore_placeholder(event, entry, placeholder):
    if entry.get() == "":
        entry.insert(0, placeholder)
        entry.config(foreground="gray")


# Function to load and display values in a structured format in the text area, then clear the entries
def load_values():
    anken_bango = anken_entry.get()
    builder_id = builder_id_entry.get()

    # Check if placeholders are still present; if so, replace them with an empty string
    if anken_bango == anken_placeholder:
        anken_bango = ""
    if builder_id == builder_placeholder:
        builder_id = ""

    # Append headers if they aren't already present
    if text_area.compare("end-1c", "==", "1.0"):
        text_area.insert(tk.END, f"{'Anken Bango':^30} {'Builder ID':^30}\n")
        text_area.insert(tk.END, "-" * 80 + "\n")  # Divider line

    # Append the entered values in a formatted way
    if anken_bango and builder_id:
        Ankens_bango_builder[str(anken_bango)] = str(builder_id)
        text_area.insert(tk.END, f"{anken_bango:^30} {builder_id:^30}\n")
    

    # Clear the entry fields and reset placeholders
    anken_entry.delete(0, tk.END)
    builder_id_entry.delete(0, tk.END)
    restore_placeholder(None, anken_entry, anken_placeholder)
    restore_placeholder(None, builder_id_entry, builder_placeholder)


# Function to browse and upload an Excel file
def upload_excel():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsm"), ("All Files", "*.*")]
    )
    if file_path:
        excel_file_entry.delete(0, tk.END)
        excel_file_entry.insert(0, file_path)
        read_excel_file(file_path)


# Function to read and display Anken Name and Builder ID columns from the Excel file
def read_excel_file(file_path):
    try:
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(file_path,dtype=str)
        # Check if the required columns exist
        if "Anken Number" in df.columns and "Builder Code" in df.columns:
            anken_data = df["Anken Number"]
            builder_data = df["Builder Code"]

            # Display headers if not already present
            if text_area.compare("end-1c", "==", "1.0"):
                text_area.insert(tk.END, f"{'Anken Bango':^30} {'Builder ID':^30}\n")
                text_area.insert(tk.END, "-" * 80 + "\n")  # Divider line

            # Display the data row by row
            for anken, builder in zip(anken_data, builder_data):
                Ankens_bango_builder[anken] = builder
                text_area.insert(tk.END, f"{anken:^30} {builder:^30}\n")

        else:
            messagebox.showerror("Error", "The file does not contain 'Anken Name' and 'Builder ID' columns.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read the Excel file:\n{str(e)}")

# Function to start the Excel check
def Start_Excel_check():
    print("It works")
    
   

# Initialize the main window
root = tk.Tk()
root.title("Excel Check")
root.geometry("1000x700")  # Increased height for extra area
root.configure(bg="#d6e0f0")

# Logo placeholder
logo_frame = tk.Frame(root, bg="#d6e0f0", height=50)
logo_frame.pack(fill='x', padx=10, pady=5)
logo_label = tk.Label(logo_frame, text="Nasiwak", font=("Arial", 12, "bold"), bg="#d6e0f0")
logo_label.pack(side='left')

# Title
title_label = tk.Label(root, text="Excel Check", font=("Arial", 14, "bold"), bg="#d6e0f0", fg="#333333")
title_label.pack(pady=(10, 5))

# Instruction label
instruction_label = tk.Label(
    root, text="Enter Anken Bango and the Builder Id, Click on the START button to start Excel Check.",
    font=("Arial", 10), bg="#d6e0f0", fg="#555555"
)
instruction_label.pack()

# Entry Frame
entry_frame = tk.Frame(root, bg="#d6e0f0")
entry_frame.pack(pady=(10, 10))

# Placeholder texts
anken_placeholder = "Enter Anken Bango..."
builder_placeholder = "Enter Builder Id..."

# Anken Bango entry with placeholder
anken_entry = ttk.Entry(entry_frame, width=30, foreground="gray")
anken_entry.insert(0, anken_placeholder)
anken_entry.grid(row=0, column=0, padx=10, pady=5)

anken_entry.bind("<FocusIn>", lambda event: clear_placeholder(event, anken_entry, anken_placeholder))
anken_entry.bind("<FocusOut>", lambda event: restore_placeholder(event, anken_entry, anken_placeholder))

# Builder Id entry with placeholder
builder_id_entry = ttk.Entry(entry_frame, width=30, foreground="gray")
builder_id_entry.insert(0, builder_placeholder)
builder_id_entry.grid(row=0, column=1, padx=10, pady=5)

builder_id_entry.bind("<FocusIn>", lambda event: clear_placeholder(event, builder_id_entry, builder_placeholder))
builder_id_entry.bind("<FocusOut>", lambda event: restore_placeholder(event, builder_id_entry, builder_placeholder))

# Load button
load_button = ttk.Button(entry_frame, text="LOAD", command=load_values)
load_button.grid(row=0, column=2, padx=10, pady=5)

# Excel file upload frame
excel_frame = tk.Frame(root, bg="#d6e0f0")
excel_frame.pack(pady=(10, 10))

# Upload Excel button
upload_button = ttk.Button(excel_frame, text="Upload Excel File", command=upload_excel)
upload_button.grid(row=0, column=0, padx=10, pady=5)

# Entry box to show the uploaded file path
excel_file_entry = ttk.Entry(excel_frame, width=50, state="normal")
excel_file_entry.grid(row=0, column=1, padx=10, pady=5)

# Main text area for displaying entries in a structured way
text_area = tk.Text(root, height=15, width=80, bg="#e6eefc", font=("Courier", 10))
text_area.pack(pady=10)

# Start button
start_button = ttk.Button(root, text="START", width=10,command=Start_Excel_check)
start_button.pack(pady=10)

# Footer
footer_label = tk.Label(root, text="Nasiwak Services Pvt Ltd     v9.0.0", bg="#d6e0f0", fg="#333333")
footer_label.pack(side='bottom', pady=(5, 0))

# Run the main event loop
root.mainloop()
