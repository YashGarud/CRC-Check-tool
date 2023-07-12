#GENERATE CODE
#======================================================================
import os
import crcmod.predefined
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

# Create a CRC-32 function with the desired polynomial
crc32_func = crcmod.predefined.mkCrcFun('crc-32')

# Define the root window for the Tkinter file dialog
root = tk.Tk()
root.withdraw()

# Prompt the user to select the Excel file to process
file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])

# Create the main window for the UI
window = tk.Tk()
window.title("Generate CRC")

# Define the callback function for the "Generate CRC" button
def generate_crc():
    # Get the number of columns to use for generating the CRC
    num_columns_str = num_columns_entry.get()
    num_columns = int(num_columns_str)

    # Automatically populate the columns entry widget
    columns_entry.delete(0, tk.END)
    columns_entry.insert(0, ",".join(str(i) for i in range(1, num_columns + 1)))

    # Read the Excel file into a pandas dataframe
    df = pd.read_excel(file_path)

    # Get the subset of the dataframe with the specified columns
    columns_str = columns_entry.get()
    columns = [int(col.strip()) - 1 for col in columns_str.split(",")][:num_columns]
    column_range = df.iloc[:, columns]

    # Calculate the checksum for each row and add it to the dataframe
    now = datetime.now().strftime('%Y%m%d_%H%M%S')
    column_name = f'CRC_{now:}_{num_columns:02d}_{columns_str.replace(",", "_")}'
    checksums = column_range.apply(lambda row: '{:08x}'.format(crc32_func(bytes(str(row), "utf-16"))), axis=1)

    # Add the checksum column to the original dataframe
    df[column_name] = checksums

    # Save the changes to the Excel file
    df.to_excel(file_path, index=False)

    # Display success message
    success_label.config(text=f"CRC column '{column_name}' added to file '{os.path.basename(file_path)}'")

# Create the UI elements
num_columns_label = tk.Label(window, text="Number of columns:")
num_columns_entry = tk.Entry(window)
columns_label = tk.Label(window, text="Columns (comma-separated):")
columns_entry = tk.Entry(window)
generate_button = tk.Button(window, text="Generate CRC", command=generate_crc)
success_label = tk.Label(window, text="")

# Pack the UI elements
num_columns_label.pack()
num_columns_entry.pack()
columns_label.pack()
columns_entry.pack()
generate_button.pack()
success_label.pack()

# Start the UI main loop
window.mainloop()
