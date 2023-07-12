#COMPARE

#=============================================================================================================
import tkinter as tk
from tkinter import filedialog
import pandas as pd

def compare(file, column1, column2):
    df = pd.read_excel(file)
    df['match'] = df[column1] == df[column2]
    df['change'] = df.apply(lambda row: 'Changed' if not row['match'] else '', axis=1)
    return df

def browse_file():
    file_path = filedialog.askopenfilename()
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def run_comparison():
    file_path = file_entry.get()
    column1 = column1_entry.get()
    column2 = column2_entry.get()
    df = compare(file_path, column1, column2)
    output_text.delete('1.0', tk.END)
    output_text.insert(tk.END, str(df))
    save_file(df)
    generate_report(df)

def save_file(df):
    file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    if file_path:
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
        df.to_excel(writer, index=False)
        writer.save()

def generate_report(df):
    changed_rows = df[df['change'] == 'Changed']
    report = ''
    if not changed_rows.empty:
        report += 'The following rows have been changed:\n\n'
        report += str(changed_rows) + '\n\n'
    else:
        report += 'No rows have been changed.\n\n'
    report += 'Summary:\n\n'
    report += 'Total rows: ' + str(len(df)) + '\n'
    report += 'Matching rows: ' + str(df['match'].sum()) + '\n'
    report += 'Non-matching rows: ' + str(len(df) - df['match'].sum()) + '\n'
    report_text.delete('1.0', tk.END)
    report_text.insert(tk.END, report)

# Create the tkinter GUI
root = tk.Tk()

# Add widgets to browse for the file
file_label = tk.Label(root, text='Select Excel file to compare:')
file_label.pack()
file_entry = tk.Entry(root)
file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
file_button = tk.Button(root, text='Browse', command=browse_file)
file_button.pack(side=tk.LEFT)

# Add widgets to enter the column names
column1_label = tk.Label(root, text='Enter first column name:')
column1_label.pack()
column1_entry = tk.Entry(root)
column1_entry.pack()
column2_label = tk.Label(root, text='Enter second column name:')
column2_label.pack()
column2_entry = tk.Entry(root)
column2_entry.pack()

# Add a button to run the comparison
compare_button = tk.Button(root, text='Compare', command=run_comparison)
compare_button.pack()

# Add a text box to display the output
output_label = tk.Label(root, text='Comparison output:')
output_label.pack()
output_text = tk.Text(root)
output_text.pack(fill=tk.BOTH, expand=True)

# Add a text box to display the report
report_label = tk.Label(root, text='Comparison report:')
report_label.pack()
report_text = tk.Text(root)
report_text.pack(fill=tk.BOTH, expand=True)

# Start the tkinter main loop
root.mainloop()
