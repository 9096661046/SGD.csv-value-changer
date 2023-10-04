import os
import csv
import openpyxl
import tkinter as tk
from tkinter import filedialog

def ask_for_path():
    # Create the Tkinter root window
    root = tk.Tk()

    # Set the window size dynamically based on screen resolution
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.geometry(f"{int(screen_width*0.5)}x{int(screen_height*0.5)}+{int(screen_width*0.2)}+{int(screen_height*0.2)}")

    # Hide the root window
    root.withdraw()

    # Show the folder selection dialog and return the selected path
    path = filedialog.askdirectory(title="Select path of the SGD.csv files")
    return path

def ask_for_xlsx_file():
    # Create the Tkinter root window
    root = tk.Tk()

    # Set the window size dynamically based on screen resolution
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.geometry(f"{int(screen_width*0.5)}x{int(screen_height*0.5)}+{int(screen_width*0.2)}+{int(screen_height*0.2)}")

    # Hide the root window
    root.withdraw()

    # Show the file selection dialog and return the selected file path
    file_path = filedialog.askopenfilename(title="Select path of replacement.xlsx file", filetypes=(("XLSX files", "*.xlsx"), ("All files", "*.*")))
    return file_path



path = ask_for_path()
replacements_file = ask_for_xlsx_file()

# Load the Excel file with the replacement information
wb = openpyxl.load_workbook(replacements_file)
ws = wb.active

# Define a function to read and yield rows from a CSV file
def read_csv_rows(filename):
    with open(filename, 'r', newline='') as file:
        reader = csv.reader(file)
        for row in reader:
            yield row

# Define a function to write rows to a CSV file
def write_csv_rows(filename, rows):
    with open(filename, 'w', newline='') as file:
        writer = csv.writer(file)
        for row in rows:
            writer.writerow(row)
            
def shift_lines(csv_path):
    # Read the contents of the CSV file
    with open(csv_path, 'r') as csv_file:
        content = csv_file.readlines()

    # Iterate through each line of the file and shift as needed
    for i in range(len(content)):
        if i >= 2 and content[i].startswith('--') and content[i-1].startswith('C') and content[i+1].startswith('V'):
            content[i], content[i-1] = content[i-1], content[i]

    # Write the modified contents back to the file
    with open(csv_path, 'w') as csv_file:
        csv_file.writelines(content)

def process_csv_files(path):
    for file_name in os.listdir(path):
        if file_name.endswith('.csv'):
            file_path = os.path.join(path, file_name)
            shift_lines(file_path)


# Get the list of CSV files to search and replace in
csv_files = [os.path.join(path, f) for f in os.listdir(path) if f.endswith('.csv')]

# Loop through each CSV file
for csv_file in csv_files:
    # Load the rows of the CSV file
    with open(csv_file, 'r', newline='') as file:
        rows = list(csv.reader(file))
    
    # Loop through each replacement specified in the Excel file
    for row in ws.iter_rows(min_row=2, max_col=3):
        search_line = row[0].value
        new_line = row[1].value
        comment = row[2].value
        
        # Skip this replacement if the search and replace content are the same
        if search_line == new_line:
            continue
        
        # Replace the matching rows in the CSV file and add the comment
        new_rows = []
        for i, row in enumerate(read_csv_rows(csv_file)):
            if row == search_line.split(','):
                if row == new_line.split(','):
                    new_rows.append(row)
                else:
                    new_rows.append([comment])
                    new_rows.append(new_line.split(','))
            else:
                new_rows.append(row)
    
        # Write the updated CSV file
        write_csv_rows(csv_file, new_rows)

# Example usage:
process_csv_files(path)