import openpyxl
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import os

def search_excel(filename, word):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    
    results = []
    
    columns_found = False
    for row in sheet.iter_rows(values_only=True):
        if not columns_found:
            if "table" in str(row[0]).lower():
                columns_found = True
        else:
            table_name = row[0].strip() if row[0] else ""  
            description = row[2].strip() if row[2] else ""  
            if word.lower() in table_name.lower() or word.lower() in description.lower():
                results.append((table_name, description))
    
    return results

def search_button_click():
    word = entry.get()
    results = search_excel(filename, word)
    
    if results:
        result_text.config(state=tk.NORMAL)
        result_text.delete("1.0", tk.END)
        for result in results:
            table_name = result[0]
            columns = result[1]
            
            # Find the position of the "-------- columns --------" header
            index = columns.lower().find("-------- columns --------")
            
            if index != -1:
                columns_data = "\n" + "    " + columns[index + len("-------- columns --------"):].replace(":", ":\n")
            else:
                columns_data = "No columns found."
            
            # Insert data into two columns
            result_text.insert(tk.END, f"{table_name:<30}", "table")  # Column 1
            
            # Search and highlight the word in the columns data
            columns_data_lower = columns_data.lower()
            word_lower = word.lower()
            word_start = columns_data_lower.find(word_lower)
            while word_start != -1:
                result_text.insert(tk.END, columns_data[:word_start], "columns")
                result_text.insert(tk.END, columns_data[word_start:word_start+len(word)], "bold")  # Highlighted word
                columns_data = columns_data[word_start+len(word):]
                columns_data_lower = columns_data.lower()
                word_start = columns_data_lower.find(word_lower)
            result_text.insert(tk.END, columns_data, "columns")                            
        result_text.config(state=tk.DISABLED)
    else:
        messagebox.showinfo("No Matches", "No matches found.")

#filename = 'data/PPAS Tables.xlsx'
filename = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'PPAS Tables.xlsx')

# Create the main window
root = tk.Tk()
root.title("Excel Search")

# Create and place GUI components
label = tk.Label(root, text="Enter a word to search:")
label.pack()

entry = tk.Entry(root)
entry.pack()

search_button = tk.Button(root, text="Search", command=search_button_click)
search_button.pack()

version = tk.Label(root, text="v1.3.0")
version.pack()

# Create a scrolled text widget with horizontal and vertical scrollbars
scrollbar_y = ttk.Scrollbar(root, orient=tk.VERTICAL)
scrollbar_x = ttk.Scrollbar(root, orient=tk.HORIZONTAL)

result_text = tk.Text(root, wrap=tk.WORD, state=tk.DISABLED, yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
scrollbar_y.config(command=result_text.yview)
scrollbar_x.config(command=result_text.xview)

scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
result_text.pack(fill=tk.BOTH, expand=True)

# Configure the tags for formatting
result_text.tag_configure("table", font=("TkDefaultFont", 12, "bold"))
result_text.tag_configure("columns", font=("TkDefaultFont", 10))
result_text.tag_configure("bold", font=("TkDefaultFont", 12, "bold"))

# Start the GUI event loop
root.mainloop()
