import pandas as pd
import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
import subprocess
from tkinter import messagebox

# File path
excel_file = 'papers.xlsx'

# Load Excel file
def load_excel():
    try:
        return pd.read_excel(excel_file)
    except FileNotFoundError:
        return pd.DataFrame(columns=['Year', 'Title', 'Authors', 'Publication', 'Link', 'Summary', 'Techniques'])

# Save DataFrame to Excel
def save_excel(df):
    df.to_excel(excel_file, index=False)

# Update displayed table
def refresh_table():
    for row in tree.get_children():
        tree.delete(row)
    if df.empty:
        tree.insert('', 'end', values=["No papers"] + [''] * (len(columns) - 1))
    else:
        for idx, row in df.iterrows():
            tree.insert('', 'end', values=list(row))

# Clear form fields
def clear_fields():
    year_combo.set('')
    title_entry.delete(0, tk.END)
    authors_entry.delete(0, tk.END)
    publication_entry.delete(0, tk.END)
    link_entry.delete(0, tk.END)
    summary_text.delete('1.0', tk.END)
    techniques_text.delete('1.0', tk.END)

# Add a new paper
def add_paper():
    global df
    new_row = {
        'Year': '',
        'Title': '',
        'Authors': '',
        'Publication': '',
        'Link': '',
        'Summary': '',
        'Techniques': ''
    }
    df = df.append(new_row, ignore_index=True)
    save_excel(df)
    refresh_table()
    tree.selection_set(tree.get_children()[-1])  # Highlight the new blank entry
    clear_fields()

# Delete selected paper
def delete_paper():
    global df
    selected_item = tree.selection()
    if selected_item:
        paper_index = tree.index(selected_item[0])
        df = df.drop(df.index[paper_index]).reset_index(drop=True)
        save_excel(df)
        refresh_table()
        # Highlight the previous entry
        children = tree.get_children()
        if children:
            prev_index = max(0, paper_index - 1)
            tree.selection_set(children[prev_index])

# Modify selected paper
def modify_paper():
    global df
    selected_item = tree.selection()
    if selected_item:
        paper_index = tree.index(selected_item[0])
        df.loc[paper_index, 'Year'] = year_combo.get()
        df.loc[paper_index, 'Title'] = title_entry.get()
        df.loc[paper_index, 'Authors'] = authors_entry.get()
        df.loc[paper_index, 'Publication'] = publication_entry.get()
        df.loc[paper_index, 'Link'] = link_entry.get()
        df.loc[paper_index, 'Summary'] = summary_text.get('1.0', tk.END).strip()
        df.loc[paper_index, 'Techniques'] = techniques_text.get('1.0', tk.END).strip()
        save_excel(df)
        refresh_table()

# Populate fields with selected paper data
def populate_fields(event):
    selected_item = tree.selection()
    if selected_item:
        paper_index = tree.index(selected_item[0])
        paper = df.iloc[paper_index]
        clear_fields()
        year_combo.set(paper['Year'])
        title_entry.insert(0, paper['Title'])
        authors_entry.insert(0, paper['Authors'])
        publication_entry.insert(0, paper['Publication'])
        link_entry.insert(0, paper['Link'])
        summary_text.insert('1.0', paper['Summary'])
        techniques_text.insert('1.0', paper['Techniques'])

# Run generate_bib.py and update HTML
def update_html():
    try:
        subprocess.run(['python', 'generate_bib.py'], check=True)
        messagebox.showinfo("Success", "HTML updated successfully")
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", "Failed to update HTML. Please check the script.")

# Safely close the application
def close_app():
    save_excel(df)
    root.destroy()

# GUI Setup
root = ThemedTk(theme='equilux')  # Dark theme
root.title("Meta-complexity Bibliography Editor")
root.geometry("1024x768")
root.configure(bg='#2b2b2b')

# Load data
df = load_excel()

# Treeview for listing papers
columns = ['Year', 'Title', 'Authors', 'Publication', 'Link', 'Summary', 'Techniques']
tree_frame = ttk.Frame(root)
tree_frame.pack(pady=10, fill='both', expand=True, padx=20)

tree = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='browse')
tree.pack(side='left', fill='both', expand=True)

# Add scrollbars to the treeview
x_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=tree.xview)
x_scrollbar.pack(side='bottom', fill='x')
tree.configure(xscrollcommand=x_scrollbar.set)

y_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
y_scrollbar.pack(side='right', fill='y')
tree.configure(yscrollcommand=y_scrollbar.set)

# Configure tree columns and headings
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor='center', width=150)

tree.bind('<<TreeviewSelect>>', populate_fields)

# Form frame for entry fields
form_frame = ttk.Frame(root)
form_frame.pack(pady=10, fill='x', padx=20)

# Entry fields
fields = ['Year', 'Title', 'Authors', 'Publication', 'Link', 'Summary', 'Techniques']
entries = []

# Year dropdown
year_label = ttk.Label(form_frame, text="Year:", style="TLabel")
year_label.grid(row=0, column=0, sticky='w', padx=5, pady=5)
year_combo = ttk.Combobox(form_frame, values=list(range(1900, 2051)), state='readonly', width=37)
year_combo.grid(row=0, column=1, sticky='w', padx=5, pady=5)

# Title, Authors, Publication, and Link entries
for idx, field in enumerate(fields[1:5]):
    label = ttk.Label(form_frame, text=f"{field}:", style="TLabel")
    label.grid(row=idx + 1, column=0, sticky='w', padx=5, pady=5)
    entry = ttk.Entry(form_frame, width=40)
    entry.grid(row=idx + 1, column=1, sticky='w', padx=5, pady=5)
    entries.append(entry)

title_entry, authors_entry, publication_entry, link_entry = entries

# Multiline textboxes for Summary and Techniques
for idx, field in enumerate(fields[5:]):
    label = ttk.Label(form_frame, text=f"{field}:", style="TLabel")
    label.grid(row=idx + 5, column=0, sticky='w', padx=5, pady=5)
    text_box = tk.Text(form_frame, height=5, width=40, wrap='word', bg='#3c3f41', fg='white', insertbackground='white')
    text_box.grid(row=idx + 5, column=1, sticky='w', padx=5, pady=5)
    if field == 'Summary':
        summary_text = text_box
    else:
        techniques_text = text_box

# Buttons
button_frame = ttk.Frame(root)
button_frame.pack(pady=10, fill='x', padx=20)

add_button = ttk.Button(button_frame, text="Add Paper", command=add_paper)
add_button.grid(row=0, column=0, padx=5, pady=5)

modify_button = ttk.Button(button_frame, text="Modify Paper", command=modify_paper)
modify_button.grid(row=0, column=1, padx=5, pady=5)

delete_button = ttk.Button(button_frame, text="Delete Paper", command=delete_paper)
delete_button.grid(row=0, column=2, padx=5, pady=5)

update_html_button = ttk.Button(button_frame, text="Update HTML", command=update_html)
update_html_button.grid(row=0, column=3, padx=5, pady=5)

close_button = ttk.Button(button_frame, text="Close App", command=close_app)
close_button.grid(row=0, column=4, padx=5, pady=5)

# Style
style = ttk.Style(root)
style.configure('Treeview', background='#3c3f41', foreground='white', fieldbackground='#3c3f41', rowheight=25)
style.map('Treeview', background=[('selected', '#4a4e51')], foreground=[('selected', 'white')])
style.configure('TLabel', background='#2b2b2b', foreground='white')
style.configure('TEntry', background='#3c3f41', foreground='white', fieldbackground='#3c3f41', insertcolor='white')
style.configure('TCombobox', fieldbackground='#3c3f41', background='#3c3f41', foreground='white')
style.configure('TButton', background='#4a4e51', foreground='white')

# Populate table
refresh_table()

# Run the application
root.mainloop()
