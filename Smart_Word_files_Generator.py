import tkinter as tk
import win32com.client
import os
import tkinter.messagebox
import docx2txt as dx
import re
import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


global py_path, cy_path, Application
global Checkbutton1, Check_Path, off_color, on_color

# Initializing tCheckbutton1kinter window--
window=tk.Tk()
window.title("Smart-Templates-Updater")
Rwidth = 350
Rheight = 150
window.geometry('350x150')
window.configure(bg='#856ff8')

Checkbutton1 = tk.StringVar (window)
Checkbutton1.set (0)

# to check if the check box is checked or not------
def compare():

    if Checkbutton1.get() == "1":
        Check_Path["fg"]= on_color


    else:
        Check_Path["fg"] - off_color


# function to be called to compare PY and CY Templates--
def Compare_word_docs():
    global Original_doc, Comparison_doc

    Original_doc=dx.process(py_path.get())
    Comparison_doc = dx.process(cy_path.get())




    # Create the Application word
    Application = win32com.client.gencache. EnsureDispatch("Word.Application")
    Application.Visible = False



    # Compare documents
    Application.CompareDocuments (Application.Documents.Open(py_path.get()), Application.Documents.Open(cy_path.get()))



    Application.ActiveDocument. SaveAs (os.path.join(os.path.expanduser('~'),'Documents',"Comparison.docx"))


    Application.Quit()
    tkinter.messagebox.showinfo('Thank you! ', 'File has been created, Please check Document folder')

if os.path.exists("Comparison.docx") == True:
    os. remove ("Comparison.docx")




# Fetching fields from the templates which needs to be updated---
def extract_fields():
    global var_list

    doc_text = str(dx.process (py_path.get()))
    search_all_fields = re.finditer(r"\[(.*?)*\]", doc_text)

    var_list =[]

    for item in search_all_fields:
        var_list.append(item.group())
        display_fields()
    display_make_changes()

# Display dynamic fields for each updatable field identified by the code to be updated by the user-------
def display_fields():
    window.geometry('800x500')
    global row_n
    global col_n
    global list_user_inputs
    list_user_inputs = []

    row_n = 12
    col_n = 0

    for var_index in range(len(var_list)):

        my_var = tk.StringVar ()
        list_user_inputs.append (my_var)
        field_label =tk.Label (window, text=var_list[var_index])


        field_label.grid(row=row_n, column=col_n, padx=5, pady=5)
        field_value = tk.Entry(window, textvariable=list_user_inputs[var_index])
        field_value.grid(row=row_n, column=col_n + 1, padx=5, pady=5)

        row_n =row_n + 1

# get user input field values once the commit changes button is clicked
def  get_field_vals():
    global new_doc_content
    new_doc_content = str(dx.process(py_path.get()))
    for item in range(len(list_user_inputs)):
        new_doc_content = new_doc_content.replace(str(var_list[item]), str(list_user_inputs[item].get()))

    Submit_btn = tk.Button(window, text="Click to generate Report", command=update,bg='white')
    Submit_btn.grid(row=row_n + 2, column=1, columnspan=15, padx=5, pady=5)

# To extract a copy of final template with updated field values(taken from user as an input)------

def update():
    document = docx.Document()
    para = document.add_paragraph(new_doc_content)
    para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    para = document.styles['Normal'].paragraph_format
    document.save(os.path.join(os.path.expanduser('~'),'Documents',"Final.docx"))
    tkinter.messagebox.showinfo('Thank you!', 'File has been created')


# To Show commit changes button once the user clicks on the Update Template button--

def display_make_changes():
    Commit_btn = tk.Button(window, text="Commit Changes", command=get_field_vals, bg='white')
    Commit_btn.grid(row=row_n + 2, column=1, columnspan=15, padx=5, pady=5)

# Displaying Initial Screen View here--
py_path_var = tk.StringVar()
py_path = tk.Entry(window, textvariable=py_path_var)
py_path.grid(row=1, column=2, columnspan=10, padx=5, pady=5, sticky=tk.W)
py_path_label =tk.Label(window, text="PY File Path", fg='black', bg='white')
py_path_label.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

cy_path_label = tk.Label(window, text="CY File Path", fg='black', bg='white')
cy_path_label.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
cy_path_var = tk.StringVar()
cy_path=tk.Entry(window, textvariable=cy_path_var)
cy_path.grid(row=2, column=2, columnspan=10, padx=5, pady=5, sticky=tk.W)

run_comparison = tk.Button (window, text="Comparison Analysis", command=Compare_word_docs, bg="white")
run_comparison.grid(row=0, column=22, rowspan=4, columnspan=20, padx=5, pady=5, sticky="NEWS")
off_color = "white"
on_color = "black"

Check_Path = tk.Checkbutton (window, text="click here to add path", variable=Checkbutton1, onvalue=1, offvalue=0,
height=1,width=10, command=compare, fg=off_color, bg='#856ff8')
Check_Path.grid(row=3, column=1, rowspan=4, columnspan=15, padx=5, pady=5, sticky="NEWS")
edit_btn = tk.Button(window, text="Update New Template here", command=extract_fields, bg='white')
edit_btn.grid(row=10, column=1, columnspan=15, padx=5, pady=5)
window.mainloop()
