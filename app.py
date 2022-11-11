# from __future__ import print_function
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter.filedialog import askopenfile

import pandas as pd

from mailmerge import MailMerge
from datetime import date
import sys
import pandas as pd
import os
from pathlib import Path
import glob
import comtypes.client
from PyPDF2 import PdfFileMerger


# Create a tkinter window
window = tk.Tk()
window.title("Aaruush OD App")
window.minsize(width=500, height=300)

window.geometry("700x450")

data = tk.StringVar(value="")
template = tk.StringVar(value="")
size = tk.StringVar(value="")
status = tk.StringVar(value="Select Data File and Template to Generate ODs")

def check():
    if(data.get()!="" and template.get()!=""):
        g["state"]="normal"
    else:
        g["state"]="disabled"

def open_data():
    file1 = filedialog.askopenfile(mode='r', filetypes=[('Excel Sheets', '*.xlsx')])
    if file1:
        global oddata
        data.set(file1.name)
        oddata = pd.read_excel(data.get())
        # size.set(file)
        size.set(str(len(oddata)))
        status.set(f"Data Size: {size.get()}")
        file1.close()

def open_template():
    file2 = filedialog.askopenfile(mode='r', filetypes=[('Word Documents', '*.docx')])
    if file2:
        print(file2.name)
        template.set(file2.name)
        file2.close()

def generate():
    if(data.get() and template.get()):
        global oddata
        oddata = oddata.to_dict()
        lst = {}
        final = {}
        for i in range(len(oddata['name'])):
            now = {}
            for key,value in oddata.items():
                now[key] = value[i]
            lst[i]=now
        
        size.set(str(len(lst.values())))
        for i,items in enumerate(lst.values()):

            odtemplate = template.get()
            status.set(f"Generating OD {i+1} of {size.get()}")
            document = MailMerge(odtemplate)
            # print(document.get_merge_fields())
            document.merge(
                Name=items['name'],
                Regd=items['regd'],
                dates=items['dates'],
                Dept=items['dept'],
                hours=items['hours']
            )
            try:
                os.mkdir('odtemp')
            except FileExistsError:
                pass
            document.write('odtemp/'+str(i)+items['name']+'.docx')
        combine()
    else:
        status.set("Select Data File and Template to Generate ODs")

def docxs_to_pdf(filenames):
	"""Converts all word files in pdfs and append them to pdfslist"""
	word = comtypes.client.CreateObject('Word.Application')
	pdfslist = PdfFileMerger()
	x = 0
	for i,f in enumerate(filenames):
		input_file = os.path.abspath(f);status.set(f"Converting OD {i+1} of {size.get()}")
		output_file = os.path.abspath(f"odtemp/{str(x)} OD.pdf")
		# loads each word document
        
		doc = word.Documents.Open(input_file)
		doc.SaveAs(output_file, FileFormat=16+1)
		doc.Close() # Closes the document, not the application
		pdfslist.append(open(output_file, 'rb'))
		x += 1

	word.Quit()
	return pdfslist

def joinpdf(pdfs):
	"""Unite all pdfs"""
	with open("OD.pdf", "wb") as result_pdf:
	    pdfs.write(result_pdf)
    
def cleanup():
    files = glob.glob('odtemp/*')
    for f in files:
        os.remove(f)
    os.rmdir('odtemp')
    status.set("Done")
    delay = 5
    window.after(1000 * delay, window.destroy)



def combine():
    paths = Path('.').glob("odtemp/*.docx")
    global filenames
    filenames = []
    for path in paths:
        filenames.append(path)
    pdfs = docxs_to_pdf(filenames)
    joinpdf(pdfs)
    cleanup()

# Add a Label widget
label = tk.Label(window, text="Click the Button to browse the Files", font=('Georgia 13'))
label.pack(pady=10)

# Create a Button
ttk.Button(window, text="Browse Data", command=open_data).pack(pady=20)
labelinfo1 = tk.Label(window, textvariable=data, font=('Georgia 13'))
labelinfo1.pack(pady=10)

label2 = tk.Label(window, text="Click the Button to browse the Files", font=('Georgia 13'))
label2.pack(pady=10)
label3 = tk.Label(window, text="Be sure to add ['name','regd','dept','dates','hours'] as MailMerge Template in the Word Document", font=('Georgia 8')).pack(pady=5)

ttk.Button(window, text="Browse Template", command=open_template).pack(pady=20)

labelinfo2 = tk.Label(window, textvariable=template, font=('Georgia 13')
)
labelinfo2.pack(pady=10)


g = ttk.Button(window, text="Generate OD", command=generate).pack(pady=20)
labelinfo3  = tk.Label(window, textvariable=status).pack(pady=10)
window.mainloop()
