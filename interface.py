import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog, Label, Button, Listbox, Tk, StringVar
import re, windnd
import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter
import threading
import time
import os



root = Tk()
root.title("PDF manipulator")
root.geometry('650x400')


filename = StringVar()
pages = StringVar()



Label(root, text="PDF Manipulator", font=("Berlin Sans FB Demi", 16, "bold")).\
                                    grid(row=0, column=0, columnspan=7, sticky=('N'))

Button(root, text= "Add a File", command=lambda: add_file()).grid(row=2, column=0)
Button(root, text= "Remove File", command=lambda: remove_file()).grid(row=3, column=0)
Button(root, text= "Merge PDF Files", command=lambda:pdf_merger()).grid(row=8, column=0)
Button(root, text= "Quit", command=lambda: want_to_quit()).grid(row=8, column=2)
Button(root, text= "Split PDF Files").grid(row=8, column=1)


Label(root, text="File: ").grid(row= 2, column= 3)
Label(root, textvariable=filename, width=25).grid(row=2, column= 4, columnspan=2)

Label(root, text= "Pages: ").grid(row=3, column= 3)
Label(root, textvariable=pages).grid(row=3, column= 4)

lb = Listbox(root, selectmode= "extended", height=15, width=35)
lb.grid(row=1, column=2, rowspan=7, sticky= ('N', 'S', 'E', 'W'))



lb.insert("end")
dataLocations = []
data = []
lb.delete(0, "end") 

def insert():
    lb.delete(0, len(data))
    for ind, _ in enumerate(data):
        lb.insert("end", data[ind])

def findAmount(od, d="", c=1):
    d = od+" [{}]".format(str(c))
    return d if not d in data else findAmount(od, d, c+1)

def selector(ind):
    data[ind] = data[ind].split(" <")[0]
    data[current] += " <"

insert()
current = 0

def select(event):
    global current
    last = current
    current = lb.nearest(event.y)
    selector(last)
    print("Current changed to", current)

def move(event):
    global current
    element = lb.nearest(event.y)
    try:
        if element != current:
            From = data[current]
            To = data[element]
            FromL = dataLocations[current]
            ToL = dataLocations[element]
            print("{} -> {}".format(From[:-2], To))
            data[current] = To
            data[element] = From
            print("{} -> {}".format(FromL, ToL))
            dataLocations[current] = ToL
            dataLocations[element] = FromL
            print(dataLocations)
            insert()
            current = element
    except Exception as e: print(e)

lb.bind('<B1-Motion>', move)
lb.bind('<Button-1>', select)

def add_file(file = None):
    global data
    global dataLocations
    if not file: d = [filedialog.askopenfilename(filetypes=(('PDF','*.pdf'),('Word files','.docx')))]
    else: d = file
    hasErrored = False
    for i in d:
        splitdata = i.split("/")[-1].split("\\")[-1]
        if splitdata.endswith(".pdf" or "PDF"):
            print(i)
            if i in data:
                i = findAmount(d)
            data += [splitdata]
            dataLocations += [i]
        elif not hasErrored:
            messagebox.showerror("You can only use pdf files", "Wrong file type! you can only use PDF files")
            hasErrored = True
    insert()

def remove_file():
    index = int(lb.curselection()[0])
    data.pop(index)
    dataLocations.pop(index)
    lb.delete("anchor")

windnd.hook_dropfiles(root, add_file, force_unicode=True)

for child in root.winfo_children():
    child.grid(padx=10, pady=10)

def updatePages(pages: StringVar, filename: StringVar):

    while True:
        if data:
            filename.set(data[current][:-2])
            pages.set(PdfFileReader(dataLocations[current]).getNumPages())
        time.sleep(0.1)

        
def pdf_merger():

    if len(data) < 1:
        button = False
        messagebox.showerror('An error occured','No file where selected \n Please add the files you want to merge')
    elif len(data) is 1:
        button = False
        messagebox.showerror('An error occured', 'Only one file selected \n Please add the files you want to merge')
    else:
        button = True
    
    while button is True:
        userfilename = simpledialog.askstring('Name the new file', 'What do you want to call the new file?')
        if len(userfilename) < 1:
            messagebox.showerror('An error has oucurred','You forgot to name the file')
            pdf_merger()
            outFolder = filedialog.askdirectory(title='Where do you want to save the file', initialdir='/')  
        try:
            os.chdir(outFolder)
        except FileNotFoundError: 
            messagebox.showerror('An error has ocurred!', "The Choosen folder dosn't exsist")
            continue
        else:
            break

        pdf2merge = []

        for filename in dataLocations:
            if filename.endswith('.pdf'):
                pdf2merge.append(filename)

        pdf_writer = PdfFileWriter()

        for filename in pdf2merge:
            pdf_file_obj = open(filename, 'rb')
            pdf_reader = PdfFileReader(pdf_file_obj)
            for page_num in range(pdf_reader.numPages):
                page_obj = pdf_reader.getPage(page_num)
                pdf_writer.addPage(page_obj)
        pdf_output = open(userfilename + '.pdf', 'wb')
        pdf_writer.write(pdf_output)
        pdf_output.close()

        wish_quit_mes = messagebox.askyesno(title='Do you wish to quit?',
                                            message='Jobs done! \n ' + userfilename  + 'is saved at ' + outFolder +
                                            '\n' '         Do yout wish to continue    ')

        lb.delete(0, "end")
        data.clear()
        dataLocations.clear()

class NoInputError(Exception): pass

def want_to_quit():
    want_quit = messagebox.askyesno(title='Are you sure?', message='Are you sure you want to quit??')
    if want_quit is True:
        root.quit()
    else:
        pass





x = threading.Thread(target=updatePages, args=(pages, filename))
x.setDaemon(True)
x.start()


root.resizable(width=False, height=False)

root.mainloop()


