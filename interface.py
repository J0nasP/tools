import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog, Label, Button, Listbox, Tk, StringVar
import re, windnd
import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter
import threading
import time, glob
import os, re
from docx import Document
import win32com.client as client
from pdf2docx import parse
from docx2pdf import convert



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
Button(root, text= "Split PDF Files", command=lambda: pdf_selector()).grid(row=8, column=1)
Button(root, text= "PDF to Word", command=lambda: pdf_to_word()).grid(row=7, column=0)


Label(root, text="File: ").grid(row= 2, column= 3)
Label(root, textvariable=filename, width=25).grid(row=2, column= 4, columnspan=1)

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
            filename.set(data[current][:-1])
            pages.set(PdfFileReader(dataLocations[current]).getNumPages())
        time.sleep(0.1)

        
def pdf_merger():
    """Merges PDF files and saves them as one file """
    if len(data) < 1:
        button = False
        messagebox.showerror('An error occured','No file where selected \n Please add the files you want to merge')
    elif len(data) == 1:
        button = False
        messagebox.showerror('An error occured', 'Only one file selected \n Please add the files you want to merge')
    else:
        button = True
    
    while button == True:
        userfilename = simpledialog.askstring('Name the new file', 'What do you want to call the new file?')
        if len(userfilename) < 1:
            messagebox.showerror('An error has oucurred','You forgot to name the file')
            pdf_merger()
         
        try:
            outFolder = filedialog.askdirectory(title='Where do you want to save the file', initialdir='/')  
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
                                            '\n' '             Do yout wish to quit?   ')

        if wish_quit_mes == True or wish_quit_mes == None:
            root.quit()
        else: 
            pass


        lb.delete(0, "end")
        data.clear()
        dataLocations.clear()

def want_to_quit():
    """ Quit button function that makes sure if they want to close the app or continue """
    want_quit = messagebox.askyesno(title='Are you sure?', message='Are you sure you want to quit??')
    if want_quit == True:
        root.quit()
    else:
        pass

def pdf_selector():
    """Function where the user selects if they want one pdf file or many pdf files  """
    if len(data)  == 0:
        button = False
        messagebox.showerror('An error occured',' No file where selected \n Please add the files you want to split')
    else:
        button = True
    
    if len(data) > 1:
        button = False
        messagebox.showerror('An error occured!', 'Sorry you can only split one file at the time \n       We are working on it!:)')
    if len(data) == 1:
        choise = simpledialog.askstring(title="Do you want one or many files?", 
        prompt="Do you want to save the splitted file in one file or one file for each page? \n" 
                                                "                                     Enter either ONE or MANY")

        if choise.lower() == "ONE".lower():
            pdf_splitter_one()
        elif choise.lower() == "MANY".lower():
            pdf_splitter_many()
        else:
            messagebox.showerror('An error occured!', 'Sorry wrong input! Please type MANY or ONE in the text field!')


def pdf_splitter_many():
    """Split a Pdf file and saves it as each page as a file"""

    userfilename = simpledialog.askstring('Name the new file', 'What do you want to call the new file?')
    if len(userfilename) < 1:
        messagebox.showerror('An error has oucurred','You forgot to name the file')
        pdf_splitter_many()
      
    try:
        outFolder = filedialog.askdirectory(title='Where do you want to save the file?', initialdir='/')  
    except FileNotFoundError: 
        messagebox.showerror('An error has ocurred!', "The Choosen folder dosn't exsist")
               
    pdf2split = []

    for filename in dataLocations:
        if filename.endswith('.pdf'):
            pdf2split.append(filename)
        
    pageCounter=0

    for Files in pdf2split:
        pages = PdfFileReader(filename).getNumPages()
        pageCount = pages + pageCounter

    numInput = False

    while numInput == False:
        page_range = simpledialog.askstring('which pages do you want?',
                                            'which pages do you want? ex. 1-3 or 1,5,7 ({} total pages)'.format(pageCount))
            
        inp = page_range
        inps = (re.findall("[0-9]-[0-9]|[0-9]", inp))
        output = [list(range(int(i[0]), int(i[1]) + 1)) for i in [i.split("-") if "-" in i else [i, i] for i in inps]]
        rangeList = [x for i in output for x in i]
        page_ranges = (x.split("-") for x in page_range.split(","))
        range_list = [i for r in page_ranges for i in range(int(r[0]), int(r[-1]) + 1)]

        print(range_list)
        if max(range_list) >  pageCount:
            messagebox.showerror('An error occured!', 'The page number is to high!')
            numInput = False

        if pageCount < 2:
            messagebox.showerror('An error ocurred!', "There isn't enough pages to split")
            numInput = False
            
        if max(range_list) <= pageCount:
            numInput = True


        
    outFolder = ''.join(outFolder)

    for filename in pdf2split:
        inputPdf = PdfFileReader(filename)
        try:
            for page in range_list:
                outputWriter = PdfFileWriter()
                outputWriter.addPage(inputPdf.getPage(page - 1))
                outFile = os.path.join(outFolder, '{} page {}.pdf'.format(userfilename, page))
                with open(outFile, 'wb') as out:
                    outputWriter.write(out)
        except IndexError:
            messagebox.showerror('An error has occured', 'Range has exceeded number of pages in the input. \nFile will still be saved')
                    
    fileCount = str(len(range_list))
    wish_quit_mes = messagebox.askyesno(title='Do you wish to quit?',
                                            message='Jobs done! \n ' +  fileCount + ' files are saved at ' + outFolder +
                                            '\n' '            Do you wish to quit?    ')

    if wish_quit_mes == True or wish_quit_mes == None:
        root.quit()
    else:
        pass

    lb.delete(0, "end")
    data.clear()
    dataLocations.clear()
    button = False
    return

def pdf_splitter_one():
    """Split a Pdf file and saves it as one file containing the selected files"""
    userfilename = simpledialog.askstring('Name the new file', 'What do you want to call the new file?')
    if len(userfilename) < 1:
        messagebox.showerror('An error has oucurred','You forgot to name the file')
        pdf_splitter_many()
      
    try:
        outFolder = filedialog.askdirectory(title='Where do you want to save the file?', initialdir='/')  
    except FileNotFoundError: 
        messagebox.showerror('An error has ocurred!', "The Choosen folder dosn't exsist")
               
    pdf2split = []

    for filename in dataLocations:
        if filename.endswith('.pdf'):
            pdf2split.append(filename)
        
    pageCounter=0

    for Files in pdf2split:
        pages = PdfFileReader(filename).getNumPages()
        pageCount = pages + pageCounter

    numInput = False

    while numInput == False:
        page_range = simpledialog.askstring('which pages do you want?',
                                            'which pages do you want? ex. 1-3 or 1,5,7 ({} total pages)'.format(pageCount))
            
        inp = page_range
        inps = (re.findall("[0-9]-[0-9]|[0-9]", inp))
        output = [list(range(int(i[0]), int(i[1]) + 1)) for i in [i.split("-") if "-" in i else [i, i] for i in inps]]
        rangeList = [x for i in output for x in i]
        page_ranges = (x.split("-") for x in page_range.split(","))
        range_list = [i for r in page_ranges for i in range(int(r[0]), int(r[-1]) + 1)]

        print(range_list)
        if max(range_list) >  pageCount:
            messagebox.showerror('An error occured!', 'The page number is to high!')
            numInput = False

        if pageCount < 2:
            messagebox.showerror('An error ocurred!', "There isn't enough pages to split")
            numInput = False
            
        if max(range_list) < pageCount:
            numInput = True

        os.chdir(outFolder)

        output_writer = PdfFileWriter()
        input_pdf = PdfFileReader(open(filename, 'rb'))
        output_file = open(userfilename + '.pdf', "wb")
        inp = page_range
        inps = (re.findall("[0-9]-[0-9]|[0-9]", inp))
        output = [i for x in [list(range(int(i[0]), int(i[1]) + 1)) for i in [i.split("-") if "-" in i else [i, i] for i in inps]] for i in x]


        for p in output:

            try:

                output_writer.addPage(input_pdf.getPage(p - 1))
            except IndexError:
                # Alert the user and stop adding pages
                messagebox.showerror(title="Info",
                                        message="Range exceeded number of pages in input.\nFile will still be saved.")

        output_writer.write(output_file)
        output_file.close()
        numInput= True  

    wish_quit_mes = messagebox.askyesno(title='Do you wish to quit?',
                                            message='Jobs done! \n ' +  userfilename + ' files are saved at ' + outFolder +
                                            '\n' '            Do you wish to quit?    ')

    if wish_quit_mes == True or wish_quit_mes == None:
        root.quit()
    else:
        pass

    lb.delete(0, "end")
    data.clear()
    dataLocations.clear()
    button = False
    return
    
def pdf_to_word():
    if len(data) < 1:
        button = False
        messagebox.showerror('An error occured','No file where selected \n  Please add the files you want to convert')    
    else:
        button = True

        while button == True:
            word = client.gencache.EnsureDispatch("Word.Application")


            pdfs_path = dataLocations
            reqs_path = filedialog.askdirectory(title='Where do you want to save the files?',
                                    initialdir='\\')

            pdf2convert = []

            for filename in dataLocations:
                if filename.endswith('.pdf'):
                    pdf2convert.append(filename)
            

            pdfs_path += '\\'
            reqs_path += '\\'
            please = reqs_path.replace('/','\\', -1)
            print(please)

            for doc in pdf2convert:

                filename = doc.split('\\')[-1]
                wb = word.Documents.Open(filename)
                out_file = os.path.abspath(filename[0:-4] + ".docx".format(doc))
                outfile = out_file.replace('/','\\')
                print(outfile)
                wb.SaveAs(outfile, FileFormat=16) # file format for docx
                wb.Close()
                        
            word.Quit()
            button = False
        wish_quit_mes = messagebox.askyesno(title='Do you wish to quit?',
                                            message='Jobs done! \n ' +  userfilename + ' files are saved at ' + outFolder +
                                            '\n' '            Do you wish to quit?    ')

    if wish_quit_mes == True or wish_quit_mes == None:
        root.quit()
    else:
        pass

    lb.delete(0, "end")
    data.clear()
    dataLocations.clear()
    button = False
    return
    
def word_to_pdf():
    if len(data) < 1:
        button = False
        messagebox.showerror('An error occured','No file where selected \n  Please add the files you want to convert')    
    else:
        button = True

        while button == True:


            input_folder = dataLocations

            pdfs_path = dataLocations
            reqs_path = filedialog.askdirectory(title='Where do you want to save the files?',
                                    initialdir='/')
            
            word2pdf = []

            for files in dataLocations:
                if files.endswith('docx'):
                    word2pdf.append(files)
            
            for filenames in word2pdf:
                convert(filenames,reqs_path)

x = threading.Thread(target=updatePages, args=(pages, filename))
x.setDaemon(True)
x.start()


root.resizable(width=False, height=False)

root.mainloop()


