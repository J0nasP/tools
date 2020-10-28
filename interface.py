import tkinter as tk
from tkinter import *
import re, windnd
from tkinter import filedialog


root = Tk()
root.title("PDF manipulator")


filename = StringVar()
pages = StringVar()
start = StringVar()
end = StringVar()



Label(root, text="PDF Manipulator", font=("Berlin Sans FB Demi", 16, "bold")).\
                                    grid(row=0, column=0, columnspan=4, sticky=(E, N, S))

Button(root, text= "Add a File", command=lambda: add_file()).grid(row=2, column=0)
Button(root, text= "Remove File", command=lambda: remove_file()).grid(row=4, column=0)
Button(root, text= "Merger").grid(row=6, column=1, rowspan=2)
Button(root, text= "Quit", command=lambda: root.quit()).grid(row=6, column=2, rowspan=2)
Button(root, text= "splitter").grid(row=6, column=3, rowspan=2)

Label(root, text="File: ").grid(row= 1, column= 3)
Label(root, textvariable=filename, width=20).grid(row=1, column= 4, sticky= (N, S, E, W))

Label(root, text= "Pages: ").grid(row=2, column= 3)
Label(root, textvariable=pages).grid(row=2, column= 4)

Label(root, text="Start: ").grid(row=4, column=3)
start_entry= Entry(root, textvariable=start, width=3)
start_entry.grid(row=4, column=4)

Label(root, text="End: ").grid(row=5, column= 3)
end_entry= Entry(root, textvariable=end, width=3)
end_entry.grid(row=5, column= 4)

lb = Listbox(root, selectmode= EXTENDED)
lb.grid(row=1, column=2, rowspan=5, sticky= (N, S, E, W))

lb.insert(END)
dataLocations = []
data = []
lb.delete(0, END) 


def insert():
    lb.delete(0, len(data))
    for ind, _ in enumerate(data):
        lb.insert(END, data[ind])

def findAmount(od, d="", c=1):
    d = od+" [{}]".format(str(c))
    return d if not d in data else findAmount(od, d, c+1)

def selector(ind):
    data[ind] = data[ind].split(" <")[0]
    data[current] += " <"
    insert()

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
    if not file: d = [filedialog.askopenfilename(filetypes=(('PDF File','*.pdf'),('Word files','.docx')))]
    else: d = file
    for i in d:
        print(i)
        if i in data:
            i = findAmount(d)
        data += [i.split("/")[-1].split("\\")[-1]]
        dataLocations += [i]
    insert()

def remove_file():
    index = int(lb.curselection()[0])
    data.pop(index)
    lb.delete(ANCHOR)

windnd.hook_dropfiles(root, add_file, force_unicode=True)


for child in root.winfo_children():
    child.grid(padx=10, pady=10)

root.mainloop()