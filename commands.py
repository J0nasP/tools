from tkinter import *
from tkinter import filedialog, dnd
from tkinter.dnd import DndHandler
import re, windnd

#root = Tk()
master = Frame()
master.grid(row=0,column=0,columnspan=2)

lb = Listbox(master, selectmode=EXTENDED)
lb.pack()



lb.insert(END)
dataLocations = []
data = []
lb.delete(0, END) # clear


def insert():
    lb.delete(0, len(data))
    for ind, _ in enumerate(data):
        lb.insert(END, data[ind])

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
    if not file: d = [filedialog.askopenfilename()]
    else: d = file
    for i in d:
        print(i)
        if i in data:
            i = findAmount(d)
        data += [i.split("/")[-1].split("\\")[-1]]
        dataLocations += [i]
    insert()

windnd.hook_dropfiles(master, add_file, force_unicode=True)

#Button(root, text="Add", command=lambda: add_file()).grid(row=1,column=0)
#Button(root, text="Open", command=lambda: [print(i) for i in dataLocations]).grid(row=1, column=1)



#root.mainloop()