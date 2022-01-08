import parseExcel
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from bs4 import BeautifulSoup
import html2text

# This program should be refactored with OO later, but for now
# We'll use global variables to communicate between functions.
# normally these would be class variables.

# program variables, declaring them here, to make it obviuos they are global.
root = None
c = None
listView = None
studentInfo = None
codeDump = None
filename = ""
excelWorkbook = None
shadowList = None


def updateUI():
    global shadowList
    global excelWorkbook

    shadowList = []
    listView.delete(first=0, last=listView.size())

    for row in excelWorkbook:
        info_str = str(row["Asker User ID"]) + "," + str(row["Asker Email ID"]) + "," + str(
            row["Asker First Name"]) + " " + str(row["Asker Last Name"])
        shadowList.append(info_str)
        listView.insert(END, info_str)


def populateStudentInfo(infoText):
    global studentInfo
    global root
    if studentInfo is not None:
        studentInfo.destroy()
    studentInfo = createTextView(c)
    studentInfo.grid(column=1, row=0, rowspan=1, sticky=(N, W, E), padx=2, pady=2)
    studentInfo.insert(INSERT, infoText)

def selectedItem(event):
    idx = int(listView.curselection()[0])
    row = excelWorkbook[idx]
    cols = ["Question ID", "Question Date", "Answer Date", "Asker User ID", "Asker First Name", "Asker Last Name",
            "Asker Email ID",
            "Asker IP Address", "Asker School Name", ]
    output_str = ""
    for item in cols:
        output_str += item + ":" + str(row[item]) + "\n"

    populateStudentInfo(output_str)

    text = html2text.html2text(str(row["Answer"]))

    codeDump.delete(index1=CURRENT, index2=END)
    codeDump.insert(END, text)
    studentInfo['state'] = 'disabled'


def load_file():
    global filename
    global excelWorkbook
    succeeded = True
    filename = filedialog.askopenfilename()
    try:
        excelWorkbook = parseExcel.parseChegg(filename)

    except:
        messagebox.showwarning(title="File Error", message="Not a valid Chegg Excel file")
        succeeded = False
    if succeeded:
        updateUI()
    return succeeded


def createListView(parent, defaultChoices=[]):
    studentList = Listbox(parent,width=75, listvariable=defaultChoices)
    return studentList


def createTextView(parent, text="", height=20):
    textView = Text(parent, height=height)
    return textView


def createMenu(parent, handler=load_file):
    menu = Menu(parent)

    # menu Items
    file_menu = Menu(menu)
    menu.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Open", command=load_file)
    file_menu.add_command(label="Exit", command=root.quit)
    return menu


root = Tk()
root.title('CHEGG View')
menu = createMenu(root)
root.configure(menu=menu)

c = ttk.Frame(root, padding=(5, 5, 12, 0))
c.grid(column=0, row=0, rowspan=2, sticky=(N, W, E, S))

# Create and place listBox
listView = Listbox(c)
listView.grid(column=0, row=0, rowspan=1, columnspan=1, sticky=(N, E, W, S), padx=2, pady=2)
listView.bind('<<ListboxSelect>>', selectedItem)

# Create Textbox
populateStudentInfo("Welcome to CheggView alpha")
# Add List View

codeDump = createTextView(c)
codeDump.grid(column=0, row=1, columnspan=2, sticky=(N, W, E, S), padx=2, pady=2)

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
# configure resizing behavior
c.columnconfigure(0, weight=0, minsize = 200)
c.columnconfigure(1, weight=1)
c.rowconfigure(0, weight=0)
c.rowconfigure(1, weight=1)

root.mainloop()
