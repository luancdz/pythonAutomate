import tkinter 
from tkinter import *
from tkinter import messagebox

from numpy.lib.function_base import place
import fileController

#button para o upload
#def datafetch():
#    x = lbx.get(ACTIVE)
#    messagebox.showinfo("data", "You Selected" + x)

def createview():

    #button para o upload
    def uploadFile():
        #x = lbx.get(ACTIVE)
        #messagebox.showinfo("data", "You Selected" + x)
        path = fileController.getPath()
        print(fileController.getSO(path))
        SOList = fileController.getSO(path)
        lbx.delete(0, END)

        idx = 0
        for SONumber in SOList:
            lbx.insert(idx, SONumber)
            idx += 1
        

    root = tkinter.Tk()
    root.geometry("800x400")
    root.title("NF SAP")

    fr = Frame(root)
    fr.pack(side=LEFT)

    lbl1 = Label(fr, text="", width=6)
    lbl1.pack(side=LEFT)

    lbl2 = Label(fr, text="Sales Order Document", font=("Verdana",16))
    lbl2.pack(side=TOP)

    sbr = Scrollbar(
        fr,
    )

    sbr.pack(side=RIGHT,fill="y")

    lbx = Listbox(
        fr,
        font = ("Verdana", 16)

    )

    #lbx.place(x=20, y=25)

    lbx.pack(expand=True, side=LEFT, fill=BOTH)

    #for data in range(50):
    #    lbx.insert(data, "Sample Data" + str(data+1))

    sbr.config(command=lbx.yview)
    lbx.config(yscrollcommand=sbr.set)

    btn = Button(
        root,
        text = "Upload file",
        font=("Verdana",16),
        command = uploadFile,
        
    )

    btn.place(x=175, y=350)

    root.mainloop()

