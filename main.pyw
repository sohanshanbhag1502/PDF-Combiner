from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from PyPDF2 import PdfFileMerger
from tkinter import messagebox as ms
from subprocess import Popen

import tkinter as tk;

class ReorderableListbox(tk.Listbox):
    """ A Tkinter listbox with drag & drop reordering of lines """
    def __init__(self, master, li, **kw):
        kw['selectmode'] = tk.EXTENDED
        tk.Listbox.__init__(self, master, kw)
        self.bind('<Button-1>', self.setCurrent)
        self.bind('<Control-1>', self.toggleSelection)
        self.bind('<B1-Motion>', self.shiftSelection)
        self.bind('<Leave>',  self.onLeave)
        self.bind('<Enter>',  self.onEnter)
        self.selectionClicked = False
        self.left = False
        self.unlockShifting()
        self.ctrlClicked = False
        self.list=li
    def orderChangedEventHandler(self):
        pass

    def onLeave(self, event):
        # prevents changing selection when dragging
        # already selected items beyond the edge of the listbox
        if self.selectionClicked:
            self.left = True
            return 'break'
    def onEnter(self, event):
        #TODO
        self.left = False

    def setCurrent(self, event):
        self.ctrlClicked = False
        i = self.nearest(event.y)
        self.selectionClicked = self.selection_includes(i)
        if (self.selectionClicked):
            return 'break'

    def toggleSelection(self, event):
        self.ctrlClicked = True

    def moveElement(self, source, target):
        global dir_
        if not self.ctrlClicked:
            element = self.get(source)
            self.delete(source)
            self.insert(target, element)
        self.list=self.get(0, END)
        dir_=self.list

    def unlockShifting(self):
        self.shifting = False
    def lockShifting(self):
        # prevent moving processes from disturbing each other
        # and prevent scrolling too fast
        # when dragged to the top/bottom of visible area
        self.shifting = True

    def shiftSelection(self, event):
        if self.ctrlClicked:
            return
        selection = self.curselection()
        if not self.selectionClicked or len(selection) == 0:
            return

        selectionRange = range(min(selection), max(selection))
        currentIndex = self.nearest(event.y)

        if self.shifting:
            return 'break'

        lineHeight = 15
        bottomY = self.winfo_height()
        if event.y >= bottomY - lineHeight:
            self.lockShifting()
            self.see(self.nearest(bottomY - lineHeight) + 1)
            self.master.after(500, self.unlockShifting)
        if event.y <= lineHeight:
            self.lockShifting()
            self.see(self.nearest(lineHeight) - 1)
            self.master.after(500, self.unlockShifting)

        if currentIndex < min(selection):
            self.lockShifting()
            notInSelectionIndex = 0
            for i in selectionRange[::-1]:
                if not self.selection_includes(i):
                    self.moveElement(i, max(selection)-notInSelectionIndex)
                    notInSelectionIndex += 1
            currentIndex = min(selection)-1
            self.moveElement(currentIndex, currentIndex + len(selection))
            self.orderChangedEventHandler()
        elif currentIndex > max(selection):
            self.lockShifting()
            notInSelectionIndex = 0
            for i in selectionRange:
                if not self.selection_includes(i):
                    self.moveElement(i, min(selection)+notInSelectionIndex)
                    notInSelectionIndex += 1
            currentIndex = max(selection)+1
            self.moveElement(currentIndex, currentIndex - len(selection))
            self.orderChangedEventHandler()
        self.unlockShifting()
        return 'break'

dir_=[]

root=Tk()
root.geometry('500x400')
root.title('PDF Combiner')
root.resizable(0,0)

def openfiles():
    global dir_
    dir__=filedialog.askopenfilenames(defaultextension='.pdf',filetypes=(('Portable document format(PDF)','*.pdf'),
    ('All files','*.*')),initialdir='/Documents',title='Open PDF File')
    if dir__:
        dir_+=list(dir__)
        lb.delete(0, END)
        lb.insert(END,*dir_)

def combine(pdflist):
    global dir_
    pdfmerger=PdfFileMerger()
    for j in pdflist:
        try:
            pdfmerger.append(j)
        except :
            ms.showerror('Error','PDF file merging operation unsuccessful')
            pdfmerger.close()
            return
    save_dir=filedialog.asksaveasfilename(title='Save merged PDF file as',confirmoverwrite=True,defaultextension='.pdf',
    filetypes=(('Portable document format(PDF)','*.pdf'),),initialdir='/Documents',initialfile='mergedpdf')
    pdfmerger.write(save_dir)
    pdfmerger.close()
    Popen(['powershell', '-noprofile', f"&'{save_dir}'".replace('/','\\')], shell=True)
    root.bell()
    root.update_idletasks()
    return

def delete_select(_=None):
    if lb.selection_get():
        global dir_
        sel=lb.selection_get()
        dir_.remove(sel)
        lb.delete(ACTIVE)

def delete():
    if lb.get(0, END):
        global dir_
        dir_.clear()
        lb.delete(0, END)

Label(text='PDF Combiner',font='Arial 20 bold').pack()
Label(text='Version 1.0.3',font='Arial 15 italic').pack(padx=30)

frame=Frame(width=300,height=200)
frame.place(x=10,y=100)

xscroll=ttk.Scrollbar(frame,orient='horizontal')
yscroll=ttk.Scrollbar(frame,orient='vertical')
lb=ReorderableListbox(frame,xscrollcommand=xscroll.set,yscrollcommand=yscroll.set,height=13,width=75, selectmode=EXTENDED,
                        li=dir_)
lb.bind('<Delete>',delete_select)
xscroll.config(command=lb.xview)
yscroll.config(command=lb.yview)
yscroll.pack(side='right',fill=Y)
xscroll.pack(side='bottom',fill=X)
lb.pack(fill='both',expand=1)

b1=Button(text='Open files',relief='ridge',font='Arial 12',bg='white',activebackground='#d8a469',cursor='hand2',overrelief='ridge',command=openfiles)
b1.bind('<Motion>',lambda _:b1.config(bg='#d8a469'))
b1.bind('<Leave>',lambda _:b1.config(bg='white'))
b1.place(x=10,y=350)

b2=Button(text='Combine',relief='ridge',font='Arial 12',bg='white',activebackground='#d8a469',cursor='hand2',overrelief='ridge',command=lambda :combine(list(dir_)))
b2.bind('<Motion>',lambda _:b2.config(bg='#d8a469'))
b2.bind('<Leave>',lambda _:b2.config(bg='white'))
b2.place(x=400,y=350)

b3=Button(text='Clear Queue', relief='ridge', font='Arial 12', bg='white', activebackground='#d8a469', cursor='hand2', overrelief='ridge', command=delete)
b3.bind('<Motion>',lambda _:b3.config(bg='#d8a469'))
b3.bind('<Leave>',lambda _:b3.config(bg='white'))
b3.place(x=188,y=350)

root.mainloop()