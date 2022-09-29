from turtle import bgcolor
import win32ui
import win32print
import win32con
import time
import threading
from tkinter import messagebox
from pygame import mixer

INCH = 1440

hDC = win32ui.CreateDC ()

from tkinter import *
from tkinter import ttk


window = Tk()
window.title("Hitung Kertasku")
window.geometry('400x160')
window.configure(background = "white");

valPrinter  = StringVar()
valJumlah   = StringVar()
valDelay    = StringVar()
valOption   = StringVar()
valOption   = StringVar()
valchcbox   = IntVar()
hitung_text =StringVar()
hitung_text.set("0 Detik")

all_printers1 = [printer[2] for printer in win32print.EnumPrinters(5)]
all_printers2 = [printer[2] for printer in win32print.EnumPrinters(2)]
all_printers  = all_printers1+all_printers2
valPrinter.set(win32print.GetDefaultPrinter())
valJumlah.set(1)
valDelay.set(10)

printer = Label(window ,background="white",text = "Printer").grid(row = 0,column = 0)
jumlah  = Label(window ,background="white",text = "Jumlah").grid(row = 1,column = 0)
delay   = Label(window ,background="white",text = "delay (Detik)").grid(row = 2,column = 0)
hitung   = Label(window ,background="white",textvariable = hitung_text).grid(row = 5,column = 1)
# delay   = Label(window ,background="white",text = "Berulang").grid(row = 3,column = 0)


# inPrinter = Entry(window,textvariable = valPrinter,width="50").grid(row = 0,column = 1)
w = OptionMenu(window,valPrinter, *all_printers)
w.config(width=40)
w.grid(row = 0,column = 1)

inJumlah  = Entry(window,textvariable = valJumlah,width="50").grid(row = 1,column = 1)
inDelay   = Entry(window,textvariable = valDelay,width="50").grid(row = 2,column = 1)
cprint    = Checkbutton(window, text='Print Berulang ',variable=valchcbox, onvalue=1, offvalue=0)
cprint.config(width=40)
cprint.grid(row = 3,column = 1)


def clicked():
    removebtn()
    t1=threading.Thread(target=mulai)
    t1.start()
    global mulaiPrint
    mulaiPrint = 1

def stopclicked():
    t2=threading.Thread(target=berhenti)
    t2.start()

mixer.init()
def mulai():
    mixer.music.load('play.mp3')
    mixer.music.play()
    i = 0   
    hDC.CreatePrinterDC (win32print.GetDefaultPrinter ())
    while i < int(valJumlah.get()):
        if mulaiPrint == 1:
            hDC.StartDoc ("tes")
            hDC.StartPage ()
            hDC.SetMapMode (win32con.MM_TWIPS)
            hDC.DrawText ("", (0, INCH * -1, INCH * 8, INCH * -2), win32con.DT_CENTER)
            hDC.EndPage ()
            hDC.EndDoc ()
            i += 1

    if mulaiPrint == 1:
        mixer.music.load('stop.mp3')
        mixer.music.play()

        if valchcbox.get() == True :
            mulaihitung(int(valDelay.get()))
            # time.sleep(int(valDelay.get()))
            #  label.configure(text="This is updated Label text")
        else:
            berhenti()

def mulaihitung(i):

    x = reversed(range(i))
    for n in x:
        if mulaiPrint == 1:
            time.sleep(1)
            text = str(n)+" Detik";
            hitung_text.set(text)
            if n == 0:
                clicked()
        else:
            text = "0 Detik";
            hitung_text.set(text)



def berhenti():
    global mulaiPrint
    mulaiPrint = 0
    btnStop.grid_remove()
    btnStart.grid(row=5,column=0)


btnStart = Button(window ,text="Print",bg='green',fg='white',command=clicked)
btnStop = Button(window ,text="Cancel",command=stopclicked)


def removebtn():
    btnStart.grid_remove()
    btnStop.grid(row=5,column=0)

btnStart.grid(row=5,column=0)
window.mainloop()






















