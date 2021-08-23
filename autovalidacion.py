from tkinter import *
from tkinter import filedialog
import os
import autoval

root = Tk()
root.title('SET autoval 0.1.0')


def subir_archivo():
    global filepath
    filepath = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"),("All files", "*.*")))
    filename = os.path.basename(filepath)
    up = Label(root, text='Archivo subido: '+ filename)
    up.grid(row=1, column=1)

def ejecutar_prog():
    autoval.autoval_func(filepath)
    
myLabel = Label(root, text='Validacion de Facturas SET').grid(row=0, column=1)
space1 = Label(root,text='').grid(row=4, column=0)
space2 = Label(root, text='').grid(row=5, column=0)

upload_butn = Button(root, text='Subir archivo Excel', command=subir_archivo).grid(row=6, column=0)

execute_butn = Button(root, text='Ejecutar programa', command=ejecutar_prog).grid(row=6, column=2)

root.mainloop()