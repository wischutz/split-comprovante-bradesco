from PyPDF2 import PdfFileReader, PdfFileWriter
from pdfminer import high_level
import random
import os
import subprocess
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

def getFolderPath():
    folder_selected = filedialog.askdirectory()
    folderpath.set(folder_selected)

def openFile():
    filepath = filedialog.askopenfilename(#initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                          title="Selecionar Arquivo de Comprovantes Bradesco",
                                          filetypes= (("Arquivos pdf","*.pdf"),))
    
    infile =  open(filepath, 'rb')
    reader = PdfFileReader(infile)
    num_pages = reader.getNumPages()
    for x in range(num_pages):
        writer = PdfFileWriter()
        page = reader.getPage(x)
        text = high_level.extract_text(filepath, "", [x])
        lines = text.split("\n\n")

        ammount = ''
        description = ''

        for line in lines:
            if "Valor total:" in line:
                ammount = line.split('Valor total: R$ ')[1].strip()
            elif "Descrição: " in line:
                description = line.split('Descrição: ')[1].strip()
                
        filename = f'{folderpath.get()}/{description.lower().replace(" ", "_")}_{ammount.replace(".","").replace(",","")}_{random.getrandbits(32)}.pdf'
        path = os.path.join(folderpath.get(), filename)
        writer.addPage(page)    
        outfile = open(path, 'wb')
        writer.write(outfile)
    if os.name == 'nt':
        subprocess.Popen(f'explorer "{folderpath.get()}"')

    messagebox.showinfo(title="Sucesso!", message="Arquivo separado com sucesso!")

window = Tk()
window.geometry("650x400")
window.title("Separar Comprovantes Bradesco")

folderpath = StringVar()
folderpath.set("Escolha o pasta de destino")
btnFind = ttk.Button(window, text="Selecionar Destino", command=getFolderPath)
btnFind.grid(row=0,column=0)

a = Label(window ,textvariable=folderpath)
a.grid(row=0,column = 2)

buttonFile = ttk.Button(window, text="Selecionar Comprovante", command=openFile)
buttonFile.grid(row=2, column=2)
window.mainloop()


