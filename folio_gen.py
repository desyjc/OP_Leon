from email import message
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

from pathlib import Path
import os
import re
import PyPDF2
from numpy.core.numeric import NaN
import pandas as pd
import xlrd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from PDF_gen import CrearPdf
from orden_folio import Orden_final
from email_mime import Envio_pdf
import numpy as np


root=Tk()
root.geometry("600x600")
root.iconbitmap('./media/jiroicon.ico')
root.title('Foliador de órdenes de trabajo')
miFrame=Frame(root, widt="600", height="600")
miFrame.pack()

cwd = os.getcwd()

def Salir():
    valor=messagebox.askokcancel("Salir", "Deseas salir de la aplicacion?")
    if valor==True:
        root.destroy()

def sel_origen():
    global dir_origen
    dir_origen = filedialog.askdirectory()
    #folder_path.set(dirname)
    if dir_origen != "":
        label2.config(text = "Origen:" + dir_origen)
        dirFinal = dir_origen + "/final"
        dirfolio = dir_origen + "/folio"
        cbx_asunto.config(state='readonly')
        try:
            os.mkdir(dirFinal)
        except OSError:
            #messagebox.showwarning(message="La creación del directorio %s falló" % dirFinal)
            label_e.config(text="No se creó el directorio \Final")
        else:
            print("Se ha creado el directorio: %s " % dirFinal)
            label_e.config(text="Se creó el directorio \\final")
        try:
            os.mkdir(dirfolio)
        except OSError:
            #messagebox.showwarning(message="La creación del directorio %s falló" % dirfolio)
            label_e2.config(text="No se creó el directorio \\folio")
        else:
            print("Se ha creado el directorio: %s " % dirfolio)
            label_e2.config(text="Se creó el directorio \\folio")

    else:
        label2.config(text = "Directorio no seleccionado")
        dir_origen = ""
        cbx_asunto.config(state='disabled')
        b2.config(state='disabled')
    
    return dir_origen

def Foliar_ordenes():
    try:
        archivo_f = dir_origen + '/excel/Folios.xlsx'
        print("archivo_f: " + archivo_f)
        datos_f = pd.read_excel(archivo_f, sheet_name='Hoja1')

    except FileNotFoundError:
        messagebox.showerror(message="Archivo Folios.xlsx no encontrado", title="Error")
    
    df_folio = pd.DataFrame(datos_f)
    path = dir_origen
    arr = os.listdir(path)

    #   Iterar archivo

    for i in arr:
        archivoPdf = path + "/" + str(i)
        if os.path.isfile(archivoPdf): # para que solo lea archvos
            # extraer texto del pdf  pdf_text=extracto
            pdfFile = open(archivoPdf, 'rb')
            reader = PyPDF2.PdfFileReader(pdfFile)
            page = reader.getPage(0)
            text = page.extractText()
            PdfFileName = i

            #print('texto del pdf: ' + str( text))
            #print('Nombre del pdf :' + str (PdfFileName))
            
            # buscar la cadena del vendedor con Regex (Vendedor: \w*\s\w*\s)
            vendedor = re.compile(r'Vendedor: ((\w*\.?\s){1,6}(\w*\.?))')
            vendedor_found = vendedor.search(text)
            try:
                texto_vend = vendedor_found.group()
            except AttributeError:
                messagebox.showerror(message=f"El pdf {PdfFileName} no es una Orden válida", title="Error")

            vendedor_real = texto_vend[10:]
    
            # Buscar en Folios y extraer cadena de folio
            try:
                searchd = df_folio.query("VendNombre == @vendedor_real")
                Folio_str = (searchd.iloc[0][16])
                CrearPdf(Folio_str, dir_origen)

                Orden_final(dir_origen, Folio_str, PdfFileName)

            except IndexError:
                messagebox.showerror(message=f"El vendedor {vendedor_real} no fue encontrado en folios", title="Resultado")
    messagebox.showinfo(message="Ordenes procesadas", title="Resultado")

def sel(evento):
    global e_asunto
    selection = "Tipo de Orden de Pago: " + cbx_asunto.get()
    label.config(text = selection)
    e_asunto = 'OP ' + var.get()
    b2.config(state='normal')
    return e_asunto

def my_insert(): # adding data to Combobox
    #if e1.get() not in cbx_asunto['values']:
    cbx_asunto['values'] +=(e1.get(),) # add option
    nuevo_asunto = e1.get()
    messagebox.showinfo(message=f"Se agergó el asunto {nuevo_asunto.upper()}", title="Agregado")


def envio_pdf():
    try:
        archivo_c = cwd + '/recursos/Correos.xlsx'
        datos_c = pd.read_excel(archivo_c, sheet_name='Hoja1')
    except FileNotFoundError:
        messagebox.showerror(message="Archivo Folios.xlsx no encontrado", title="Error")

    df_correo = pd.DataFrame(datos_c)
    path = dir_origen + '/final/'
    arr = os.listdir(path)
    
    if not arr:
        #raise Exception("No existe ningun archivo en la carpeta seleccionada")
        messagebox.showerror(message="No existe ningun archivo en la carpeta seleccionada", title="Error")

    for i in arr:
        archivoPdf = path + str(i)
        if os.path.isfile(archivoPdf):
            pdfFile = open(archivoPdf, 'rb')
            reader = PyPDF2.PdfFileReader(pdfFile)
            page = reader.getPage(0)
            text = page.extractText()
            PdfFileName = i
            vendedor = re.compile(r'Vendedor: ((\w*\.?\s){1,6}(\w*\.?))')
            vendedor_found = vendedor.search(text)
            texto_vend = vendedor_found.group()
            vendedor_real = texto_vend[10:]
            searchd = df_correo.query("NombreCompleto == @vendedor_real")

            if searchd.empty:
                messagebox.showerror(message=f"No se encontró el vendedor {vendedor_real} en la lista de correos", title="Resultado")

            else:
                e_vendedor = (searchd.iloc[0][3])
                if pd.isna(e_vendedor):
                    e_vendedor = "jcdesigaud@jiro.mx"
                
                e_ejecutivo = (searchd.iloc[0][4])
                if pd.isna(e_ejecutivo):
                    e_ejecutivo = "jcdesigaud@jiro.mx"

                Envio_pdf(dir_origen, e_vendedor, e_ejecutivo, e_asunto, PdfFileName)
    messagebox.showinfo(message="Envíos terminados", title="Resultado")      


# ---------------------M E N U ----------------
barraMenu=Menu(miFrame)
root.config(menu=barraMenu, width=300, height=300)
root.title("Foliador de órdenes de trabajo")
barraMenu.config(font = ("Arial", 14))



archivoSalir=Menu(barraMenu, tearoff=0)
archivoSalir.add_command(label="Salir", command=Salir)

barraMenu.add_cascade(label="Salir", menu=archivoSalir)

#------------- B O T O N E S  -----------------------------------------

nombreNuevo=StringVar()
noCuenta=StringVar()
var = StringVar()
var.set(None)

Title = Label(miFrame, text="FOLIADO DE OPs BAJIO")
Title.config(
    fg="#F7F9F9",
    bg="#2980B9",
    padx=150,
    pady=5,
    font=("Arial", 16)
)
Title.grid(row=0, columnspan=4, sticky="ew" , padx=15, pady=10)

s2=Button(miFrame, text="Seleccionar Folder pdfs", command=sel_origen)
s2.grid(row=1, column=0, sticky="w", padx=15, pady=5)

label2 = Label(miFrame, fg="blue", bg="white", width=50)
label2.grid(row=2, column=0, sticky="w", padx=15, pady=5)

label_e = Label(miFrame, fg="blue", bg="white", width=50)
label_e.grid(row=3, column=0, sticky="w", padx=15, pady=5)

label_e2 = Label(miFrame, fg="blue", bg="white", width=50)
label_e2.grid(row=4, column=0, sticky="w", padx=15, pady=5)


b=Button(miFrame, text="Foliar PDF's", command=Foliar_ordenes)
b.config(
    fg="#1C2833",
    bg="#A3E4D7",
    font=("Arial", 14)
)
b.grid(row=2, column=1, sticky="e", padx=10, pady=5)

separator = ttk.Separator(miFrame, orient=HORIZONTAL)
separator.grid(row=5, columnspan=4, sticky="ew", pady=20)

Title = Label(miFrame, text="ENVIO DE OPs")
Title.config(
    fg="#F7F9F9",
    bg="#2980B9",
    padx=150,
    pady=5,
    font=("Arial", 16)
)
Title.grid(row=6, columnspan=4, sticky="w", padx=15, pady=10)

c = Label(miFrame, text="Seleccionar tipo de OP:")
c.grid(row=7, column=0, sticky="w", padx=15, pady=5)

cbx_asunto = ttk.Combobox(miFrame, width = 27, state='disabled', textvariable = var)
cbx_asunto['values']= ("Daños pesos",
                        "Vida pesos",
                        "Daños Dls",
                        "Vida Dls",
                        "Retenciones")
#cbx_asunto['state'] = 'disabled','readonly'
cbx_asunto.grid(row=8, column=0, sticky="w", padx=15, pady=5)
cbx_asunto.bind('<<ComboboxSelected>>', sel)

label = Label(miFrame, fg="red")
label.grid(row=9, column=0, sticky="w", padx=15, pady=5)

e1 = Entry(miFrame,bg='Yellow',width=30)
e1.grid(row=8,column=0, sticky="e")

b3 = Button(miFrame,text='Add', command=lambda: my_insert())
b3.grid(row=8,column=1, sticky="w")

b2=Button(miFrame, text="Enviar OP's por corrreo", state=DISABLED, command=envio_pdf)
b2.config(
    fg="#F4F6F7",
    bg="#F04F17",
    font=("Arial", 14)
)
b2.grid(row=10, column=0, sticky="w", padx=15, pady=20)

version = Label(miFrame, text="Hermes OP, version 1.3")
version.grid(row=12, column=1, sticky="e")
version.config(font=("Arial", 8, "italic"))
root.mainloop()