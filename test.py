from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as dlg # Carregar arquivo
import os
import pandas as pd

def importar():
    arquivo = dlg.askopenfilename(filetypes=[("Planilha", "*.xlsx")]) # dialog para buscar o xlsx (arquivo excel)
    arq = pd.read_excel(arquivo, header=None)
    lista = [i for i in arq[0]]
    print(lista)

App = Tk()
App.title("Buscador")

lista = []

framePrincipal = Frame(App)

nome = Label(framePrincipal, text="Carregar Arquivo")
nome.pack(pady="5", padx="5")

entradaCod = Button(framePrincipal, text="Importar", width=34, command=importar)
entradaCod.pack(pady="5", padx="5")

butPesquisar = Button(framePrincipal, text="Buscar", width=34)
butPesquisar.pack(pady="5", padx="5")


framePrincipal.pack(padx="10", pady="10")


#---------------------------------------------------------#

App.mainloop()