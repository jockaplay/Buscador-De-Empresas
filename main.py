from tkinter import *
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog as dlg # Carregar arquivo
import pandas as pd
import runner
import xlsxwriter

#CNPJ Diego: 40378230

lista = []

def importar():
    global lista
    try:
        arquivo = dlg.askopenfilename(filetypes=[("Planilha", "*.xlsx")]) # dialog para buscar o xlsx (arquivo excel)
        arq = pd.read_excel(arquivo, header=None)
        lista = [i for i in arq[0]] # coluna dos códigos
        nome.configure(text="Importado com sucesso!", foreground="#00aa10")
    except:
        nome.configure(text="Não foi possivel importar\nTente novamente.", foreground="#ff1010", justify=CENTER)

App = Tk()
App.minsize(240, 130)
App.geometry("240x130")
style = ttk.Style("litera")
App.title("Buscador")
App.iconbitmap("img\R-_1_.ico")
img = PhotoImage(file="img\R.png")
App.iconphoto(False, img, img)

# arquivo = dlg.askopenfile(mode='r') # dialog para carregar xlsx (arquivo excel)
# print(arquivo)
# arquivo.close()

def buscar(codigos:list):
    try:
        if codigos != []:
            retorno = runner
            if retorno.ok == True:
                scores = []
                retorno.execute(codigos)
                
                scores = retorno.data
                #  for i in scores:                              #
                #      res = Label(framePrincipal, text="")      #   Adiciona uma label falando se a empresa é local ou 
                #      res.pack(pady="5", padx="5")              #   externa para cada pesquisa.
                #      res.configure(text=f'{i[3]}')             #
                    
                workbook = xlsxwriter.Workbook('Relatório.xlsx')
                worksheet = workbook.add_worksheet("Empresas")
                format = workbook.add_format()
                format.set_align("center")
                worksheet.set_column(2, 2, 25, format)
                worksheet.set_column(0, 1, 35)
                row = 0
                col = 0
                for pessoa, empresa, local, l in scores:
                    worksheet.write(row, col, pessoa)
                    worksheet.write(row, col + 1, empresa)
                    worksheet.write(row, col + 2, local)
                    row += 1
                workbook.close()
                nome.configure(text="Concluído!", foreground="#00aa10")
            else:
                nome.configure(text="Ocorreu um erro de execução, por favor\nreinicie o programa.", foreground="#ff1010", justify=CENTER)
        else:
            nome.configure(text="Primeiro importe os CNPJs.", foreground="#ff1010", justify=CENTER)
    except:
        print("Deu ruim")
    

#------------------------Interface----------------------#

framePrincipal = Frame(App)

nome = ttk.Label(framePrincipal, text="Importar Arquivo")
nome.pack(pady="5", padx="5")

entradaCod = ttk.Button(framePrincipal, text="Importar", width=34, command=importar, bootstyle=(SUCCESS, OUTLINE))
entradaCod.pack(pady="5", padx="5")

butPesquisar = ttk.Button(framePrincipal, text="Buscar", width=34, command=lambda:buscar(lista), bootstyle=(SUCCESS, OUTLINE))
butPesquisar.pack(pady="5", padx="5")


framePrincipal.pack(padx="10", pady="10")

#---------------------------------------------------------#

App.mainloop()