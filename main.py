import tkinter as tk
import ttkbootstrap as ttk
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
        src = pd.read_excel(arquivo, header=None)
        lista = [i for i in src[0]] # coluna dos códigos
        nome.configure(text="Importado com sucesso!", foreground="#00aa10")
    except:
        nome.configure(text="Não foi possivel importar\nTente novamente.", foreground="#ff1010", justify=tk.CENTER)


# - - - Inicializar aplicativo - - - #
App = tk.Tk()
App.minsize(240, 130)
style = ttk.Style("litera")
App.title("Buscador")
img = tk.PhotoImage(file="img/R.png")
App.iconphoto(False, img, img)
# - - - - - - - - - - - - - - - - - #


def buscar(codigos:list):
    try:
        if codigos != []:
            retorno = runner
            if retorno.ok == True:
                nome.configure(text="Aguarde...", foreground="#cccc10")
                scores = []
                retorno.execute(codigos)
                
                scores = retorno.data
                save = dlg.asksaveasfilename(filetypes=[("Arquivo Excel", "*.xlsx")])
                if save:
                    workbook = xlsxwriter.Workbook(f'{save}.xlsx')
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
                    nome.configure(text="Ocorreu um erro de execução, por favor\nreinicie o programa.", foreground="#ff1010", justify=tk.CENTER)
            else:
                nome.configure(text="Ocorreu um erro de execução, por favor\nreinicie o programa.", foreground="#ff1010", justify=tk.CENTER)
        else:
            nome.configure(text="Primeiro importe os CNPJs.", foreground="#ff1010", justify=tk.CENTER)
    except:
        print("Deu ruim")
    

#------------------------Interface----------------------#

framePrincipal = tk.Frame(App)

nome = ttk.Label(framePrincipal, text="Importar Arquivo")
nome.pack(pady="5", padx="5")

entradaCod = ttk.Button(framePrincipal, text="Importar", width=34, command=importar, style="success-outline")
entradaCod.pack(pady="5", padx="5")

butPesquisar = ttk.Button(framePrincipal, text="Buscar", width=34, command=lambda:buscar(lista), style="success-outline")
butPesquisar.pack(pady="5", padx="5")


framePrincipal.pack(padx="10", pady="10")

#---------------------------------------------------------#

App.mainloop()