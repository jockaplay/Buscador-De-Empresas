from tkinter import *
from tkinter.ttk import *
import runner
import xlsxwriter

#CNPJ Diego: 40378230

App = Tk()
App.title("Buscador")

def buscar():
    retorno = runner
    retorno.execute(f'{entradaText.get()}')
    
    res.configure(text=f'{retorno.datareturn()}')
    
    workbook = xlsxwriter.Workbook('Relatório.xlsx') 
    worksheet = workbook.add_worksheet("Empresas") 
    scores = []
    scores.append(retorno.datareturn())
    scores.append(retorno.datareturn())
    row = 0
    col = 0
    for i, j, k in scores:
        worksheet.write(row, col, i)
        worksheet.write(row, col + 1, j)
        worksheet.write(row, col + 2, k)
        row += 1
    workbook.close()

framePrincipal = Frame(App)

nome = Label(framePrincipal, text="Carregar Arquivo")
nome.pack(pady="5", padx="5")

entradaText = Entry(framePrincipal, width='34')
entradaText.pack(pady="5", padx="5")

butPesquisar = Button(framePrincipal, text="Buscar", width='34', command=lambda:buscar())
butPesquisar.pack(pady="5", padx="5")

res = Label(framePrincipal, text="")
res.pack(pady="5", padx="5")

framePrincipal.pack(padx="10", pady="10")

App.mainloop()