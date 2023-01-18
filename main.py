from pathlib import Path
import tkinter as tk
from tkinter import font
import ttkbootstrap as ttk
from tkinter import filedialog as dlg  # Carregar arquivo
import pandas as pd
import runner
import xlsxwriter

lista = []
cidades = []

bw = int()


def select(data):
    global bw
    bw = data


item = 0


def saveFile(alinhamento, scores):
    save = dlg.asksaveasfilename(filetypes=[("Arquivo Excel", "*.xlsx")])
    if save:
        workbook = xlsxwriter.Workbook(f'{save}.xlsx')
        worksheet = workbook.add_worksheet("Empresas")
        formatado = workbook.add_format()
        formatado.set_align("center")
        worksheet.set_column(2, 4, 25, formatado)
        worksheet.set_column(0, 1, 35)
        row = 0
        col = 0
        for pessoa, empresa, local, email, text in scores:
            worksheet.write(row, col, pessoa)
            worksheet.write(row, col + 1, empresa)
            worksheet.write(row, col + 2, local)
            worksheet.write(row, col + 3, email)
            worksheet.write(row, col + 4, text)
            row += 1
        workbook.close()
        nome.configure(text="Concluído!", foreground="#00aa10", justify=alinhamento)
    else:
        nome.configure(text="O arquivo não foi salvo.", foreground="#ff1010", justify=alinhamento)


def verificar(a):
    global item
    if a == "Raiz do CNPJ":
        item = 0
    elif a == "CACEAL":
        item = 1
    elif a == "CNPJ":
        item = 2


def marcar(mark, city: str):
    if mark.instate(['selected']):
        cidades.append(city)
    else:
        cidades.remove(city)


def importar():
    global lista
    try:
        arquivo = dlg.askopenfilename(filetypes=[("Planilha", "*.xlsx")])  # dialog para buscar o xlsx (arquivo excel)
        src = pd.read_excel(arquivo, header=None)
        lista = [i for i in src[0]]  # coluna dos códigos
        nome.configure(text="Importado com sucesso!", foreground="#00aa10")
    except FileNotFoundError:
        if lista:
            return
        nome.configure(text="Não foi possivel importar\nTente novamente.", foreground="#ff1010", justify=tk.LEFT)


def processo(codigos, alinhamento):
    retorno = runner
    if cidades:
        retorno.Execute(codigos, cidades, bw, item)
        scores = retorno.data
        if scores:
            saveFile(alinhamento, scores)
        else:
            nome.configure(text="Nenhuma empresa encontrada.", foreground="#ff1010", justify=alinhamento)
        scores.clear()
    else:
        nome.configure(text="Selecione pelo menos uma cidade.", foreground="#ff1010", justify=alinhamento)
    retorno.wipedata()

# - - - Inicializar aplicativo - - - #


App = tk.Tk()
App.minsize(260, 290)
style = ttk.Style("litera") 
App.title("Buscador")
img = tk.PhotoImage(file=Path(".\img\R.png"))
App.iconphoto(False, img, img)


# - - - - - - - - - - - - - - - - - #

def buscar(codigos: list):
    global cidades
    alinhamento = tk.CENTER
    try:
        if codigos:
            nome.configure(text="Aguarde...", foreground="#cccc10", justify=alinhamento)
            processo(codigos, alinhamento)
        else:
            nome.configure(text="Primeiro importe os CNPJs.", foreground="#ff1010", justify=alinhamento)
    except:
        nome.configure(text="Não foi encontrado nenhuma empresa\nnos locais selecionados", foreground="#ff1010",
                       justify=alinhamento)


# ------------------------Interface----------------------#


listaOpt = ["Raiz do CNPJ", "Raiz do CNPJ", "CACEAL", "CNPJ"]
value_inside = tk.StringVar(App)
framePrincipal = tk.Frame(App)
optbut = ttk.OptionMenu(App, value_inside, *listaOpt, style="success", command=verificar, direction="below")
verificar(optbut)
optbut.pack(fill="x")

frameSide = tk.Frame(framePrincipal)

ttk.Label(frameSide, text="Cidades:").pack(padx="3", pady=(0, 10))
# Select City
limoeiro = ttk.Checkbutton(frameSide, text="Limoeiro de Anadia", command=lambda: marcar(limoeiro, "LIMOEIRO DE ANADIA"),
                           style="success")
limoeiro.pack(pady=(0, 10), padx=5, fill="both")
limoeiro.invoke()
limoeiro.invoke()
giral = ttk.Checkbutton(frameSide, text="Giral do Ponciano", command=lambda: marcar(giral, "GIRAL DO PONCIANO"),
                        style="success")
giral.pack(pady=(0, 10), padx=5, fill="both")
giral.invoke()
giral.invoke()
taquarana = ttk.Checkbutton(frameSide, text="Taquarana", command=lambda: marcar(taquarana, "TAQUARANA"),
                            style="success")
taquarana.pack(pady=(0, 10), padx=5, fill="both")
taquarana.invoke()
taquarana.invoke()
arapiraca = ttk.Checkbutton(frameSide, text="Arapiraca", command=lambda: marcar(arapiraca, "ARAPIRACA"),
                            style="success")
arapiraca.pack(padx=5, fill="both")
arapiraca.invoke()
arapiraca.invoke()
# Select City
ttk.Separator(frameSide).pack(pady=(15, 0), expand=1, fill="x")

# --------------- linha ---------------

frameSide.pack(side="top", expand=1, fill="x")

ttk.Label(framePrincipal, text="Status:").pack(pady=(5, 0))
nome = ttk.Label(framePrincipal, text="Importar Arquivo", justify=tk.CENTER)
nome.pack(pady="5", padx="5")

frameRadials = ttk.LabelFrame(framePrincipal, text="Navegador")
browser = ttk.Radiobutton(frameRadials, text="Firefox", variable=bw, command=lambda: select(0), value=0,
                          style="success")
browser.pack(side=tk.LEFT, padx=(5, 0))
browser.invoke()
browser2 = ttk.Radiobutton(frameRadials, text="Chrome", variable=bw, command=lambda: select(1), value=1,
                           style="success")
browser2.pack(side=tk.RIGHT, padx=(0, 5))
frameRadials.pack(fill="x", ipadx="5", ipady="5", padx=5, pady=(5, 10))

entradaCod = ttk.Button(framePrincipal, text="Importar", width=34, command=importar, style="success-outline")
entradaCod.pack(pady="5", padx="5")

butPesquisar = ttk.Button(framePrincipal, text="Buscar", width=34, command=lambda: buscar(lista),
                          style="success-outline")
butPesquisar.pack(pady="5", padx="5")

version = ttk.Label(framePrincipal, text="v1.2", foreground="#999", font=font.Font(size=8)).pack()

framePrincipal.pack(padx="10", pady="10", expand=1, fill="both")

# ---------------------------------------------------------#

App.mainloop()
