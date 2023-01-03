import tkinter as tk
import ttkbootstrap as ttk
from tkinter import filedialog as dlg  # Carregar arquivo
import pandas as pd
import runner
import xlsxwriter
# openpyxl

lista = []
cidades = []


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


# - - - Inicializar aplicativo - - - #
App = tk.Tk()
App.minsize(240, 270)
App.wm_geometry("1x1+550+220")
style = ttk.Style("litera")
App.title("Buscador")
img = tk.PhotoImage(file="img/R.png")
App.iconphoto(False, img, img)


# - - - - - - - - - - - - - - - - - #


def buscar(codigos: list):
    global cidades
    try:
        if codigos:
            nome.configure(text="Aguarde...", foreground="#cccc10")
            retorno = runner
            if retorno.ok:
                retorno.Execute(codigos, cidades)
                scores = retorno.data
                save = dlg.asksaveasfilename(filetypes=[("Arquivo Excel", "*.xlsx")])
                if save:
                    workbook = xlsxwriter.Workbook(f'{save}.xlsx')
                    worksheet = workbook.add_worksheet("Empresas")
                    formatado = workbook.add_format()
                    formatado.set_align("center")
                    worksheet.set_column(2, 2, 25, formatado)
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
                    nome.configure(text="Ocorreu um erro de execução, por favor\nreinicie o programa.",
                                   foreground="#ff1010", justify=tk.LEFT)
                    print("Arquivo não salvo")
            else:
                nome.configure(text="Ocorreu um erro de execução, por favor\nreinicie o programa.",
                               foreground="#ff1010", justify=tk.LEFT)
                print("Não foi possível realizar as operações")
        else:
            nome.configure(text="Primeiro importe os CNPJs.", foreground="#ff1010", justify=tk.LEFT)
            print("else numero 2")
    except:
        nome.configure(text="Não foi encontrado nenhuma empresa\nnos locais selecionados", foreground="#ff1010",
                       justify=tk.LEFT)
        print("Deu ruim")


# ------------------------Interface----------------------#

framePrincipal = tk.Frame(App)

frameSide = tk.Frame(framePrincipal)

ttk.Label(frameSide, text="Cidades").pack(padx="3", expand=1, fill="x")
# Select City
limoeiro = ttk.Checkbutton(frameSide, text="Limoeiro de Anadia", command=lambda: marcar(limoeiro, "LIMOEIRO DE ANADIA"))
limoeiro.pack(pady=5, padx=5, fill="both")
limoeiro.invoke()
limoeiro.invoke()
giral = ttk.Checkbutton(frameSide, text="Giral do Ponciano", command=lambda: marcar(giral, "GIRAL DO PONCIANO"))
giral.pack(pady=(0, 10), padx=5, fill="both")
giral.invoke()
giral.invoke()
# Select City
ttk.Separator(frameSide).pack(pady=(5, 0), expand=1, fill="x")

frameSide.pack(side="top", expand=1, fill="x")

label = ttk.Label(framePrincipal, text="Status")
label.pack(pady=(5, 0), padx="3", expand=1, fill="x")
nome = ttk.Label(framePrincipal, text="Importar Arquivo")
nome.pack(pady="5", padx="5", expand=1, fill="both")

entradaCod = ttk.Button(framePrincipal, text="Importar", width=34, command=importar, style="success-outline")
entradaCod.pack(pady="5", padx="5")

butPesquisar = ttk.Button(framePrincipal, text="Buscar", width=34, command=lambda: buscar(lista),
                          style="success-outline")
butPesquisar.pack(pady="5", padx="5")

framePrincipal.pack(padx="10", pady="10", expand=1, fill="both")

# ---------------------------------------------------------#

App.mainloop()
