from tkinter import *
from tkinter.ttk import *
import runner

#CNPJ Diego: 40378230

App = Tk()
App.title("Buscador")

framePrincipal = Frame(App)

nome = Label(framePrincipal, text="Código")
nome.pack(pady="5", padx="5")

entradaText = Entry(framePrincipal, width='34')
entradaText.pack(pady="5", padx="5")

butPesquisar = Button(framePrincipal, text="Buscar", width='34', command = lambda:runner.execute(f'{entradaText.get()}'))
butPesquisar.pack(pady="5", padx="5")

framePrincipal.pack(padx="10", pady="10")

App.mainloop()
