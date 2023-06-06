import os.path
import pandas as pd
import openpyxl
import database
from tkinter import *
from tkinter import ttk
from mundial_bot import *
from os import *

janela = Tk()
mundial_bot = Mundial()
pastaApp = os.path.dirname(__file__)

class Tela:
    def __init__(self):
        self.janela = janela
        self.tela()
        self.frames()
        self.botoes()
        self.labels()
        self.entrada()
        self.lista_frame2()
        self.get_database()
        self.dropdown()
        self.csv()
        self.excel()
        self.atualizar = mundial_bot.database_inserir()

        janela.mainloop()

    def tela(self):
        self.janela.title("www.mundialcalcados.com.br")
        self.janela.configure(background="#fff")
        self.janela.resizable(True, True)
        self.imagem_logo()

    def frames(self):
        self.frame_1 = Frame(self.janela, bg='grey', highlightthickness=0.5, highlightbackground="#fff")
        self.frame_1.place(relx=0.03, rely=0.05, relwidth=0.94, relheight=0.40)

        self.frame_2 = Frame(self.janela, bg='grey', highlightthickness=0.5, highlightbackground="#fff")
        self.frame_2.place(relx=0.03, rely=0.50, relwidth=0.94, relheight=0.45)

    def botoes(self):
        self.botao_atualizar = Button(self.frame_1, font=("GothamPro", 13), background="#fff", text="Atualizar", command= self.database_atualizar)
        self.botao_atualizar.place(relx=0.63, rely=0.56, relwidth=0.1, relheight=0.13)

        self.botao_pesquisar = Button(self.frame_1, font=("GothamPro", 13), background="#fff", text="Pesquisar", command= self.database_pesquisar)
        self.botao_pesquisar.place(relx=0.27, rely=0.56, relwidth=0.1, relheight=0.13)

        self.botao_exportar = Button(self.frame_1, font=("GothamPro", 13), background="#fff", text="Exportar", command=self.botao_exportar)
        self.botao_exportar.place(relx=0.63, rely=0.80, relwidth=0.1, relheight=0.13)

    def labels(self):
        self.lb_input = Label(self.frame_1, font=("GothamPro", 15), text='Fabricante', bg='grey', fg="#fff")
        self.lb_input.place(relx=0.40, rely=0.40, relwidth=0.20, relheight=0.15)

        self.lbStatus = Label(self.frame_1, image=self.logo)
        self.lbStatus.place(relx=0.05, rely=0.05, relwidth=0.90, relheight=0.21)

    def entrada(self):
        self.entrada_usuario = Entry(self.frame_1)
        self.entrada_usuario.place(relx=0.40, rely=0.57, relwidth=0.2, relheight=0.13)

    def lista_frame2(self):
        self.listaCli = ttk.Treeview(self.frame_2, height=3, columns=('col0', 'col1', 'col2','col3'))
        self.listaCli.heading('#0', text='')
        self.listaCli.heading('#1', text='Marca')
        self.listaCli.heading('#2', text='Modelo')
        self.listaCli.heading('#3', text='Valor')

        self.listaCli.column('#0', width=5)
        self.listaCli.column('#1', width=35)
        self.listaCli.column('#2', width=188)
        self.listaCli.column('#3', width=188)

        self.listaCli.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.listaCli.configure(yscrollcommand=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.85)

    def database_pesquisar(self):
        self.listaCli.delete(*self.listaCli.get_children())
        for marcas in database.cur.execute(f"SELECT * FROM tenis WHERE marca LIKE '%{self.entrada_usuario.get()}%'"):
            self.listaCli.insert(parent="", index=0, values=marcas)
        database.var.commit()

    def get_database(self):
        for coluna in database.cur.execute(f"SELECT * FROM tenis"):
            self.listaCli.insert(parent="", index=0, values=coluna)
        database.var.commit()

    def database_atualizar(self):
        Mundial()
        mundial_bot.database_inserir()
        mundial_bot.database_delete()
        mundial_bot.database_inserir()
    def imagem_logo(self):
        self.logo = PhotoImage(file=pastaApp+"//logo.png")

    def dropdown(self):
        options = [
            "CSV",
            "XLSX"
        ]
        self.clicked = StringVar()
        self.clicked.set("Escolha um formato")
        drop = OptionMenu(self.frame_1, self.clicked, *options)
        drop.place(relx=0.25, rely=0.80, relwidth=0.20, relheight=0.15)

    def csv(self):
        lista_marca = []
        lista_nome = []
        lista_preco = []

        for i in database.cur.execute(f"SELECT * FROM tenis WHERE marca LIKE '%{self.entrada_usuario.get()}%'"):
            lista_marca.append(i[0])
            lista_nome.append(i[1])
            lista_preco.append(i[2])
            print(lista_marca)
            print(lista_preco)
            print(lista_nome)

            self.cabecalho = {'Marca': lista_marca,
                              'modelo': lista_nome,
                              'valor': lista_preco}

            data = pd.DataFrame(data=self.cabecalho)

            data.to_csv(f'C:/Users/47190845836/Documents/teste/{self.entrada_usuario.get()}.csv', sep=";", index=False)

    def excel(self):
        lista_marca = []
        lista_nome = []
        lista_preco = []

        for i in database.cur.execute(f"SELECT * FROM tenis WHERE marca LIKE '%{self.clicked.get().lower()}%'"):
            lista_marca.append(i[0])
            lista_nome.append(i[1])
            lista_preco.append(i[2])
            print(lista_marca)
            print(lista_preco)
            print(lista_nome)

            self.cabecalho = {'Marca': lista_marca,
                              'modelo': lista_nome,
                              'valor': lista_preco}

            data = pd.DataFrame(data=self.cabecalho)

            data.to_excel(f'C:/Users/47190845836/Documents/teste/{self.entrada_usuario.get()}.xlsx', index=False)

    def botao_exportar(self):
        if self.clicked.get() == "CSV":
            return self.csv()
        elif self.clicked.get() == "XLSX":
            return self.excel()


Tela()
