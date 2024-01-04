from datetime import date
import tkinter
from tkinter import *
from tkinter import ttk

from models.busca15min import Busca15min
from models.buscaDiaria import BuscaDiaria
from models.interpolacao15min import Interpolar_15min
from models.interpolacaoDiaria import InterpolarDiarias
try:
    class Principal():
        # Seleção da estação de consulta
        def executa(self):
            data1 = self.entry2.get()
            data2 = self.entry3.get()
            codigo1 = self.entry4.get()
            nomesalvar = self.entry5.get()
            nomearquivo = self.entry6.get()
            data3 = self.entry7.get()
            intervalo = self.rb_value.get()
            interpolar = self.cb_value.get()
            resultado = str("Arquivo excel criado com sucesso!!!")
            self.v.set("")
            Window.update()
            # Validação de dados
            if ((data1 == "") or (data2 == "") or (codigo1 == "") or (nomesalvar == "") or (intervalo == 0)):
                validacao1 = str("Preencha os campos corretamente!!!")
                self.v.set(validacao1)
                Window.update()
            elif intervalo == 1 and interpolar == False:  # Dados 15 minutos
                objeto = Busca15min()
                instancia = objeto.buscar(
                    data1, data2, codigo1, nomesalvar)
                self.v.set(resultado)
            elif intervalo == 2 and interpolar == False:  # Media Diaria
                objeto = BuscaDiaria()
                instancia = objeto.buscar(
                    data1, data2, codigo1, nomesalvar)
                self.v.set(resultado)
            elif intervalo == 1 and interpolar == True and nomearquivo != "" and data3 != "":  # Interpolacao Dados 15 minutos
                objeto = Interpolar_15min()
                instancia = objeto.calcular(
                    data1, data2, data3, nomearquivo, nomesalvar)
                self.v.set(resultado)
            elif intervalo == 2 and interpolar == True and nomearquivo != "" and data3 != "":  # Interpolacao Dados Diarios
                objeto = InterpolarDiarias()
                instancia = objeto.calcular(
                    data1, data2, data3, nomearquivo, nomesalvar)
                self.v.set(resultado)
            else:
                validacao2 = str(
                    "Preencha as informações para interpolação de dados!!!")
                self.v.set(validacao2)
                Window.update()

        def limpar(self):
            self.entry2.delete(0, END)
            self.entry3.delete(0, END)
            self.entry4.delete(0, END)
            self.entry5.delete(0, END)
            self.entry6.delete(0, END)
            self.entry6.config(state="disabled")
            self.entry7.delete(0, END)
            self.entry7.config(state="disabled")
            self.rb_value.set(0)
            self.cb_value.set(False)
            self.v.set("")
            self.entry2.focus()
            Window.update()

        def habilitar(self):
            if self.cb_value.get():
                self.entry6.config(state="normal")
                self.entry7.config(state="normal")
            else:
                self.entry6.config(state="disabled")
                self.entry7.config(state="disabled")
        # Parâmetros da tela
        def __init__(self, instancia_de_Tk):
            frame1 = tkinter.Frame(instancia_de_Tk)
            frame1.configure(border=5)
            frame1.pack()
            frame2 = Frame(instancia_de_Tk)
            frame2.configure(border=5)
            frame2.pack()
            frame3 = Frame(instancia_de_Tk)
            frame3.configure(border=5)
            frame3.pack()
            frame4 = Frame(instancia_de_Tk)
            frame4.configure(border=5)
            frame4.pack()
            frame5 = Frame(instancia_de_Tk)
            frame5.configure(border=5)
            frame5.pack()
            containerP = Frame(instancia_de_Tk)
            containerP.configure(bd=3, relief="groove")
            containerP.pack()
            containerE = Frame(master=containerP)
            containerE.configure(border=5)
            containerE.pack(side="left")
            containerD = Frame(master=containerP)
            containerD.configure(border=5)
            containerD.pack(side="right")
            frame6 = Frame(master=containerE)
            frame6.configure(border=5)
            frame6.pack(side="left")
            frame7 = Frame(master=containerD)
            frame7.configure(border=5)
            frame7.pack(side="right")
            frame8 = Frame(instancia_de_Tk)
            frame8.configure(border=5)
            frame8.pack()
            frame9 = Frame(instancia_de_Tk)
            frame9.configure(border=5)
            frame9.pack()
            divisoria = ttk.Separator(containerP, orient="vertical")
            divisoria.pack(fill="y", expand=True)
            # Parâmetros dos dados apresentados na tela
            label1 = Label(frame1, text="Formato de data: YYYY-mm-dd HH:mm:ss")
            label1.pack()
            label2 = Label(
                frame2, text="Período de inicio: ")
            label2.pack()
            self.entry2 = Entry(frame2, width=30)
            self.entry2.pack()
            label3 = Label(
                frame3, text="Período final: ")
            label3.pack()
            self.entry3 = Entry(frame3, width=30)
            self.entry3.pack()
            label4 = Label(frame4, text="Código da estação: ")
            label4.pack()
            self.entry4 = Entry(frame4, width=5)
            self.entry4.pack()
            label5 = Label(
                frame5, text="Nome para salvar o arquivo: ")
            label5.pack()
            self.entry5 = Entry(frame5, width=15)
            self.entry5.pack()
            self.rb_value = IntVar()
            self.rb1 = Radiobutton(
                frame6, text="Dados 15 minutos", value=1, variable=self.rb_value).pack()
            self.rb2 = Radiobutton(
                frame6, text="Dados Média Diária", value=2, variable=self.rb_value).pack()
            self.cb_value = BooleanVar()
            self.cb1 = Checkbutton(
                frame7, text="Interpolar", variable=self.cb_value, command=lambda: self.habilitar()).pack(pady=10)
            label6 = Label(frame7, text="Nome do arquivo: ")
            label6.pack()
            self.entry6 = Entry(frame7, width=15, state="disabled")
            self.entry6.pack()
            label7 = Label(
                frame7, text="Período para iniciar a interpolação: ")
            label7.pack()
            self.entry7 = Entry(frame7, width=30, state="disabled")
            self.entry7.pack()
            label8 = Label(frame8, text="Resultado: ")
            label8.pack()
            self.v = StringVar()
            label9 = Label(frame8, text="", textvariable=self.v,
                           background="white", font="14")
            label9.pack()
# Parâmetros de execução
            button1 = Button(frame9, text="Buscar", borderwidth=5,
                             command=lambda: self.executa())
            button1.pack(side="right")

            button2 = Button(frame9, text="Limpar",
                             borderwidth=5, command=lambda: self.limpar())
            button2.pack(side="left")
# Parâmetros da tela
    Window = tkinter.Tk()
    Window.title("Dados Planilha")
    Window.geometry("400x500")
    Principal(Window)
    Window.mainloop()
except OSError as e:
    print("Erro: ", e)
