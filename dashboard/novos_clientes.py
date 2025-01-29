import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
from tkinter.messagebox import showinfo

import sqlite3

from datetime import datetime, date, time, timedelta    # Para poder utilizar dados de data e hors
import locale                                           # Para poder utilizar data em Pt-pt
from tkcalendar import Calendar, DateEntry                         # Para poder escolher datas com um calendário

from random import choice
import matplotlib
matplotlib.use('TkAgg',force=True)
import matplotlib.pyplot as plt                                 # Para poder trabalhar com gráficos
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg # Para o gráfico aparecer na app, não em pop-up
#FigureCanvasTkAgg.use('TkAgg',force=True)
import numpy as np                                      # Para usar algumas fórmulas

def clientes_novos(self):
    self.espaco=Label(self.relatorio, text='')
    self.espaco.grid(row=19, column=1, columnspan=3)
        
    self.etiqueta_ultimos_clientes = Label(self.relatorio, text="Últimos Clientes Registados", font=('Arial', 12, 'bold'), foreground='blue')
    self.etiqueta_ultimos_clientes.grid(row=20, column=1, columnspan=3)

    # Estilização da Tabela
    style = ttk.Style()
    style.configure("mystyle.Treeview", highlightthikness=0, bd=0, font=('Arial', 10))
    style.configure("mystyle.Treeview.Heading", font=('Arial', 11, 'bold'), padding = (0,20), anchor="center")
    style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky':'nswe'})])        # Eliminar as bordas
    
    columns = (('ID\nCliente', 60), ('Nome', 120), ('Sobrenome', 120), ('Telemóvel', 120), ('E-mail', 250))
    self.tabela = ttk.Treeview(self.relatorio, height=5, columns=[x[0] for x in columns], style="mystyle.Treeview", show="headings")
    self.tabela.grid(row=21, column=1, columnspan=3)
    self.tabela.heading('#1', text='ID\nCliente', anchor=CENTER)
    self.tabela.heading('#2', text='Nome', anchor=CENTER)    
    self.tabela.heading('#3', text='Sobrenome', anchor=CENTER)   
    self.tabela.heading('#4', text='Telemóvel', anchor=CENTER)   
    self.tabela.heading('#5', text='E-mail', anchor=CENTER)    
    
    for col, width in columns:
        self.tabela.heading(col, text=col)
        self.tabela.column(col, width=width, anchor=tk.CENTER)
        
    # Consultar SQL        
    query_IDCliente = 'SELECT * FROM clientes ORDER BY IDCliente DESC'  # ID pela ordem decrescente
    registos_db_clientes = self.db_consulta(query_IDCliente)            # Faz-se a chamada ao método db_consultas

    contagem_ID = 1
    for linha in registos_db_clientes:
        if contagem_ID <=5:                         # 5 Clientes com maior ID => Últimos 5 clientes
            self.tabela.insert('', 'end', values=(linha[0], linha[1], linha[2], linha[4], linha [5]))
            contagem_ID +=1
