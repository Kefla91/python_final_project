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

def inspecao(self):
    self.etiqueta_veic_inspecao_expirar = Label(self.relatorio, text='Inspecao', font=('Arial', 12, 'bold'), foreground='blue')
    self.etiqueta_veic_inspecao_expirar.grid(row=1, column=3)
        
    # Estilização da Tabela
    style = ttk.Style()
    style.configure("mystyle.Treeview", highlightthikness=0, bd=0, font=('Arial', 10))
    style.configure("mystyle.Treeview.Heading", font=('Arial', 11, 'bold'), padding = (0,20), anchor="center")
    style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky':'nswe'})])        # Eliminar as bordas
        
    columns = (('ID\nVeículo', 60), ('Matrícula', 120), ('Data\nInspecao', 120),('Dias para\nInspecao', 90))
    self.tabela_inspecao = ttk.Treeview(self.relatorio, height=5, columns=[x[0] for x in columns], style="mystyle.Treeview", show="headings")
    self.tabela_inspecao.grid(row=2, column=3)
    self.tabela_inspecao.heading('#1', text='ID\nVeículo', anchor=CENTER)
    self.tabela_inspecao.heading('#2', text='Matrícula', anchor=CENTER)    
    self.tabela_inspecao.heading('#3', text='Data de Inspecao', anchor=CENTER)
    self.tabela_inspecao.heading('#4', text='Dias para Inspecao', anchor=CENTER)    
        
    for col, width in columns:
        self.tabela_inspecao.heading(col, text=col)
        self.tabela_inspecao.column(col, width=width, anchor=tk.CENTER)
        
    # Limpar a tabela case haja dados residuais ou antigos
    registos_tabela = self.tabela_inspecao.get_children()            # Obter todos os dados da tabela
    for linha in registos_tabela:
        self.tabela_inspecao.delete(linha)
        
    parametros = (self.tabela_inspecao.heading('#1'))
        
    # Consultar SQL        
    query = 'SELECT * FROM veiculos ORDER BY Data_Inspecao DESC'
    registos_db = self.db_consulta(query)  # Faz-se a chamada ao método db_consultas
        
    
    # Inserir dados:
    for linha in registos_db:
        data_inspecao = datetime.strptime(linha[10], self.formato_data_num)                   # Formatar string para datetime
        diferenca_dias = data_inspecao-self.dt                                           # Quanto tempo falta para a Revisão

        if diferenca_dias <= self.duas_semanas:    # Insere caso falte 15 dias ou menos para a próxima revisão
            if diferenca_dias >= self.dif_data_nula:
                self.tabela_inspecao.insert('', 'end', values=(linha[0], linha[1], linha[10], diferenca_dias.days))
            else:
                self.tabela_inspecao.tag_configure('aviso',foreground='red')
                self.tabela_inspecao.insert('', 'end', values=(linha[0], linha[1], linha[10], diferenca_dias.days), tags=('aviso',))

   
    # Scrollbar:
    scrollbar = ttk.Scrollbar(self.relatorio, orient="vertical", command=self.tabela_inspecao.yview)
    self.tabela_inspecao.configure(yscroll=scrollbar.set)

    # Posição scrollbar:
    scrollbar.grid(row=2, column=4, sticky="ns")
    self.tabela_inspecao.grid_rowconfigure(0, weight=1)
    self.tabela_inspecao.grid_columnconfigure(0, weight=1)