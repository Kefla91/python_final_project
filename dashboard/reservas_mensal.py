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

def reservas_mes(self):
    self.espaco_linha = Label(self.relatorio, text='')
    self.espaco_linha.grid(row=4, column=1, columnspan=3)
        
    self.etiqueta_reservas_mes = Label(self.relatorio, text="Reservas do mês", font=('Arial', 12, 'bold'), foreground='blue')
    self.etiqueta_reservas_mes.grid(row=5, column=1, columnspan=3)

    # Estilização da Tabela
    style = ttk.Style()
    style.configure("mystyle.Treeview", highlightthikness=0, bd=0, font=('Arial', 10))
    style.configure("mystyle.Treeview.Heading", font=('Arial', 11, 'bold'), padding = (0,20), anchor="center")
    style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky':'nswe'})])        # Eliminar as bordas
    
    columns = (('ID\nCliente', 60), ('ID\nReserva', 60), ('ID\nVeículo', 60), ('Data\nInício', 120), ('Data\nFim', 120),('Valor', 90), ('Pago', 90))
    self.tabela = ttk.Treeview(self.relatorio, height=5, columns=[x[0] for x in columns], style="mystyle.Treeview", show="headings")
    self.tabela.grid(row=6, column=1, columnspan=3)
    self.tabela.heading('#1', text='ID\nCliente', anchor=CENTER)
    self.tabela.heading('#2', text='ID\nReserva', anchor=CENTER)    
    self.tabela.heading('#3', text='ID\nVeículo', anchor=CENTER)
    self.tabela.heading('#4', text='Data Inínio', anchor=CENTER) 
    self.tabela.heading('#5', text='Data Fim', anchor=CENTER)
    self.tabela.heading('#6', text='Valor', anchor=CENTER)
    self.tabela.heading('#7', text='Pago', anchor=CENTER)

    # Para a largura da coluna ficar adaptada:
    for col, width in columns:
        self.tabela.heading(col, text=col)
        self.tabela.column(col, width=width, anchor=tk.CENTER)
        
    # 1 - Ao iniciar a app vamos LIMPAR a tabela se tiver dados residuais ou antigos
    registos_tabela = self.tabela.get_children()            # Obter todos os dados da tabela
    for linha in registos_tabela:
        self.tabela.delete(linha)
            
    # 2 - Consultar SQL        
    query_reservas = 'SELECT * FROM reservas ORDER BY Início DESC'
    registos_db_reservas = self.db_consulta(query_reservas)  # Faz-se a chamada ao método db_consultas

    # 3 - ESCREVER dados no ecrã:
    for linha in registos_db_reservas:
        data_reserva = datetime.strptime(linha[3], self.formato_data_num)                   # Formatar string para datetime
        diferenca_dias = data_reserva - self.dt                                           # Quanto tempo falta para a reserva
        if diferenca_dias <= self.trinta_dias and diferenca_dias <= self.dif_data_nula:
            if linha[7] == 'Sim':
                self.tabela.insert('', 'end', values=(linha[1], linha[0], linha[2],linha[3], linha[4],linha[6], linha[7]))
            else:
                self.tabela.tag_configure('aviso',foreground='red')
                self.tabela.insert('', 'end', values=(linha[1], linha[0], linha[2],linha[3], linha[4],linha[6], linha[7]), tags=('aviso',))                
    # Adiciona barra de rolagem usando grid
    scrollbar = ttk.Scrollbar(self.relatorio, orient="vertical", command=self.tabela.yview)

    # Posiciona o Treeview e a scrollbar com grid
    scrollbar.grid(row=6, column=4, sticky="ns")

    # Configura o redimensionamento
    self.tabela.grid_rowconfigure(0, weight=1)
    self.tabela.grid_columnconfigure(0, weight=1)
    
def total_financeiro(self):
    self.espacoLinha = Label(self.relatorio, text='')
    self.espacoLinha.grid(row=7, column=1)
                  
    # 2 - Consultar SQL        
    query = 'SELECT * FROM reservas ORDER BY Valor_Final DESC'
    registos = self.db_consulta(query)  # Faz-se a chamada ao método db_consultas

    total_mes=0
    # 3 - Somar pagamentos das reservas efetuadas nos ultimos 30dias:
    for linha in registos:
        data_reserva = datetime.strptime(linha[5], self.formato_data_num)                   # Formatar string para datetime
        diferenca_dias = data_reserva - self.dt                                             # Há quanto tempo reservou
        if diferenca_dias <= self.trinta_dias and diferenca_dias <= self.dif_data_nula:
            if linha[7] == 'Sim':                                                           # Se pagou
                total_mes += linha[6]
        
    self.totalFinanceiro = Label(self.relatorio, text=f"Total Mensal: {total_mes}€", font=('Arial', 15, 'bold'))
    self.totalFinanceiro.grid(row=8, column=1, columnspan=3)  
    