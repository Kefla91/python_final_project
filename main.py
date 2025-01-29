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

import openpyxl                                         # Para poder trabalhar com ficheiros Excel
from openpyxl import Workbook

import csv                                              # Para poder trabalhar com ficheiros Csv

class LuxuryWeels:
    
    db = "database/LuxuryWeels.db"    
    
    from menu.veiculos import veiculos, get_veiculos, novo_veiculo, add_veiculo, confirmar_eliminar_veiculos, eliminar_veiculos, editar_veiculo, confirmar_edit_veiculo, guardar_edit_veiculo, exportar_veiculos, exp_veiculo_excel, exp_veiculo_csv
    from menu.clientes import clientes, get_clientes, novo_cliente, add_cliente, confirmar_eliminar_cliente, eliminar_cliente, editar_clientes, confirmar_edit_cliente, guardar_edit_cliente, exportar_clientes, exp_clientes_excel, exp_clientes_csv
    from menu.formas_pag import formas_pagamento, get_formas_pag, nova_formaPagamento, add_formaPagamento, editar_formaPagamento, confirmar_edit_formaPag, guardar_edit_formaPag, confirmar_eliminar_formaPag, eliminar_formaPag, exportar_formaPag, exp_formaPag_excel, exp_formaPag_csv
    from menu.reservas import reservas, get_reservas, nova_reserva, add_reserva, confirmar_eliminar_reserva, eliminar_reserva, editar_reservas, confirmar_edit_reservas, guardar_edit_reservas, exportar_reservas, exp_reservas_excel, exp_reservas_csv

    from dashboard.inspecao import inspecao
    from dashboard.novos_clientes import clientes_novos
    from dashboard.reservas_mensal import reservas_mes, total_financeiro
    from dashboard.revisao import revisao
    from dashboard.veic_alug_disp import veic_alugados, get_veic_alugados, veic_disponiveis
    
    # Método Construtor:
    def __init__(self, root):
              
        self.janela = root
        self.janela.title("Gestor Luxury Weels")        # Título da janela
        self.janela.geometry('1100x980')
        self.janela.resizable(0,0)                      # Activar redimensionamento da janela. Para ativá-la:(1,1)  
        self.janela.iconbitmap('recursos/icon.ico')     # Ícon da janela
        
        # Data e hora:
        locale.setlocale(locale.LC_ALL, 'pt-PT.UTF-8') # Idioma da data em português
        self.dt = datetime.now() # Agora
        self.hora_actual = self.dt.strftime("%H:%M")
        self.data_actual_extenso = self.dt.strftime("%A, %d de %B de %Y")
        self.formato_data_num = "%d/%m/%Y"
        self.data_actual_numerica = self.dt.strftime(self.formato_data_num)
        self.dif_data_nula = timedelta(days=0)
        self.cinco_dias = timedelta(days=5)
        self.duas_semanas = timedelta(days=15)
        self.trinta_dias = timedelta(days=30)
        
        # Criação do Menu:
        menu= tk.Canvas(self.janela)
        menu.grid(row=0, column=0, pady=20, sticky=N)
        # Botões do Menu:
        self.botão_veiculos = ttk.Button(menu, text="Veículos", command=self.veiculos)      # Botão Veículos
        self.botão_veiculos.grid(row=1, column=0)
        self.botão_clientes = ttk.Button(menu, text="Clientes", command=self.clientes)      # Botão Clientes
        self.botão_clientes.grid(row=2, column=0)
        self.botão_reservas = ttk.Button(menu, text="Reservas", command=self.reservas)      # Botão Reservas
        self.botão_reservas.grid(row=3, column=0)
        self.botão_formasPag = ttk.Button(menu, text="Formas de\nPagamento", command=self.formas_pagamento) # Botão Formas de Pagamento
        self.botão_formasPag.grid(row=4, column=0)
        
        # Criação DashBoard:
        self.relatorio = LabelFrame(self.janela, text=f"Relatório\n{self.data_actual_extenso}, {self.hora_actual}", font=('Arial', 12, 'bold'))
        self.relatorio.grid(row=0, column=3, pady=20)
        
        self.revisao()
        self.inspecao()
        self.reservas_mes()
        self.total_financeiro()
        self.veic_alugados()
        self.veic_disponiveis()
        self.clientes_novos()
        
        # Mensagem de aviso:
        self.aviso_revisao()
        
    # Método para consulta SQL:  
    def db_consulta(self, consulta, parametros = ()):
        with sqlite3.connect(self.db) as con:                   # Inicia-se uma connexão c a base de dados (con)
            self.cursor = con.cursor()                          # Cria-se um cursor da conexão p poder operar na base de dados
            self.cursor.execute(consulta, parametros)           # Prepara-se a consulta SQL (c parâmetros se os há)
            resultado = self.cursor.fetchall()
            con.commit()                                        # Executa-se a consulta SQL pereparada anteriormente
        return resultado                                        # Restitui-se o resultado da consulta
    
    # Método para validação de dados:
    def validacao_dados(self, parametros = ()):
        validacao = True
        for parametro in parametros:
            if not parametro: # Se algum parâmetro estiver em branco
                validacao = False
                break
        return validacao
    
    def aviso_revisao(self):
              
        query = 'SELECT * FROM veiculos ORDER BY Data_Revisao DESC'
        registos_db = self.db_consulta(query)
        
        for linha in registos_db:
            data_revisao = datetime.strptime(linha[9], self.formato_data_num)                   # Formatar string para datetime
            diferenca_dias = data_revisao-self.dt                                               # Quanto tempo falta para a Revisão
                
            if linha[11] == "Sim" and diferenca_dias <= self.cinco_dias:    # Insere caso falte 5 dias ou menos para a próxima revisão, e ainda não está como "não disponível"
                if diferenca_dias >= self.dif_data_nula:
                    messagebox.OK = 'yesno'                 # Botões para colocar como não disponivel, ou manter disponivel
                    self.indisponivel = messagebox.showwarning("Aviso", f"O Veículo com matrícula {linha[1]} precisa de revisão em {diferenca_dias.days} dias.\nPretende colocá-lo como indisponível?")
                    if self.indisponivel == 'yes':
                        #self.veic_indisponivel()
                        query=f'UPDATE veiculos SET Disponivel=? WHERE Matrícula=?'
                        parametros = ('Não', linha[1])
                        self.db_consulta(query, parametros)      # Actualizar Disponibilidade para "Não"
                        #self.db_consulta.close()
                        self.veic_disponiveis()
    

if __name__ == '__main__':
    root = Tk()             # Instância da janela principal
    app = LuxuryWeels(root)
    root.mainloop()         # Começamos o ciclo de aplicação. É como um while true

