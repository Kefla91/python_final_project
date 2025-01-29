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

# Método Nova janela Formas de Pagamento 
def formas_pagamento(self):
    self.janela_formas_pagamento = Toplevel()                   # Criat uma janela à frente da principal
    self.janela_formas_pagamento.title = "Formas de Pagamento"  # Título da janela
    self.janela_formas_pagamento.resizable(1,1)                 # Ativar a resimensão da janela
    self.janela_formas_pagamento.iconbitmap('recursos/icon.ico')  # Ícon da janela

    titulo = LabelFrame(self.janela_formas_pagamento, text = 'Formas de Pagamento', font=('Arial', 14, 'bold'))
    titulo.grid(column=0, row=0)

    self.colunas_formaPag = (("ID Forma\nPagamento", 120), ("Modo", 120))
        
    #botão Registar Novo
    self.botão_registar_novo = ttk.Button(titulo, text="Nova Forma de Pagamento", command=self.nova_formaPagamento)
    self.botão_registar_novo.grid(row=2, column=0, sticky=W+E)
    #botão Editar/Alterar
    botao_editar = ttk.Button(titulo, text='Editar', command=self.editar_formaPagamento)
    botao_editar.grid(row=2, column=1, sticky=W+E)
    #botão Remover
    botao_eliminar = ttk.Button(titulo, text = 'Eliminar', command= self.confirmar_eliminar_formaPag)
    botao_eliminar.grid(row=2, column=2, sticky=W+E)
    #botão Exportar Infos
    self.botão_exp_infos = ttk.Button(titulo, text="Exportar Informações", command=self.exportar_formaPag)
    self.botão_exp_infos.grid(row=2, column=3, sticky=W+E)

    # Mensagem informativa para o utilizador:
    self.mensagem = Label(self.janela_formas_pagamento, text = '', fg = 'green')
    self.mensagem.grid(row=4, column=0, columnspan=3, sticky=W + E) 
        
    # Criação Tabela Formas de Pagamento
    style = ttk.Style()
    style.configure("mystyle.Treeview", highlightthikness=0, bd=0, font=('Arial', 11))
    style.configure("mystyle.Treeview.Heading", font=('Arial', 12, 'bold'))
    style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky':'nswe'})])        # Eliminar as bordas
        
    self.tabela_formaPag = ttk.Treeview(self.janela_formas_pagamento, height=10, columns=[x[0] for x in self.colunas_formaPag], style="mystyle.Treeview", show="headings")
    self.tabela_formaPag.grid(row=5, column=0, columnspan=2)
    self.tabela_formaPag.heading('#1', text='ID\nForma Pagamento', anchor=CENTER)
    self.tabela_formaPag.heading('#2', text='Modo', anchor=CENTER) 
    
    for col, width in self.colunas_formaPag:
        self.tabela_formaPag.heading(col, text=col)
        self.tabela_formaPag.column(col, width=width, anchor=tk.CENTER)
        
    #Scrollbar:        
    scrollbar = ttk.Scrollbar(self.janela_formas_pagamento, orient="vertical", command=self.tabela_formaPag.yview)
    # Posição scrollbar:
    scrollbar.grid(row=5, column=1, sticky="ns")
    self.janela_formas_pagamento.grid_rowconfigure(0, weight=1)
    self.janela_formas_pagamento.grid_columnconfigure(0, weight=1)    

    self.get_formas_pag()
    
def get_formas_pag(self):
    # 1 - Ao iniciar a app vamos LIMPAR a tabela se tiver dados residuais ou antigos
    registos_tabela = self.tabela_formaPag.get_children()            # Obter todos os dados da tabela
    for linha in registos_tabela:
        self.tabela_formaPag.delete(linha)
            
    # 2 - Consultar SQL        
    query = 'SELECT * FROM formas_pagamento ORDER BY Modo DESC'
    registos_db = self.db_consulta(query)  # Faz-se a chamada ao método db_consultas

    # 3 - ESCREVER dados no ecrã:
    for linha in registos_db:
        self.tabela_formaPag.insert('', 'end', values=(linha[0], linha[1]))
    
#--------------------------------------------- NOVO ---------------------------------------------
# Método Janela Nova Forma de Pagamento:
def nova_formaPagamento(self):
    self.janela_nova_formaPagamento = Toplevel()       # Criar uma janela à frente da principal
    self.janela_nova_formaPagamento.title = "Adicionar Forma de Pagamento" # Título da janela
    self.janela_nova_formaPagamento.resizable(1,1)     # Ativar a resimensão da janela
        
    frame_nova_formaPagamento = LabelFrame(self.janela_nova_formaPagamento, text = 'Adicionar Forma de Pagamento', font=('Arial', 14, 'bold'))
    frame_nova_formaPagamento.grid(column=0, row=0)
      
    # Label Modo
    self.etiqueta_modo = Label(frame_nova_formaPagamento, text='Modo')
    self.etiqueta_modo.grid(row=1, column=0)
    self.modo = Entry(frame_nova_formaPagamento)
    self.modo.grid(row=1, column=1)
        
    # Botão Guardar:
    self.botao_add_formaPagamento = ttk.Button(frame_nova_formaPagamento, text='Adicionar Forma de Pagamento', command=self.add_formaPagamento)
    self.botao_add_formaPagamento.grid(row=2, columnspan=2, sticky=W+E)
    
# Método para guardar Nova Forma de Pagamento:
def add_formaPagamento(self):
    parametros = (self.modo.get(),)
    if self.validacao_dados(parametros) == True:
        query = 'INSERT INTO formas_pagamento VALUES(NULL, ?)'
        self.db_consulta(query, parametros)
        self.janela_nova_formaPagamento.destroy()
        self.janela_formas_pagamento.destroy()
        self.formas_pagamento()
    else:
        showinfo("AVISO","O Cliente não foi adicionado, pois faltam dados obrigatórios.")
    
#-------------------------------------------- EDITAR -------------------------------------------- 
# Método Janela de Edição
def editar_formaPagamento(self):
    self.mensagem['text'] = ''
    try:
        self.tabela_formaPag.item(self.tabela_formaPag.selection()[0])
    except IndexError as e:
        self.mensagem['text']='Por favor, seleccione uma forma de pagamento'
        return
        
    #Guardar os dados antigos em listas:
    formaPag_id = self.tabela_formaPag.item(self.tabela_formaPag.selection())['values'][0]
    modo = self.tabela_formaPag.item(self.tabela_formaPag.selection())['values'][1]
    
    # Criação Janela para Editar
    self.janela_editar_formaPag = Toplevel()
    self.janela_editar_formaPag.title = "Editar Forma de Pagamento"
        
    # Frame Edit Forma Pagamento
    frame_editFormaPag = LabelFrame(self.janela_editar_formaPag, text=f' Editar o Forma de Pagamento com ID {formaPag_id} ', font=('Arial', 12))
    frame_editFormaPag.grid(row=1, column=0)
        
    self.antigo = Label(frame_editFormaPag, text='Dados antigos:')
    self.antigo.grid(row=2, column=1)
    self.novo = Label(frame_editFormaPag, text='Dados Novos:')
    self.novo.grid(row=2, column=2)
    
    self.etiqueta_modo= Label(frame_editFormaPag, text='Modo')
    self.etiqueta_modo.grid(row=3, column=0)
    # Modo Pagamento Antigo:
    self.input_modo_antigo = Entry(frame_editFormaPag, textvariable=StringVar(frame_editFormaPag, value=modo), state='readonly')
    self.input_modo_antigo.grid(row=3, column=1)
    # Modo Pagamento Novo:
    self.input_modo_novo = Entry(frame_editFormaPag)
    self.input_modo_novo.grid(row=3, column=2)
     
    
    botao_guardar= ttk.Button(frame_editFormaPag, text='Guardar Alterações', command=self.confirmar_edit_formaPag)
    botao_guardar.grid(row=5, column=0)

# Método Janela de Confirmação da Edição
def confirmar_edit_formaPag(self):
    self.janela_confEditar_formaPag = Toplevel()                        # Criar uma janela à frente da principal
    self.janela_confEditar_formaPag.title = "Editar forma de Pagamento" # Título da janela
    self.janela_confEditar_formaPag.resizable(0,0)                      # Ativar a resimensão da janela
    self.janela_confEditar_formaPag.wm_iconbitmap('recursos/icon.ico')   # Ícon da janela
    
    titulo = LabelFrame(self.janela_confEditar_formaPag, text = 'Editar forma de Pagamento', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)
    
    self.mensagem = ttk.Label(titulo, text='Tem a certeza que deseja editar a forma de Pagamento?', font=('Arial', 14))
    self.mensagem.grid(row=1, column=0)
    
    self.botão_guardar = ttk.Button(titulo, text="Guardar Alterações", command = lambda: self.guardar_edit_formaPag(self.input_modo_novo.get(), self.input_modo_antigo.get()))
    self.botão_guardar.grid(row=2, column=0)
    
    self.botao_voltar = ttk.Button(titulo, text='Voltar atrás', command=self.janela_confEditar_formaPag.destroy)
    self.botao_voltar.grid(row=3, column=0)
    
def guardar_edit_formaPag(self, novo_modo, antigo_modo):
    query='UPDATE formas_pagamento SET Modo=? WHERE Modo=?'
        
    if novo_modo == '':
        novo_modo = antigo_modo

    parametros = (novo_modo, antigo_modo)

    self.db_consulta(query, parametros)
    self.get_formas_pag()
    
    self.janela_confEditar_formaPag.destroy()
    self.janela_formas_pagamento.destroy()
    self.janela_editar_formaPag.destroy()
    self.formas_pagamento() 

#------------------------------------------- ELIMINAR -------------------------------------------
# Método Janela Confirmar Eliminar Forma de Pagamento:
def confirmar_eliminar_formaPag(self):
    self.janela_eliminar_formaPag = Toplevel()                          # Criar uma janela à frente da principal
    self.janela_eliminar_formaPag.title = "Eliminar Forma de Pagamento" # Título da janela
    self.janela_eliminar_formaPag.resizable(0,0)                        # Ativar a resimensão da janela
    self.janela_eliminar_formaPag.wm_iconbitmap('recursos/icon.ico')     # Ícon da janela
        
    titulo = LabelFrame(self.janela_eliminar_formaPag, text = 'Eliminar Forma de Pagamento', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)
        
    self.mensagem = ttk.Label(titulo, text='Tem a certeza que deseja eliminar a Forma de Pagamento?', font=('Arial', 14))
    self.mensagem.grid(row=1, column=0)
        
    self.botão_eliminar = ttk.Button(titulo, text="Eliminar Forma de Pagamento", command=self.eliminar_formaPag)
    self.botão_eliminar.grid(row=2, column=0)
        
    self.botao_voltar = ttk.Button(titulo, text='Voltar atrás', command=self.janela_eliminar_formaPag.destroy)
    self.botao_voltar.grid(row=3, column=0)
    
# Método para Eliminar Forma de Pagamento:
def eliminar_formaPag(self):   
    self.mensagem['text'] = ''
    
    id_formaPag = self.tabela_formaPag.item(self.tabela_formaPag.selection())['values'][0]
    query = 'DELETE FROM formas_pagamento WHERE IDModoPagamento=?'
    
    self.db_consulta(query, (id_formaPag,))
    self.mensagem['text'] = f'A Forma de Pagamento com ID {id_formaPag} foi eliminada com êxito.'
    self.janela_eliminar_formaPag.destroy()
    self.get_formas_pag()

    
#------------------------------------------- EXPORTAR -------------------------------------------
def exportar_formaPag(self):
             
    self.janela_exportar_formaPag = Toplevel()                      # Criar uma janela à frente da principal
    self.janela_exportar_formaPag.title = "Exportar"                # Título da janela
    self.janela_exportar_formaPag.resizable(0,0)                    # Ativar a resimensão da janela
    self.janela_exportar_formaPag.wm_iconbitmap('recursos/icon.ico') # Ícon da janela
        
    titulo = LabelFrame(self.janela_exportar_formaPag, text = 'Exportar dados', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)        
        
    self.mensagem = ttk.Label(titulo, text='Qual formato pretende exportar os dados?')
    self.mensagem.grid(row=3, column=0, columnspan=2, sticky=W+E)
        
    self.botão_excel = ttk.Button(titulo, text="Excel (.xls)", command=self.exp_formaPag_excel)
    self.botão_excel.grid(row=4, column=0)
        
    self.botao_css = ttk.Button(titulo, text='CSV (.csv)', command=self.exp_formaPag_csv)
    self.botao_css.grid(row=4, column=1)
        
    #Consulta dos dados para escrever no ficheiro:
    self.query_formaPag = "SELECT * FROM formas_pagamento"
    self.dados_formaPag = self.db_consulta(self.query_formaPag)
    
def exp_formaPag_excel(self):
    self.janela_exportar_formaPag.destroy()
    
    self.cabecalho = []                     # Para tornar o cabeçalho elegível para Excel
    for item in self.colunas_formaPag:
        self.cabecalho.append(item[0])

    wb = openpyxl.Workbook()                # Criar Excel
    sheet = wb.active
        
    sheet.append(self.cabecalho)            # Introduzir nomes das colunas
    for formaPag in self.dados_formaPag:    # Introduzir os dados
        sheet.append(formaPag)
    wb.save("Relatório das Formas de Pagamento.xlsx")   # Salvar a folha num ficheiro
        
    wb.close()      # Fechar ficheiro
        
    showinfo("", "Ficheiro guardado com sucesso")       # Mensagem de ficheiro guardado
        
def exp_formaPag_csv(self):
    self.janela_exportar_formaPag.destroy()
        
    ficheiro_csv = open("Relatório das Formas de Pagamento.csv", "w")  # Criar CSV
        
    writer = csv.writer(ficheiro_csv)
    writer.writerow(self.colunas_formaPag)  # Introduzir nomes das colunas
    writer.writerows(self.dados_formaPag)   # Introduzir os dados
        
    del writer              # excluir objectos
    ficheiro_csv.close()    # Fechar ficheiro
        
    showinfo("", "Ficheiro guardado com sucesso")   # Mensagem de ficheiro guardado
