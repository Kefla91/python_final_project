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

# Método Janela Clientes
def clientes(self):
    self.janela_clientes = Toplevel()       # Criat uma janela à frente da principal
    self.janela_clientes.title = "Clientes" # Título da janela
    self.janela_clientes.resizable(1,1)     # Ativar a resimensão da janela
    self.janela_clientes.wm_iconbitmap('recursos/icon.ico')  # Ícon da janela
        
    titulo = LabelFrame(self.janela_clientes, text = 'Clientes', font=('Arial', 14, 'bold'))
    titulo.grid(column=0, row=0)
        
    self.colunas_clientes= (("ID\nCliente", 60), ("Nome", 120), ("Pronome", 120), ("Data de\nNascimento", 120), ("Telemóvel", 120), ("E-mail", 180))

    #botão Registar Novo
    self.botão_registar_novo = ttk.Button(titulo, text="Novo Cliente", command=self.novo_cliente)
    self.botão_registar_novo.grid(row=2, column=0, sticky=W+E)
    #botão Editar/Alterar
    botao_editar = ttk.Button(titulo, text='Editar', command=self.editar_clientes)
    botao_editar.grid(row=2, column=1, sticky=W+E)
    #botão Remover
    botao_eliminar = ttk.Button(titulo, text = 'Eliminar', command= self.confirmar_eliminar_cliente)
    botao_eliminar.grid(row=2, column=2, sticky=W+E)
    #botão Exportar Infos
    self.botão_exp_infos = ttk.Button(titulo, text="Exportar Informações", command=self.exportar_clientes)
    self.botão_exp_infos.grid(row=2, column=3, sticky=W+E)
                
    # Mensagem informativa para o utilizador:
    self.mensagem = Label(self.janela_clientes, text = '', fg = 'green')
    self.mensagem.grid(row=4, column=0, columnspan=3, sticky=W + E)        
        
    style = ttk.Style()
    style.configure("mystyle.Treeview", highlightthikness=0, bd=0, font=('Arial', 11))
    style.configure("mystyle.Treeview.Heading", font=('Arial', 12, 'bold'))
    style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky':'nswe'})])        # Eliminar as bordas
        
    self.tabela_cliente= ttk.Treeview(self.janela_clientes, height=15, columns=[x[0] for x in self.colunas_clientes], style="mystyle.Treeview", show="headings")
    self.tabela_cliente.grid(row=5, column=0, columnspan=1)
        
    self.tabela_cliente.heading('#1', text='ID\nCliente', anchor=CENTER)
    self.tabela_cliente.heading('#2', text='Nome', anchor=CENTER)
    self.tabela_cliente.heading('#3', text='Sobrenome', anchor=CENTER)
    self.tabela_cliente.heading('#4', text='Data de Nascimento', anchor=CENTER)
    self.tabela_cliente.heading('#5', text='Telemóvel', anchor=CENTER)
    self.tabela_cliente.heading('#6', text='E-mail', anchor=CENTER)
    
    for col, width in self.colunas_clientes:
        self.tabela_cliente.heading(col, text=col)
        self.tabela_cliente.column(col, width=width, anchor=tk.CENTER)
        
    #Scrollbar:        
    scrollbar = ttk.Scrollbar(self.janela_clientes, orient="vertical", command=self.tabela_cliente.yview)

    # Posição scrollbar:
    scrollbar.grid(row=5, column=1, sticky="ns")
    self.janela_clientes.grid_rowconfigure(0, weight=1)
    self.janela_clientes.grid_columnconfigure(0, weight=1)
        
    #opcao_clientes = self.combobox_clientes.get()                                             # AINDA N FUNCIONA

    self.get_clientes()
    
def get_clientes(self):

    # 1 - Ao iniciar a app vamos LIMPAR a tabela se tiver dados residuais ou antigos
    registos_tabela = self.tabela_cliente.get_children()            # Obter todos os dados da tabela
    for linha in registos_tabela:
        self.tabela_cliente.delete(linha)
            
    # 2 - Consultar SQL        
    query = 'SELECT * FROM clientes ORDER BY Nome DESC'

    registos_db = self.db_consulta(query)  # Faz-se a chamada ao método db_consultas
        
    # 3 - ESCREVER dados no ecrã:
    for linha in registos_db:
        self.tabela_cliente.insert('', 'end', values=(linha[0], linha[1], linha[2], linha[3], linha[4], linha[5]))
    
#--------------------------------------------- NOVO --------------------------------------------- 
# Método Janela Novo Cliente:
def novo_cliente(self):
    self.janela_novo_cliente = Toplevel()       # Criat uma janela à frente da principal
    self.janela_novo_cliente.title = "Adicionar Veículo" # Título da janela
    self.janela_novo_cliente.resizable(1,1)     # Ativar a resimensão da janela
        
    frame_novo_cliente = LabelFrame(self.janela_novo_cliente, text = 'Adicionar Cliente', font=('Arial', 14, 'bold'))
    frame_novo_cliente.grid(column=0, row=0)
      
    # Label Nome tabela
    self.etiqueta_nome = Label(frame_novo_cliente, text='Nome')
    self.etiqueta_nome.grid(row=1, column=0)
    self.nome=Entry(frame_novo_cliente)
    self.nome.grid(row=1, column=1)
        
    # Label Sobrenome
    self.etiqueta_sobrenome = Label(frame_novo_cliente, text='Sobrenome')
    self.etiqueta_sobrenome.grid(row=1, column=3)
    self.sobrenome=Entry(frame_novo_cliente)
    self.sobrenome.grid(row=1, column=4)
    
    # Data Nascimento:
    self.etiqueta_dataNascimento = Label(frame_novo_cliente, text='Data de Nascimento')
    self.etiqueta_dataNascimento.grid(row=2, column=0)
    self.calentadario_dataNascimento=DateEntry(frame_novo_cliente, locale='pt_PT', date_pattern='dd/mm/y')   # Label com data Seleccionada
    self.calentadario_dataNascimento.grid(row=2, column=1, sticky=W+E)    
       
    # Label Telemovel
    self.etiqueta_telemovel = Label(frame_novo_cliente, text='Telemóvel')
    self.etiqueta_telemovel.grid(row=3, column=0)
    self.telemovel=Entry(frame_novo_cliente)
    self.telemovel.grid(row=3, column=1)
        
    # Label e-mail
    self.etiqueta_email = Label(frame_novo_cliente, text='E-mail')
    self.etiqueta_email.grid(row=3, column=3)
    self.email=Entry(frame_novo_cliente)
    self.email.grid(row=3, column=4)
        
    # Botão Guardar:
    self.botao_add_cliente = ttk.Button(frame_novo_cliente, text='Adicionar Cliente', command=self.add_cliente)
    self.botao_add_cliente.grid(row=4, columnspan=5, sticky=W+E)

# Método para guardar Novo Cliente
def add_cliente(self):
    parametros = (self.nome.get(), self.sobrenome.get(), self.calentadario_dataNascimento.get(), self.telemovel.get(), self.email.get())
    if self.validacao_dados(parametros) == True:
        query = 'INSERT INTO clientes VALUES(NULL, ?, ?, ?, ?, ?)'
        self.db_consulta(query, parametros)
        self.janela_novo_cliente.destroy()
        self.janela_clientes.destroy()
        self.clientes()
        #Para actualizar o dashboard:
        self.clientes_novos()
        
    else:
        showinfo("AVISO","O Cliente não foi adicionado, pois faltam dados obrigatórios.")

#------------------------------------------- ELIMINAR ------------------------------------------- 
# Método Janela Confirmar Eliminar Cliente:
def confirmar_eliminar_cliente(self):
    self.mensagem['text'] = ''
    try:
        self.tabela_cliente.item(self.tabela_cliente.selection()[0])
    except IndexError as e:
        self.mensagem['text']='Por favor, seleccione um cliente'
        return
    
    self.janela_eliminar_cliente = Toplevel()                   # Criat uma janela à frente da principal
    self.janela_eliminar_cliente.title = "Eliminar Cliente"  # Título da janela
    self.janela_eliminar_cliente.resizable(0,0)                 # Ativar a resimensão da janela
    self.janela_eliminar_cliente.wm_iconbitmap('recursos/icon.ico')  # Ícon da janela
        
    titulo = LabelFrame(self.janela_eliminar_cliente, text = 'Eliminar Cliente', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)
    
    self.mensagem = ttk.Label(titulo, text='Tem a certeza que deseja eliminar o Cliente?', font=('Arial', 14))
    self.mensagem.grid(row=1, column=0)
    
    self.botão_eliminar = ttk.Button(titulo, text="Eliminar Cliente", command=self.eliminar_cliente)
    self.botão_eliminar.grid(row=2, column=0)
    
    self.botao_voltar = ttk.Button(titulo, text='Voltar atrás', command=self.janela_eliminar_cliente.destroy)
    self.botao_voltar.grid(row=3, column=0)
    
# Método para Eliminar Cliente:
def eliminar_cliente(self):
    self.mensagem['text'] = ''
    
    id_cliente = self.tabela_cliente.item(self.tabela_cliente.selection())['values'][0]
    query = 'DELETE FROM clientes WHERE IDCliente = ?'
    
    self.db_consulta(query, (id_cliente,))
    self.mensagem['text'] = f'O cliente com ID {id_cliente} foi eliminado com êxito.'
    self.janela_eliminar_cliente.destroy()
    self.get_clientes()

#-------------------------------------------- EDITAR -------------------------------------------- 
# Método Janela de Edição de Clientes:
def editar_clientes(self):
    self.mensagem['text'] = ''
    try:
        self.tabela_cliente.item(self.tabela_cliente.selection()[0])  
    except IndexError as e:
        self.mensagem['text']='Por favor, seleccione um cliente'
        return
        
    #Guardar os dados antigos em listas:
    cliente_id = self.tabela_cliente.item(self.tabela_cliente.selection())['values'][0]
    nome = self.tabela_cliente.item(self.tabela_cliente.selection())['values'][1]
    sobrenome = self.tabela_cliente.item(self.tabela_cliente.selection())['values'][2]
    dataNasc = self.tabela_cliente.item(self.tabela_cliente.selection())['values'][3]
    telemovel = self.tabela_cliente.item(self.tabela_cliente.selection())['values'][4]
    email = self.tabela_cliente.item(self.tabela_cliente.selection())['values'][5]
        
    self.janela_editar_cliente = Toplevel()
    self.janela_editar_cliente.title = "Editar Cliente"
        
    # Frame Edit Cliente
    frame_editCliente = LabelFrame(self.janela_editar_cliente, text=f" Editar Cliente com ID {cliente_id} ", font=('Arial', 12))
    frame_editCliente.grid(row=1, column=0)
        
    self.antigo = Label(frame_editCliente, text='Dandos antigo:')
    self.antigo.grid(row=2, column=1)
    self.novo = Label(frame_editCliente, text='Dados Novos:')
    self.novo.grid(row=2, column=2)

    self.etiqueta_nome= Label(frame_editCliente, text='Nome:')
    self.etiqueta_nome.grid(row=3, column=0)
    # Nome Antigo:
    self.input_nome_antigo = Entry(frame_editCliente, textvariable=StringVar(self.janela_editar_cliente, value=nome), state='readonly')
    self.input_nome_antigo.grid(row=3, column=1)
    # Nome Nova:
    self.input_nome_novo = Entry(frame_editCliente)
    self.input_nome_novo.grid(row=3, column=2)
    
    self.etiqueta_sobrenome = Label(frame_editCliente, text='Sobrenome:')
    self.etiqueta_sobrenome.grid(row=4, column=0)
    # Sobrenome Antigo:
    self.input_sobrenome_antigo = Entry(frame_editCliente, textvariable=StringVar(self.janela_editar_cliente, value=sobrenome), state='readonly')
    self.input_sobrenome_antigo.grid(row=4, column=1)
    # Sobrenome Novo:
    self.input_sobrenome_novo = Entry(frame_editCliente)
    self.input_sobrenome_novo.grid(row=4, column=2)
    
    # Data de Nascimento:
    self.etiqueta_dataNascimento= Label(frame_editCliente, text='Data de Nascimento:')
    self.etiqueta_dataNascimento.grid(row=5, column=0)
    # Data de Nascimento Antiga:
    self.input_dataNascimento_antiga = Entry(frame_editCliente, textvariable=StringVar(self.janela_editar_cliente, value=dataNasc), state='readonly')
    self.input_dataNascimento_antiga.grid(row=5, column=1)
    # Calendário para facilitar o preenchimento:
    data_NascAntiga = datetime(int(dataNasc[6:]), int(dataNasc[3:5]), int(dataNasc[:2]))
    self.input_dataNascimento_nova=DateEntry(frame_editCliente, locale='pt_PT', date_pattern='dd/mm/y', year=data_NascAntiga.year, month=data_NascAntiga.month, day=data_NascAntiga.day)
    self.input_dataNascimento_nova.grid(row=5, column=2)

    # Telemóvel:
    self.etiqueta_telemovel= Label(frame_editCliente, text='Telemóvel:')
    self.etiqueta_telemovel.grid(row=6, column=0)
    # Telemóvel Antigo:
    self.input_telemovel_antigo = Entry(frame_editCliente, textvariable=StringVar(self.janela_editar_cliente, value=telemovel), state='readonly')
    self.input_telemovel_antigo.grid(row=6, column=1)
    # Telemóvel Novo:
    self.input_telemovel_novo = Entry(frame_editCliente)
    self.input_telemovel_novo.grid(row=6, column=2)

    # Email:
    self.etiqueta_email= Label(frame_editCliente, text='E-mail:')
    self.etiqueta_email.grid(row=7, column=0)
    # Email Antigo:
    self.input_email_antigo = Entry(frame_editCliente, textvariable=StringVar(self.janela_editar_cliente, value=email), state='readonly')
    self.input_email_antigo.grid(row=7, column=1)
    # Email Novo:
    self.input_email_novo = Entry(frame_editCliente)
    self.input_email_novo.grid(row=7, column=2)
    
    botao_guardar= ttk.Button(self.janela_editar_cliente, text='Guardar Alterações', command=self.confirmar_edit_cliente)
    botao_guardar.grid(row=9, column=0)
    
# Método Janela de Confirmação da Edição de Clientes:
def confirmar_edit_cliente(self):
    self.janela_confEditar_cliente = Toplevel()                     # Criat uma janela à frente da principal
    self.janela_confEditar_cliente.title = "Editar Cliente"         # Título da janela
    self.janela_confEditar_cliente.resizable(0,0)                   # Ativar a resimensão da janela
    self.janela_confEditar_cliente.iconbitmap('recursos/icon.ico') # Ícon da janela
    
    titulo = LabelFrame(self.janela_confEditar_cliente, text = 'Editar Cliente', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)
    
    self.mensagem = ttk.Label(titulo, text='Tem a certeza que deseja editar o cliente?', font=('Arial', 14))
    self.mensagem.grid(row=1, column=0)

    self.botão_guardar = ttk.Button(titulo, text="Guardar Alterações", command = lambda:self.guardar_edit_cliente(
        self.input_nome_novo.get(), 
        self.input_sobrenome_novo.get(),
        self.input_dataNascimento_nova.get(), 
        self.input_telemovel_novo.get(), 
        self.input_email_novo.get(), 
        self.input_nome_antigo.get(), 
        self.input_sobrenome_antigo.get(), 
        self.input_dataNascimento_antiga.get(), 
        self.input_telemovel_antigo.get(), 
        self.input_email_antigo.get()
        )
                                    )
    self.botão_guardar.grid(row=2, column=0)
    
    self.botao_voltar = ttk.Button(titulo, text='Voltar atrás', command=self.janela_confEditar_cliente.destroy)
    self.botao_voltar.grid(row=3, column=0)
    
# Método Janela de Guardar a Edição de Clientes:
def guardar_edit_cliente(self, novo_nome, novo_sobrenome, nova_dataNascimento, novo_telemovel, novo_email, antigo_nome, antigo_sobrenome, antiga_dataNascimento, antigo_telemovel, antigo_email):
   
    query='UPDATE clientes SET Nome=?, Sobrenome=?, Data_Nascimento=?, Telemóvel=?, Email=? WHERE Nome=? AND Sobrenome=? AND Data_Nascimento=? AND  Telemóvel=? AND Email=?'
        
    if novo_nome == '':
        novo_nome = antigo_nome
    if novo_sobrenome == '':
        novo_sobrenome = antigo_sobrenome
    if nova_dataNascimento == '':
        nova_dataNascimento = antiga_dataNascimento
    if novo_telemovel == '':
        novo_telemovel = antigo_telemovel
    if novo_email == '':
        novo_email = antigo_email
        
    parametros = (novo_nome, novo_sobrenome, nova_dataNascimento, novo_telemovel, novo_email, antigo_nome, antigo_sobrenome, antiga_dataNascimento, antigo_telemovel, antigo_email)

    self.db_consulta(query, parametros)
    self.get_clientes()
        
    self.janela_editar_cliente.destroy()
    self.janela_confEditar_cliente.destroy()
    self.janela_clientes.destroy()
    
    self.clientes()
    #Para actualizar o dashboard:
    self.clientes_novos()

#------------------------------------------- EXPORTAR -------------------------------------------
# Método para exportar
def exportar_clientes(self):
             
    self.janela_exportar_clientes = Toplevel()              # Criat uma janela à frente da principal
    self.janela_exportar_clientes.title = "Exportar"        # Título da janela
    self.janela_exportar_clientes.resizable(0,0)            # Ativar a resimensão da janela
    self.janela_exportar_clientes.wm_iconbitmap('recursos/icon.ico') # Ícon da janela
        
    titulo = LabelFrame(self.janela_exportar_clientes, text = 'Exportar dados', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)        
        
    self.mensagem = ttk.Label(titulo, text='Qual formato pretende exportar os dados?')
    self.mensagem.grid(row=3, column=0, columnspan=2, sticky=W+E)
        
    self.botão_excel = ttk.Button(titulo, text="Excel (.xls)", command=self.exp_clientes_excel)
    self.botão_excel.grid(row=4, column=0)
        
    self.botao_css = ttk.Button(titulo, text='CSV (.csv)', command=self.exp_clientes_csv)
    self.botao_css.grid(row=4, column=1)
        
    #Consulta dos dados para escrever no ficheiro:
    self.query_clientes = "SELECT * FROM clientes"
    self.dados_clientes = self.db_consulta(self.query_clientes)
   
# Método para exportar em Excel 
def exp_clientes_excel(self):
    self.janela_exportar_clientes.destroy()

    self.cabecalho = []                     # Para tornar o cabeçalho elegível para Excel
    for item in self.colunas_clientes:
        self.cabecalho.append(item[0])
        
    wb = openpyxl.Workbook()                # Criar Excel
    sheet = wb.active
        
    sheet.append(self.cabecalho)         # Introduzir nomes das colunas
    for cliente in self.dados_clientes:     # Introduzir os dados
        sheet.append(cliente)
    wb.save("Relatório dos clientes.xlsx")  # Salvar a folha num ficheiro
    
    wb.close()      # Fechar ficheiro
    
    showinfo("", "Ficheiro guardado com sucesso")   # Mensagem de ficheiro guardado
        
# Método para exportar em CSV
def exp_clientes_csv(self):
    self.janela_exportar_clientes.destroy()
        
    ficheiro_csv = open("Relatório dos clientes.csv", "w")  # Criar CSV
        
    writer = csv.writer(ficheiro_csv)
    writer.writerow(self.colunas_clientes)          # Introduzir nomes das colunas
    writer.writerows(self.dados_clientes)       # Introduzir os dados
        
    del writer              # excluir objectos
    ficheiro_csv.close()    # Fechar ficheirod
        
    showinfo("", "Ficheiro guardado com sucesso")   # Mensagem de ficheiro guardado
