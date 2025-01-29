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

# Método Janela Reservas
def reservas(self):
    self.janela_reservas = Toplevel()       # Criat uma janela à frente da principal
    self.janela_reservas.title = "Reservas" # Título da janela
    self.janela_reservas.resizable(1,1)     # Ativar a resimensão da janela
    self.janela_reservas.iconbitmap('recursos/icon.ico')  # Ícon da janela
        
    titulo = LabelFrame(self.janela_reservas, text = 'Reservas', font=('Arial', 14, 'bold'))
    titulo.grid(column=0, row=0)

    self.colunas_reservas = (("ID da\nReserva", 75), ("ID do\nCliente", 70), ("ID do\nVeículo", 70), ("Data\nInício", 120), ("Data\nFim", 120), ("Data\nReserva", 120), ("Valor\nFinal", 120), ("Pago", 90))

    #botão Registar Novo
    self.botão_registar_novo = ttk.Button(titulo, text="Registar Nova Reserva", command=self.nova_reserva)
    self.botão_registar_novo.grid(row=2, column=0, sticky=W+E) 
    #botão Editar/Alterar
    botao_editar = ttk.Button(titulo, text='Editar', command= self.editar_reservas)
    botao_editar.grid(row=2, column=1, sticky=W+E)
    #botão Remover
    botao_eliminar = ttk.Button(titulo, text = 'Eliminar', command= self.confirmar_eliminar_reserva)
    botao_eliminar.grid(row=2, column=2, sticky=W+E)
    #botão Exportar Infos
    self.botão_exp_infos = ttk.Button(titulo, text="Exportar Informações", command=self.exportar_reservas)
    self.botão_exp_infos.grid(row=2, column=3, sticky=W+E)



        
    # Mensagem informativa para o utilizador:
    self.mensagem = Label(self.janela_reservas, text = '', fg = 'green')
    self.mensagem.grid(row=4, column=0, columnspan=3, sticky=W + E)        
        
    # Criação Tabela Reservas
    style = ttk.Style()
    style.configure("mystyle.Treeview", highlightthikness=0, bd=0, font=('Arial', 11))
    style.configure("mystyle.Treeview.Heading", font=('Arial', 12, 'bold'))
    style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky':'nswe'})])        # Eliminar as bordas
        
    self.tabela_reserva = ttk.Treeview(self.janela_reservas, height=15, columns=[x[0] for x in self.colunas_reservas], style="mystyle.Treeview", show="headings")
    self.tabela_reserva.grid(row=5, column=0, columnspan=2)
    self.tabela_reserva.heading('#1', text='ID\nReserva', anchor=CENTER)
    self.tabela_reserva.heading('#2', text='ID\nCliente', anchor=CENTER)
    self.tabela_reserva.heading('#3', text='ID\nVeículo',anchor=CENTER)
    self.tabela_reserva.heading('#4', text='Data Início', anchor=CENTER)
    self.tabela_reserva.heading('#5', text='Data Fim', anchor=CENTER)       
    self.tabela_reserva.heading('#6', text='Data da\nReserva', anchor=CENTER)
    self.tabela_reserva.heading('#7', text='Valor Final', anchor=CENTER)
    self.tabela_reserva.heading('#8', text='Pago', anchor=CENTER)
    
    for col, width in self.colunas_reservas:
        self.tabela_reserva.heading(col, text=col)
        self.tabela_reserva.column(col, width=width, anchor=tk.CENTER)
        
    #Scrollbar:        
    scrollbar = ttk.Scrollbar(self.janela_reservas, orient="vertical", command=self.tabela_reserva.yview)
    self.tabela_reserva.configure(yscroll=scrollbar.set)
    # Posição scrollbar:
    scrollbar.grid(row=5, column=1, sticky="ns")
    self.janela_reservas.grid_rowconfigure(0, weight=1)
    self.janela_reservas.grid_columnconfigure(0, weight=1)
    
    #opcao_reservas = self.combobox_reservas.get()                                             # AINDA N FUNCIONA

    self.get_reservas()
    
def get_reservas(self):
    # 1 - Ao iniciar a app vamos LIMPAR a tabela se tiver dados residuais ou antigos
    registos_tabela = self.tabela_reserva.get_children()            # Obter todos os dados da tabela
    for linha in registos_tabela:
        self.tabela_reserva.delete(linha)
            
    # 2 - Consultar SQL        
    query = 'SELECT * FROM reservas ORDER BY Início DESC'

    registos_db = self.db_consulta(query)  # Faz-se a chamada ao método db_consultas

    # 3 - ESCREVER dados no ecrã:
    for linha in registos_db:
        self.tabela_reserva.insert('', 'end', values=(linha[0], linha[1], linha[2], linha[3], linha[4], linha[5], linha[6], linha[7]))
    
#--------------------------------------------- NOVO ---------------------------------------------
# Método Janela Nova Reserva:
def nova_reserva(self):
    self.janela_nova_reserva = Toplevel()       # Criat uma janela à frente da principal
    self.janela_nova_reserva.title = "Adicionar Reserva" # Título da janela
    self.janela_nova_reserva.resizable(1,1)     # Ativar a resimensão da janela
    
    frame_nova_reserva = LabelFrame(self.janela_nova_reserva, text = 'Adicionar Reserva', font=('Arial', 14, 'bold'))
    frame_nova_reserva.grid(column=0, row=0)
    
    # Label ID Cliente
    self.etiqueta_idCliete = Label(frame_nova_reserva, text='ID do Cliente')
    self.etiqueta_idCliete.grid(row=1, column=0)
    self.idCliete=Entry(frame_nova_reserva)
    self.idCliete.grid(row=1, column=1)
    
    # Labem ID Veículo
    self.etiqueta_idVeiculo = Label(frame_nova_reserva, text='ID do Veículo')
    self.etiqueta_idVeiculo.grid(row=1, column=3)
    self.idVeiculo=Entry(frame_nova_reserva)
    self.idVeiculo.grid(row=1, column=4)        
    
    # Label Data de início
    self.etiqueta_dtInicio = Label(frame_nova_reserva, text='Data de início')
    self.etiqueta_dtInicio.grid(row=2, column=0)
    self.dtInicio=DateEntry(frame_nova_reserva, locale='pt_PT', date_pattern='dd/mm/y')   # Label com data Seleccionada
    self.dtInicio.grid(row=2, column=1)
    
    # Label Data de Fim
    self.etiqueta_dtFim = Label(frame_nova_reserva, text='Data de Fim')
    self.etiqueta_dtFim.grid(row=2, column=3)
    self.dtFim=DateEntry(frame_nova_reserva, locale='pt_PT', date_pattern='dd/mm/y')
    self.dtFim.grid(row=2, column=4)

    # Label Data Reserva
    self.etiqueta_dtReserva = Label(frame_nova_reserva, text='Data Reserva')
    self.etiqueta_dtReserva.grid(row=3, column=0)
    self.dtReserva=DateEntry(frame_nova_reserva, locale='pt_PT', date_pattern='dd/mm/y')
    self.dtReserva.grid(row=3, column=1)

    # Label Valor Final
    self.etiqueta_valorFinal = Label(frame_nova_reserva, text='Valor Final')
    self.etiqueta_valorFinal.grid(row=4, column=0)
    self.valorFinal=Entry(frame_nova_reserva)
    self.valorFinal.grid(row=4, column=1)
    
    # Label Pago
    self.etiqueta_pago = Label(frame_nova_reserva, text='Pago')
    self.etiqueta_pago.grid(row=4, column=3,)
    self.pago=Entry(frame_nova_reserva)
    self.pago.grid(row=4, column=4)

    # Botão Guardar:
    self.botao_add_reserva = ttk.Button(frame_nova_reserva, text='Adicionar Reserva', command=self.add_reserva)
    self.botao_add_reserva.grid(row=5, columnspan=5, sticky=W+E)

# Método para guardar Nova Resera
def add_reserva(self):
    parametros = (self.idCliete.get(), self.idVeiculo.get(), self.dtInicio.get(), self.dtFim.get(), self.dtReserva.get(), self.valorFinal.get(), self.pago.get())
    if self.validacao_dados:
        query = 'INSERT INTO reservas VALUES(NULL, ?, ?, ?, ?, ?, ?, ?)'
        self.db_consulta(query, parametros)
        self.janela_nova_reserva.destroy()
        self.janela_reservas.destroy()
        self.reservas()
        #Para actualizar o dashboard:
        self.reservas_mes()
        self.total_financeiro()
        self.veic_alugados()
        self.veic_disponiveis()
    
    else:
        showinfo("AVISO","O Cliente não foi adicionado, pois faltam dados obrigatórios.")
            
#------------------------------------------- ELIMINAR -------------------------------------------
# Método Janela Confirmar Eliminar Reserva:
def confirmar_eliminar_reserva(self):
    self.mensagem['text'] = ''
    try:
        self.tabela_reserva.item(self.tabela_reserva.selection()[0])
    except IndexError as e:
        self.mensagem['text']='Por favor, seleccione um cliente'
        return
    
    self.janela_eliminar_reserva = Toplevel()                   # Criat uma janela à frente da principal
    self.janela_eliminar_reserva.title = "Eliminar Reserva"  # Título da janela
    self.janela_eliminar_reserva.resizable(0,0)                 # Ativar a resimensão da janela
    self.janela_eliminar_reserva.wm_iconbitmap('recursos/icon.ico')  # Ícon da janela
        
    titulo = LabelFrame(self.janela_eliminar_reserva, text = 'Eliminar Reserva', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)
        
    self.mensagem = ttk.Label(titulo, text='Tem a certeza que deseja eliminar a Reserva?', font=('Arial', 14))
    self.mensagem.grid(row=1, column=0)
    
    self.botão_eliminar = ttk.Button(titulo, text="Eliminar Reserva", command=self.eliminar_reserva)
    self.botão_eliminar.grid(row=2, column=0)
    
    self.botao_voltar = ttk.Button(titulo, text='Voltar atrás', command=self.janela_eliminar_reserva.destroy)
    self.botao_voltar.grid(row=3, column=0)

# Método para Eliminar Reserva:
def eliminar_reserva(self):
             
    self.mensagem['text'] = ''
    id_reserva = self.tabela_reserva.item(self.tabela_reserva.selection())['values'][0]
                             
    query = 'DELETE FROM reservas WHERE IDReserva = ?'
    self.db_consulta(query, (id_reserva,))
    self.mensagem['text'] = f'A reserva de {id_reserva} foi eliminado com êxito.'
    self.janela_eliminar_reserva.destroy()
    self.get_reservas()
      
#-------------------------------------------- EDITAR -------------------------------------------- 
# Método Janela de Edição de Reservas:
def editar_reservas(self):
    self.mensagem['text'] = ''
    try:
        self.tabela_reserva.item(self.tabela_reserva.selection()[0])
    except IndexError as e:
        self.mensagem['text']='Por favor, seleccione um cliente'
        return
        
    #Guardar os dados antigos em listas:
    id_reserva = self.tabela_reserva.item(self.tabela_reserva.selection())['values'][0]
    id_cliente = self.tabela_reserva.item(self.tabela_reserva.selection())['values'][1]
    id_veiculo = self.tabela_reserva.item(self.tabela_reserva.selection())['values'][2]
    inicio = self.tabela_reserva.item(self.tabela_reserva.selection())['values'][3]
    fim = self.tabela_reserva.item(self.tabela_reserva.selection())['values'][4]
    dataReserva = self.tabela_reserva.item(self.tabela_reserva.selection())['values'][5]
    valorFinal = self.tabela_reserva.item(self.tabela_reserva.selection())['values'][6]
    pago = self.tabela_reserva.item(self.tabela_reserva.selection())['values'][7]
    
    self.janela_editar_reserva = Toplevel()
    self.janela_editar_reserva.title = "Editar reserva"
        
    # Frame Edit Cliente
    frame_editReserva = LabelFrame(self.janela_editar_reserva, text=f'Editar reserva com ID {id_reserva} ', font=('Arial', 12))
    frame_editReserva.grid(row=1, column=0)
        
    self.antigoEsq = Label(frame_editReserva, text='Dandos antigo:')
    self.antigoEsq.grid(row=2, column=1)
    self.novoEsq = Label(frame_editReserva, text='Dados Novos:')
    self.novoEsq.grid(row=2, column=2)
    
    self.antigoDir = Label(frame_editReserva, text='Dandos antigo:')
    self.antigoDir.grid(row=2, column=5)
    self.novoDir = Label(frame_editReserva, text='Dados Novos:')
    self.novoDir.grid(row=2, column=6)

    self.etiqueta_idCliente= Label(frame_editReserva, text='ID do Cliente:')
    self.etiqueta_idCliente.grid(row=3, column=0)
    # idCliente Antigo:
    self.input_idCliente_antigo = Entry(frame_editReserva, textvariable=StringVar(frame_editReserva, value=id_cliente), state='readonly')
    self.input_idCliente_antigo.grid(row=3, column=1)
    # idCliente Novo:
    self.input_idCliente_novo = Entry(frame_editReserva)
    self.input_idCliente_novo.grid(row=3, column=2)
    
    self.etiqueta_idVeiculo = Label(frame_editReserva, text='ID do Veículo:')
    self.etiqueta_idVeiculo.grid(row=3, column=4)
    # id_veiculo Antigo:
    self.input_idVeiculo_antigo = Entry(frame_editReserva, textvariable=StringVar(frame_editReserva, value=id_veiculo), state='readonly')
    self.input_idVeiculo_antigo.grid(row=3, column=5)
    # id_veiculo Novo:
    self.input_idVeiculo_novo = Entry(frame_editReserva)
    self.input_idVeiculo_novo.grid(row=3, column=6)
    
    # Data de Inicio:
    self.etiqueta_dataInicio= Label(frame_editReserva, text='Data de Inicio:')
    self.etiqueta_dataInicio.grid(row=4, column=0)
    # Data de Inicio Antiga:
    self.input_dataInicio_antiga = Entry(frame_editReserva, textvariable=StringVar(frame_editReserva, value=inicio), state='readonly')
    self.input_dataInicio_antiga.grid(row=4, column=1)
    # Data de Inicio Nova:
    # Calendário para facilitar o preenchimento:
    dataInicio_antiga = datetime(int(inicio[6:]), int(inicio[3:5]), int(inicio[:2]))
    self.input_dataInicio_nova=DateEntry(frame_editReserva, locale='pt_PT', date_pattern='dd/mm/y', year=dataInicio_antiga.year, month=dataInicio_antiga.month, day=dataInicio_antiga.day)
    self.input_dataInicio_nova.grid(row=4, column=2)

    # Data de Fim:
    self.etiqueta_dataFim= Label(frame_editReserva, text='Data de Fim:')
    self.etiqueta_dataFim.grid(row=4, column=4)
    # Data de Fim Antiga:
    self.input_dataFim_antiga = Entry(frame_editReserva, textvariable=StringVar(frame_editReserva, value=fim), state='readonly')
    self.input_dataFim_antiga.grid(row=4, column=5)
    # Calendário para facilitar o preenchimento:
    dataFim_antiga = datetime(int(fim[6:]), int(fim[3:5]), int(fim[:2]))
    self.input_dataFim_nova=DateEntry(frame_editReserva, locale='pt_PT', date_pattern='dd/mm/y', year=dataFim_antiga.year, month=dataFim_antiga.month, day=dataFim_antiga.day)
    self.input_dataFim_nova.grid(row=4, column=6)

    # Data da Reserva:
    self.etiqueta_dataReserva= Label(frame_editReserva, text='Data da Reserva:')
    self.etiqueta_dataReserva.grid(row=5, column=0)
    # Data da Reserva Antiga:
    self.input_dataReserva_antiga = Entry(frame_editReserva, textvariable=StringVar(frame_editReserva, value=dataReserva), state='readonly')
    self.input_dataReserva_antiga.grid(row=5, column=1)
    # Calendário para facilitar o preenchimento:
    data_reserva = datetime(int(dataReserva[6:]), int(dataReserva[3:5]), int(dataReserva[:2]))
    self.input_dataReserva_nova=DateEntry(frame_editReserva, locale='pt_PT', date_pattern='dd/mm/y', year=data_reserva.year, month=data_reserva.month, day=data_reserva.day)
    self.input_dataReserva_nova.grid(row=5, column=2)
    
    # Valor Final
    self.etiqueta_valorFinal= Label(frame_editReserva, text='Valor Final:')
    self.etiqueta_valorFinal.grid(row=6, column=0)
    # Valor Final Antigo:
    self.input_valorFinal_antigo = Entry(frame_editReserva, textvariable=StringVar(frame_editReserva, value=valorFinal), state='readonly')
    self.input_valorFinal_antigo.grid(row=6, column=1)
    # Valor Final Novo:
    self.input_valorFinal_novo = Entry(frame_editReserva)
    self.input_valorFinal_novo.grid(row=6, column=2)
    
    # Pago
    self.etiqueta_pago= Label(frame_editReserva, text='Pago:')
    self.etiqueta_pago.grid(row=6, column=4)
    # Pago Antigo:
    self.input_pago_antigo = Entry(frame_editReserva, textvariable=StringVar(frame_editReserva, value=pago), state='readonly')
    self.input_pago_antigo.grid(row=6, column=5)
    # Pago Novo:
    self.input_pago_novo = Entry(frame_editReserva)
    self.input_pago_novo.grid(row=6, column=6)
    
    
    botao_guardar= ttk.Button(frame_editReserva, text='Guardar Alterações', command=self.confirmar_edit_reservas)
    botao_guardar.grid(row=8, column=0)
    
# Método Janela de Confirmação da Edição de Clientes:
def confirmar_edit_reservas(self):
    self.janela_confEditar_reserva = Toplevel()                   # Criat uma janela à frente da principal
    self.janela_confEditar_reserva.title = "Editar Reserva"  # Título da janela
    self.janela_confEditar_reserva.resizable(0,0)                 # Ativar a resimensão da janela
    self.janela_confEditar_reserva.wm_iconbitmap('recursos/icon.ico')  # Ícon da janela
    
    titulo = LabelFrame(self.janela_confEditar_reserva, text = 'Editar Reserva', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)
    
    self.mensagem = ttk.Label(titulo, text='Tem a certeza que deseja editar a Reserva?', font=('Arial', 14))
    self.mensagem.grid(row=1, column=0)
    
    self.botão_guardar = ttk.Button(titulo, text="Guardar Alterações", command = lambda: self.guardar_edit_reservas(self.input_idCliente_novo.get(), self.input_idVeiculo_novo.get(), self.input_dataInicio_nova.get(), self.input_dataFim_nova.get(), self.input_dataReserva_nova.get(), self.input_valorFinal_novo.get(), self.input_pago_novo.get(), self.input_idCliente_antigo.get(), self.input_idVeiculo_antigo.get(), self.input_dataInicio_antiga.get(), self.input_dataFim_antiga.get(), self.input_dataReserva_antiga.get(), self.input_valorFinal_antigo.get(), self.input_pago_antigo.get()))
    self.botão_guardar.grid(row=2, column=0)
    
    self.botao_voltar = ttk.Button(titulo, text='Voltar atrás', command=self.janela_confEditar_reserva.destroy)
    self.botao_voltar.grid(row=3, column=0)
    
# Método Janela de Guardar a Edição de Clientes:
def guardar_edit_reservas(self, novo_idCliente, novo_idVeiculo, nova_dataInicio, nova_dataFim, nova_dataReserva, novo_valorFinal, novo_pago, antigo_idCliente, antigo_idVeiculo, antiga_dataInicio, antiga_dataFim, antiga_dataReserva, antigo_valorFinal, antigo_pago):

    #novo = novo_idCliente, novo_idVeiculo, nova_dataInicio, nova_dataFim, nova_dataReserva, novo_valorFinal, novo_pago, antigo_idCliente, antigo_idVeiculo, antiga_dataInicio, antiga_dataFim, antiga_dataReserva, antigo_valorFinal, antigo_pago
    #antigo =antigo_idCliente, antigo_idVeiculo, antiga_dataInicio, antiga_dataFim, antiga_dataReserva, antigo_valorFinal, antigo_pago

    query='''UPDATE reservas SET IDCliente=?, IDVeiculo=?, Início=?, Fim=?, Data_Reserva=?, Valor_Final=?, Pago=? WHERE IDCliente=? AND IDVeiculo=? AND Início=? AND Fim=? AND Data_Reserva=? AND Valor_Final=? AND Pago=?'''
        
    if novo_idCliente == '':
        novo_idCliente = antigo_idCliente
    if novo_idVeiculo == '':
        novo_idVeiculo = antigo_idVeiculo
    if nova_dataInicio == '':
        nova_dataInicio = antiga_dataInicio
    if nova_dataFim == '':
        nova_dataFim = antiga_dataFim
    if nova_dataReserva == '':
        nova_dataReserva = antiga_dataReserva
    if novo_valorFinal == '':
        novo_valorFinal = antigo_valorFinal
    if novo_pago == '':
        novo_pago = antigo_pago
        
    parametros = (novo_idCliente, novo_idVeiculo, nova_dataInicio, nova_dataFim, nova_dataReserva, novo_valorFinal, novo_pago, antigo_idCliente, antigo_idVeiculo, antiga_dataInicio, antiga_dataFim, antiga_dataReserva, antigo_valorFinal, antigo_pago)
    self.db_consulta(query, parametros)
    self.get_reservas()
    
    self.janela_editar_reserva.destroy()
    self.janela_confEditar_reserva.destroy()
    self.janela_reservas.destroy()
    self.reservas()
    #Para actualizar o dashboard:
    self.reservas_mes()
    self.total_financeiro()
    self.veic_alugados()
    self.veic_disponiveis()

#------------------------------------------- EXPORTAR -------------------------------------------
def exportar_reservas(self):
             
    self.janela_exportar_reservas = Toplevel()              # Criat uma janela à frente da principal
    self.janela_exportar_reservas.title = "Exportar"        # Título da janela
    self.janela_exportar_reservas.resizable(0,0)            # Ativar a resimensão da janela
    self.janela_exportar_reservas.wm_iconbitmap('recursos/icon.ico') # Ícon da janela
    
    titulo = LabelFrame(self.janela_exportar_reservas, text = 'Exportar dados', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)        
    
    self.mensagem = ttk.Label(titulo, text='Qual formato pretende exportar os dados?')
    self.mensagem.grid(row=3, column=0, columnspan=2, sticky=W+E)
    
    self.botão_excel = ttk.Button(titulo, text="Excel (.xls)", command=self.exp_reservas_excel)
    self.botão_excel.grid(row=4, column=0)
    
    self.botao_css = ttk.Button(titulo, text='CSV (.csv)', command=self.exp_reservas_csv)
    self.botao_css.grid(row=4, column=1)
    
    #Consulta dos dados para escrever no ficheiro:
    self.query_reservas = "SELECT * FROM reservas"
    self.dados_reservas = self.db_consulta(self.query_reservas)
    
def exp_reservas_excel(self):
    self.janela_exportar_reservas.destroy()

    self.cabecalho = []                     # Para tornar o cabeçalho elegível para Excel
    for item in self.colunas_reservas:
        self.cabecalho.append(item[0])

    wb = openpyxl.Workbook()                # Criar Excel
    sheet = wb.active
        
    sheet.append(self.cabecalho)         # Introduzir nomes das colunas
    for reserva in self.dados_reservas:     # Introduzir os dados
        sheet.append(reserva)
    wb.save("Relatório das reservas.xlsx")  # Salvar a folha num ficheiro
        
    wb.close()      # Fechar ficheiro
        
    showinfo("", "Ficheiro guardado com sucesso")   # Mensagem de ficheiro guardado
        
def exp_reservas_csv(self):
    self.janela_exportar_reservas.destroy()
        
    ficheiro_csv = open("Relatório das reservas.csv", "w")  # Criar CSV
        
    writer = csv.writer(ficheiro_csv)
    writer.writerow(self.colunas_reservas)          # Introduzir nomes das colunas
    writer.writerows(self.dados_reservas)       # Introduzir os dados
        
    del writer              # excluir objectos
    ficheiro_csv.close()    # Fechar ficheirod
        
    showinfo("", "Ficheiro guardado com sucesso")   # Mensagem de ficheiro guardado
    