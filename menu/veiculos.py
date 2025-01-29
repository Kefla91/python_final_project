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
  
# Método Nova janela Veículos 
def veiculos(self):
    self.janela_veiculos = Toplevel()       # Criat uma janela à frente da principal
    self.janela_veiculos.title = "Veículos" # Título da janela
    self.janela_veiculos.resizable(1,1)     # Ativar a resimensão da janela
    self.janela_veiculos.iconbitmap('recursos/icon.ico')    # Ícon da janela
        
    titulo = LabelFrame(self.janela_veiculos, text = 'Veículos', font=('Arial', 14, 'bold'))
    titulo.grid(column=0, row=0)
        
    self.colunas_veic = [("ID\nVeículo", 60), ("Matrícula", 150), ("Tipo de\nVeículo", 150), ("Marca", 150), ("Modelo", 150), ("Categora", 150), ("Tipo de\nTransmissão", 150), ("Quantidade de\npassageiros", 90), ("Valor diário", 90), ("Data de\nRevisão", 120), ("Data de\nInspeção", 120), ("Disponivel", 90), ("Imagem", 90)]

    #botão Registar Novo
    self.botão_registar_novo = ttk.Button(titulo, text="Novo Veículo", command=self.novo_veiculo)
    self.botão_registar_novo.grid(row=2, column=0, sticky=W+E)    
    #botão Editar/Alterar
    botao_editar = ttk.Button(titulo, text='Editar', command=self.editar_veiculo)
    botao_editar.grid(row=2, column=1, sticky=W+E)  
    #botão Remover
    botao_eliminar = ttk.Button(titulo, text = 'Eliminar', command= self.confirmar_eliminar_veiculos)
    botao_eliminar.grid(row=2, column=2, sticky=W+E)        
    #botão Exportar Infos
    self.botão_exp_infos = ttk.Button(titulo, text="Exportar Informações", command=self.exportar_veiculos)
    self.botão_exp_infos.grid(row=2, column=3, sticky=W+E)

    # Mensagem informativa para o utilizador:
    self.mensagem = Label(self.janela_veiculos, text = '', fg = 'green')
    self.mensagem.grid(row=4, column=0, columnspan=3, sticky=W + E)

    # Criação Tabela Veículos
    style = ttk.Style()
    style.configure("mystyle.Treeview", highlightthikness=0, bd=0, font=('Arial', 11))
    style.configure("mystyle.Treeview.Heading", font=('Arial', 12, 'bold'), padding = (0,20), anchor="center")
    style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky':'nswe'})])        # Eliminar as bordas
        
    self.tabela_veiculos= ttk.Treeview(self.janela_veiculos, height=15, columns=[x[0] for x in self.colunas_veic], style="mystyle.Treeview", show="headings")
    self.tabela_veiculos.grid(row=5, column=0)
    self.tabela_veiculos.heading('#1', text='ID Veículo', anchor=CENTER)
    self.tabela_veiculos.heading('#2', text='Matrícula', anchor=CENTER)
    self.tabela_veiculos.heading('#3', text='Tipo de\nVeículo', anchor=CENTER)
    self.tabela_veiculos.heading('#4', text='Marca', anchor=CENTER)
    self.tabela_veiculos.heading('#5', text='Modelo', anchor=CENTER)
    self.tabela_veiculos.heading('#6', text='Categoria', anchor=CENTER)
    self.tabela_veiculos.heading('#7', text='Tipo de\nTransmissão', anchor=CENTER)
    self.tabela_veiculos.heading('#8', text='Quantidade de\npassageiros', anchor=CENTER)
    self.tabela_veiculos.heading('#9', text='Valor diário', anchor=CENTER)
    self.tabela_veiculos.heading('#10', text='Data de\nRevisão', anchor=CENTER)
    self.tabela_veiculos.heading('#11', text='Data de\nInspeção', anchor=CENTER)
    self.tabela_veiculos.heading('#12', text='Disponivel', anchor=CENTER)
    self.tabela_veiculos.heading('#12', text='Imagem', anchor=CENTER)

    for col, width in self.colunas_veic:
        self.tabela_veiculos.heading(col, text=col)
        self.tabela_veiculos.column(col, width=width, anchor=tk.CENTER)
        
    #Scrollbar:        
    scrollbar = ttk.Scrollbar(self.janela_veiculos, orient="vertical", command=self.tabela_veiculos.yview)

    # Posição scrollbar:
    scrollbar.grid(row=5, column=1, sticky="ns")
    self.janela_veiculos.grid_rowconfigure(0, weight=1)
    self.janela_veiculos.grid_columnconfigure(0, weight=1)

    #opcao = self.combobox_veiculo.get()                                             # AINDA N FUNCIONA
    #print(opcao)
    self.get_veiculos()
    
def get_veiculos(self):
    # 1 - Ao iniciar a app vamos LIMPAR a tabela se tiver dados residuais ou antigos
    registos_tabela = self.tabela_veiculos.get_children()            # Obter todos os dados da tabela
    for linha in registos_tabela:
        self.tabela_veiculos.delete(linha)
        
    # 2 - Consultar SQL        
    query_veic = 'SELECT * FROM veiculos ORDER BY IDVeiculo DESC'
    self.registos_db_veic = self.db_consulta(query_veic)  # Faz-se a chamada ao método db_consultas
    
    # 3 - ESCREVER dados no ecrã:
    for linha in self.registos_db_veic:
        self.tabela_veiculos.insert('', 'end', values=(linha[0], linha[1], linha[2], linha[3], linha[4], linha[5], linha[6], linha[7], linha[8], linha[9], linha[10], linha[11], linha[12]))

#--------------------------------------------- NOVO --------------------------------------------- 
# Método Janela Novo Veívulo:
def novo_veiculo(self):
    self.janela_novo_veiculo = Toplevel()       # Criat uma janela à frente da principal
    self.janela_novo_veiculo.title = "Adicionar Veículo" # Título da janela
    self.janela_novo_veiculo.resizable(1,1)     # Ativar a resimensão da janela
        
    self.novo_veiculo = LabelFrame(self.janela_novo_veiculo, text = 'Adicionar Veículo', font=('Arial', 14, 'bold'))
    self.novo_veiculo.grid(column=0, row=0)
        
    # Label Matrícula
    self.etiqueta_matricula = Label(self.novo_veiculo, text='Matricula')
    self.etiqueta_matricula.grid(row=1, column=0)
    self.matricula=Entry(self.novo_veiculo)
    self.matricula.grid(row=1, column=1)
        
    # Label Tipo de Veículo
    self.etiqueta_tipoVeículo = Label(self.novo_veiculo, text='Tipo de Veículo')
    self.etiqueta_tipoVeículo.grid(row=1, column=3)
    self.tipoVeículo=Entry(self.novo_veiculo)
    self.tipoVeículo.grid(row=1, column=4)
        
    # Label Marca
    self.etiqueta_marca = Label(self.novo_veiculo, text='Marca')
    self.etiqueta_marca.grid(row=2, column=0)
    self.marca=Entry(self.novo_veiculo)
    self.marca.grid(row=2, column=1)
        
    # Label Modelo
    self.etiqueta_modelo = Label(self.novo_veiculo, text='Modelo')
    self.etiqueta_modelo.grid(row=2, column=3)
    self.modelo=Entry(self.novo_veiculo)
    self.modelo.grid(row=2, column=4)
        
    # Label Categoria
    self.etiqueta_categoria = Label(self.novo_veiculo, text='Categoria')
    self.etiqueta_categoria.grid(row=3, column=0)
    self.categoria=Entry(self.novo_veiculo)
    self.categoria.grid(row=3, column=1)
        
    # Label Tipo de Transmissão
    self.etiqueta_tipoTransmissão = Label(self.novo_veiculo, text='Tipo de Transmissão')
    self.etiqueta_tipoTransmissão.grid(row=3, column=3)
    self.tipoTransmissão=Entry(self.novo_veiculo)
    self.tipoTransmissão.grid(row=3, column=4)
        
    # Label Quantidade de Passageiros
    self.etiqueta_qtdPassageiros = Label(self.novo_veiculo, text='Quantidade de Passageiros')
    self.etiqueta_qtdPassageiros.grid(row=4, column=0)
    self.qtdPassageiros=Entry(self.novo_veiculo)
    self.qtdPassageiros.grid(row=4, column=1)
        
    # Label Valor Diário
    self.etiqueta_valorDiario = Label(self.novo_veiculo, text='Valor Diário')
    self.etiqueta_valorDiario.grid(row=4, column=3)
    self.valorDiario=Entry(self.novo_veiculo)
    self.valorDiario.grid(row=4, column=4)
        
    # Calendário Próxima Revisão
    self.etiqueta_revisao = Label(self.novo_veiculo, text='Data da próxima\nRevisão')
    self.etiqueta_revisao.grid(row=5, column=0)
    self.calendario_revisao=DateEntry(self.novo_veiculo, locale='pt_PT', date_pattern='dd/mm/y')   # Label com data Seleccionada
    self.calendario_revisao.grid(row=5, column=1)
        
    # Calendário Próxima Inspeção:
    self.etiqueta_inspecao = Label(self.novo_veiculo, text='Data da próxima\nInspecao')
    self.etiqueta_inspecao.grid(row=5, column=3)
    self.calendario_inspecao=DateEntry(self.novo_veiculo, locale='pt_PT', date_pattern='dd/mm/y')   # Label com data Seleccionada
    self.calendario_inspecao.grid(row=5, column=4)
    
    # Disponibilidade
    self.etiqueta_disponib = Label(self.novo_veiculo, text='Disponíbel?')
    self.etiqueta_disponib.grid(row=6, column=0)
    self.disponib=Entry(self.novo_veiculo)
    self.disponib.grid(row=6, column=1)
    
    # Imagem
    self.etiqueta_imagem = Label(self.novo_veiculo, text='Imagem')
    self.etiqueta_imagem.grid(row=6, column=3)
    self.imagem=Entry(self.novo_veiculo)
    self.imagem.grid(row=6, column=4)
     
    # Botão Guardar:
    self.botao_add_veiculo = ttk.Button(self.novo_veiculo, text='Adicionar Veículo', command=self.add_veiculo)
    self.botao_add_veiculo.grid(row=7, columnspan=5, sticky=W+E)

# Método para Guardar Novo Veículo:
def add_veiculo(self):
    parametros = (self.matricula.get(), self.tipoVeículo.get(), self.marca.get(), self.modelo.get(), self.categoria.get(), self.tipoTransmissão.get(), self.qtdPassageiros.get(), self.valorDiario.get(), self.calendario_revisao.get(), self.calendario_inspecao.get(), self.disponib.get(), self.imagem.get())
    
    if self.validacao_dados(parametros) == True:
        query = 'INSERT INTO veiculos VALUES(NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
        self.db_consulta(query, parametros)
        self.janela_novo_veiculo.destroy()
        self.janela_veiculos.destroy()
        self.veiculos()
        #Para actualizar o dashboard:
        self.revisao()
        self.inspecao()
        self.reservas_mes()
        self.veic_disponiveis()
    else:
        showinfo("AVISO","O Veículo não foi adicionado, pois faltam dados obrigatórios.")


#------------------------------------------- ELIMINAR -------------------------------------------   
# Método Janela Confirmar Eliminar Veículo:
def confirmar_eliminar_veiculos(self):
    self.mensagem['text'] = ''
    try:
        self.tabela_veiculos.item(self.tabela_veiculos.selection()[0])
    except IndexError as e:
        self.mensagem['text']='Por favor, seleccione um veículo'
        return
        
    self.janela_eliminar_veiculos = Toplevel()                   # Criat uma janela à frente da principal
    self.janela_eliminar_veiculos.title = "Eliminar Veículo"  # Título da janela
    self.janela_eliminar_veiculos.resizable(0,0)                 # Ativar a resimensão da janela
    self.janela_eliminar_veiculos.wm_iconbitmap('recursos/icon.ico')  # Ícon da janela
        
    titulo = LabelFrame(self.janela_eliminar_veiculos, text = 'Eliminar Veículo', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)
        
    self.mensagem = ttk.Label(titulo, text='Tem a certeza que deseja eliminar o veículo?', font=('Arial', 14))
    self.mensagem.grid(row=1, column=0)
        
    self.botão_eliminar = ttk.Button(titulo, text="Eliminar Veículo", command=self.eliminar_veiculos)
    self.botão_eliminar.grid(row=2, column=0)
        
    self.botao_voltar = ttk.Button(titulo, text='Voltar atrás', command=self.janela_eliminar_veiculos.destroy)
    self.botao_voltar.grid(row=3, column=0)
    
# Método para Eliminar Veículo:
def eliminar_veiculos(self):       
    self.mensagem['text'] = ''
    
    matricula = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][1]
    query = 'DELETE FROM veiculos WHERE Matrícula=?'
    
    self.db_consulta(query, (matricula,))
    self.mensagem['text'] = f'O veiculo com matricula {matricula} foi eliminado com êxito.'
    self.janela_eliminar_veiculos.destroy()
    self.get_veiculos()
    
#-------------------------------------------- EDITAR -------------------------------------------- 
# Método Janela de Edição de Veículos:
def editar_veiculo(self):
    self.mensagem['text'] = ''
    try:
        self.tabela_veiculos.item(self.tabela_veiculos.selection()[0])
    except IndexError as e:
        self.mensagem['text']='Por favor, seleccione um veículo'
        return
        
    #Guardar os dados antigos em listas:
    veiculo_id = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][0]
    matricula_antiga = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][1]
    tipoVeiculo_antigo = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][2]
    marca_antiga= self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][3]
    modelo_antigo= self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][4]
    categoria_antiga= self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][5]
    tipoTransm_antigo= self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][6]
    qtdPassageiros_antiga= self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][7]
    valorDiario_antigo= self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][8]
    revisao_antigo= self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][9]
    inspecao_antigo= self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][10]
    disponib_antiga= self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][11]
    imagem_antiga = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][12]

    self.janela_editar_veiculo = Toplevel()
    self.janela_editar_veiculo.title = "Editar Veículo"
        
    # Frame Edit Veiculo
    frame_editVeiculo = LabelFrame(self.janela_editar_veiculo, text=f' Editar o veículo com ID {veiculo_id} ', font=('Arial', 12))
    frame_editVeiculo.grid(row=1, column=0)
        
    self.antigoEsq = Label(frame_editVeiculo, text='Dados antigos:')
    self.antigoEsq.grid(row=2, column=1)
    self.novoEsq = Label(frame_editVeiculo, text='Dados Novos:')
    self.novoEsq.grid(row=2, column=2)
    
    self.antigoEsq = Label(frame_editVeiculo, text='Dados antigos:')
    self.antigoEsq.grid(row=2, column=5)
    self.novoEsq = Label(frame_editVeiculo, text='Dados Novos:')
    self.novoEsq.grid(row=2, column=6)
    
    self.etiqueta_matricula= Label(frame_editVeiculo, text='matricula')
    self.etiqueta_matricula.grid(row=3, column=0)
    # Matrícula Antiga:
    self.input_matricula_antiga = Entry(frame_editVeiculo, textvariable=StringVar(frame_editVeiculo, value=matricula_antiga), state='readonly')
    self.input_matricula_antiga.grid(row=3, column=1)
    # Matricula Nova:
    self.input_matricula_nova = Entry(frame_editVeiculo)
    self.input_matricula_nova.grid(row=3, column=2)
    
    self.etiqueta_tipoVeiculo = Label(frame_editVeiculo, text='Tipo de Veículo:')
    self.etiqueta_tipoVeiculo.grid(row=3, column=4)
    # Tipo de Veículo Antigo:
    self.input_tipoVeiculo_antigo = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=tipoVeiculo_antigo), state='readonly')
    self.input_tipoVeiculo_antigo.grid(row=3, column=5)
    # Tipo de Veículo Novo:
    self.input_tipoVeiculo_novo = Entry(frame_editVeiculo)
    self.input_tipoVeiculo_novo.grid(row=3, column=6)
    
    # Marca:
    self.etiqueta_marca= Label(frame_editVeiculo, text='Marca:')
    self.etiqueta_marca.grid(row=4, column=0)
    # Marca Antiga:
    self.input_marca_antiga = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=marca_antiga), state='readonly')
    self.input_marca_antiga.grid(row=4, column=1)
    # Marca Nova:
    self.input_marca_nova = Entry(frame_editVeiculo)
    self.input_marca_nova.grid(row=4, column=2)
    
    # Modelo
    self.etiqueta_modelo = Label(frame_editVeiculo, text='Modelo:')
    self.etiqueta_modelo.grid(row=4, column=4)
    # Modelo Antigo:
    self.input_modelo_antigo = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=modelo_antigo), state='readonly')
    self.input_modelo_antigo.grid(row=4, column=5)
    # Modelo Novo:
    self.input_modelo_novo = Entry(frame_editVeiculo)
    self.input_modelo_novo.grid(row=4, column=6)

    # Categoria:
    self.etiqueta_categoria= Label(frame_editVeiculo, text='Categoria:')
    self.etiqueta_categoria.grid(row=5, column=0)
    # Categoria Antiga:
    self.input_categoria_antiga = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=categoria_antiga), state='readonly')
    self.input_categoria_antiga.grid(row=5, column=1)
    # Categoria Nova:
    self.input_categoria_nova = Entry(frame_editVeiculo)
    self.input_categoria_nova.grid(row=5, column=2)
    
    # Tipo de transmissão
    self.etiqueta_tipoTransm = Label(frame_editVeiculo, text='Tipo de transmissão:')
    self.etiqueta_tipoTransm.grid(row=5, column=4)
    # Tipo de transmissão Antiga:
    self.input_tipoTransm_antigo = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=tipoTransm_antigo), state='readonly')
    self.input_tipoTransm_antigo.grid(row=5, column=5)
    # Tipo de transmissão Nova:
    self.input_tipoTransm_novo = Entry(frame_editVeiculo)
    self.input_tipoTransm_novo.grid(row=5, column=6)

    # Quantidade de passageiros:
    self.etiqueta_qtdPassageiros= Label(frame_editVeiculo, text='Quantidade de passageiros:')
    self.etiqueta_qtdPassageiros.grid(row=6, column=0)
    # Quantidade de passageiros Antiga:
    self.input_qtdPassageiros_antiga = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=qtdPassageiros_antiga), state='readonly')
    self.input_qtdPassageiros_antiga.grid(row=6, column=1)
    # Quantidade de passageiros Nova:
    self.input_qtdPassageiros_nova = Entry(frame_editVeiculo)
    self.input_qtdPassageiros_nova.grid(row=6, column=2)
    
    # Valor Diário
    self.etiqueta_valorDiario = Label(frame_editVeiculo, text='Valor Diário:')
    self.etiqueta_valorDiario.grid(row=6, column=4)
    # Valor Diário Antigo:
    self.input_valorDiario_antigo = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=valorDiario_antigo), state='readonly')
    self.input_valorDiario_antigo.grid(row=6, column=5)
    # Valor Diário Novo:
    self.input_valorDiario_novo = Entry(frame_editVeiculo)
    self.input_valorDiario_novo.grid(row=6, column=6)
    
    # Calendário Próxima Revisão
    self.etiqueta_revisao = Label(frame_editVeiculo, text='Data da próxima\nRevisão')
    self.etiqueta_revisao.grid(row=7, column=0)
    self.input_revisao_antigo = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=revisao_antigo), state='readonly')
    self.input_revisao_antigo.grid(row=7, column=1)
    # Calendário para facilitar o preenchimento:
    data_rev_antiga = datetime(int(revisao_antigo[6:]), int(revisao_antigo[3:5]), int(revisao_antigo[:2]))
    self.input_revisao_nova=DateEntry(frame_editVeiculo, locale='pt_PT', date_pattern='dd/mm/y', year=data_rev_antiga.year, month=data_rev_antiga.month, day=data_rev_antiga.day)
    self.input_revisao_nova.grid(row=7, column=2)
    
    # Calendário Próxima Inspeção:
    self.etiqueta_inspecao = Label(frame_editVeiculo, text='Data da próxima\nInspecao')
    self.etiqueta_inspecao.grid(row=7, column=4)
    self.input_inspecao_antigo = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=inspecao_antigo), state='readonly')
    self.input_inspecao_antigo.grid(row=7, column=5)
    # Calendário para facilitar o preenchimento:
    data_insp_antiga = datetime(int(inspecao_antigo[6:]), int(inspecao_antigo[3:5]), int(inspecao_antigo[:2]))
    self.input_inspecao_nova=DateEntry(frame_editVeiculo, locale='pt_PT', date_pattern='dd/mm/y', year=data_insp_antiga.year, month=data_insp_antiga.month, day=data_insp_antiga.day)
    self.input_inspecao_nova.grid(row=7, column=6)
    
    # Disponibilidade:
    self.etiqueta_disponib = Label(frame_editVeiculo, text='Disponivel')
    self.etiqueta_disponib.grid(row=8, column=0)
    self.input_disponib_antiga = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=disponib_antiga), state='readonly')
    self.input_disponib_antiga.grid(row=8, column=1)
    self.input_disponib_nova=Entry(frame_editVeiculo)
    self.input_disponib_nova.grid(row=8, column=2)   
    
    # Imagem:
    self.etiqueta_imagem = Label(frame_editVeiculo, text='Imagem')
    self.etiqueta_imagem.grid(row=8, column=4)
    self.input_imagem_antiga = Entry(frame_editVeiculo, textvariable=StringVar(self.janela_editar_veiculo, value=imagem_antiga), state='readonly')
    self.input_imagem_antiga.grid(row=8, column=5)
    self.input_imagem_nova=Entry(frame_editVeiculo)
    self.input_imagem_nova.grid(row=8, column=6)      
    
    botao_guardar= ttk.Button(frame_editVeiculo, text='Guardar Alterações', command=self.confirmar_edit_veiculo)
    botao_guardar.grid(row=9, column=0)
    
# Método Janela de Confirmação da Edição de Veículos:
def confirmar_edit_veiculo(self):
    self.janela_confEditar_veiculos = Toplevel()                    # Criat uma janela à frente da principal
    self.janela_confEditar_veiculos.title = "Editar Veículo"        # Título da janela
    self.janela_confEditar_veiculos.resizable(0,0)                  # Ativar a resimensão da janela
    self.janela_confEditar_veiculos.iconbitmap('recursos/icon.ico')  # Ícon da janela
    
    titulo = LabelFrame(self.janela_confEditar_veiculos, text = 'Editar Veículo', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)
    
    self.mensagem = ttk.Label(titulo, text='Tem a certeza que deseja editar o veículo?', font=('Arial', 14))
    self.mensagem.grid(row=1, column=0)
    
    self.botão_guardar = ttk.Button(titulo, text="Guardar Alterações", command = lambda: self.guardar_edit_veiculo(self.input_matricula_nova.get(), self.input_tipoVeiculo_novo.get(), self.input_marca_nova.get(), self.input_modelo_novo.get(), self.input_categoria_nova.get(), self.input_tipoTransm_novo.get(), self.input_qtdPassageiros_nova.get(), self.input_valorDiario_novo.get(), self.input_revisao_nova.get(), self.input_inspecao_nova.get(), self.input_disponib_nova.get(), self.input_imagem_nova.get(),self.input_matricula_antiga.get(), self.input_tipoVeiculo_antigo.get(), self.input_marca_antiga.get(), self.input_modelo_antigo.get(), self.input_categoria_antiga.get(), self.input_tipoTransm_antigo.get(), self.input_qtdPassageiros_antiga.get(), self.input_valorDiario_antigo.get(), self.input_revisao_antigo.get(), self.input_inspecao_antigo.get(), self.input_disponib_antiga.get(), self.input_imagem_antiga.get()))
    self.botão_guardar.grid(row=2, column=0)
    
    self.botao_voltar = ttk.Button(titulo, text='Voltar atrás', command=self.janela_confEditar_veiculos.destroy)
    self.botao_voltar.grid(row=3, column=0)
    
# Método Janela de Guardar a Edição de Veículos:
def guardar_edit_veiculo(self, nova_matricula, novo_tipoVeiculo, nova_marca, novo_modelo, nova_categoria, novo_tipoTransmissao, nova_qtdPassageiros, novo_valorDiario, nova_dataRev, nova_dataInsp, nova_disponib, nova_imagem, antiga_matricula, antigo_tipoVeiculo, antiga_marca, antigo_modelo, antiga_categoria, antigo_tipoTransmissao, antiga_qtdPassageiros, antigo_valorDiario, antiga_dataRev, antiga_dataInsp, antiga_disponib, antiga_imagem):

    query='UPDATE veiculos SET Matrícula=?, Tipo_Veiculo=?, Marca=?, Modelo=?, Categoria=?, Tipo_Transmissao=?, Qtd_Passageiros=?, Valor_Diário=?, Data_Revisao=?, Data_Inspecao=?, Disponivel=?, Imagem=? WHERE Matrícula=? AND Tipo_Veiculo=? AND Marca=? AND Modelo=? AND Categoria=? AND Tipo_Transmissao=? AND Qtd_Passageiros=? AND Valor_Diário=? AND Data_Revisao=? AND Data_Inspecao=? AND Disponivel=? AND Imagem=?'
        
    if nova_matricula == '':
        nova_matricula = antiga_matricula
    if novo_tipoVeiculo == '':
        novo_tipoVeiculo = antigo_tipoVeiculo
    if nova_marca == '':
        nova_marca = antiga_marca
    if novo_modelo == '':
        novo_modelo = antigo_modelo
    if nova_categoria == '':
        nova_categoria = antiga_categoria
    if novo_tipoTransmissao == '':
        novo_tipoTransmissao = antigo_tipoTransmissao
    if nova_qtdPassageiros == '':
        nova_qtdPassageiros = antiga_qtdPassageiros
    if novo_valorDiario == '':
        novo_valorDiario = antigo_valorDiario
    if novo_valorDiario == '':
        novo_valorDiario = antigo_valorDiario
    if nova_dataInsp == '':
        nova_dataInsp = antiga_dataInsp
    if nova_dataRev == '':
        nova_dataRev = antiga_dataRev
    if nova_disponib == '':
        nova_disponib = antiga_disponib            
    if nova_imagem == '':
        nova_imagem = antiga_imagem

    parametros = (nova_matricula, novo_tipoVeiculo, nova_marca, novo_modelo, nova_categoria, novo_tipoTransmissao, nova_qtdPassageiros, novo_valorDiario, nova_dataRev, nova_dataInsp, nova_disponib, nova_imagem, antiga_matricula,antigo_tipoVeiculo, antiga_marca, antigo_modelo, antiga_categoria, antigo_tipoTransmissao, antiga_qtdPassageiros, antigo_valorDiario, antiga_dataRev, antiga_dataInsp, antiga_disponib, antiga_imagem)

    self.db_consulta(query, parametros)
    self.get_veiculos()
        
    self.janela_editar_veiculo.destroy()
    self.janela_confEditar_veiculos.destroy()
    self.janela_veiculos.destroy()
    self.veiculos() 
    #Para actualizar o dashboard:  
    self.revisao()
    self.inspecao()
    self.reservas_mes()
    self.veic_disponiveis()

        
#------------------------------------------- EXPORTAR -------------------------------------------
def exportar_veiculos(self):
             
    self.janela_exportar_veiculos = Toplevel()              # Criat uma janela à frente da principal
    self.janela_exportar_veiculos.title = "Exportar"        # Título da janela
    self.janela_exportar_veiculos.resizable(0,0)            # Ativar a resimensão da janela
    self.janela_exportar_veiculos.wm_iconbitmap('recursos/icon.ico') # Ícon da janela
        
    titulo = LabelFrame(self.janela_exportar_veiculos, text = 'Exportar dados', font=('Arial', 8, 'bold'))
    titulo.grid(column=0, row=0)        
        
    self.mensagem = ttk.Label(titulo, text='Qual formato pretende exportar os dados?')
    self.mensagem.grid(row=3, column=0, columnspan=2, sticky=W+E)
    
    self.botão_excel = ttk.Button(titulo, text="Excel (.xls)", command=self.exp_veiculo_excel)
    self.botão_excel.grid(row=4, column=0)
    
    self.botao_css = ttk.Button(titulo, text='CSV (.csv)', command=self.exp_veiculo_csv)
    self.botao_css.grid(row=4, column=1)
    
    #Consulta dos dados para escrever no ficheiro:
    self.query_veic = "SELECT * FROM veiculos"
    self.dados_veiculos = self.db_consulta(self.query_veic)
    
def exp_veiculo_excel(self):
    self.janela_exportar_veiculos.destroy()
    
    self.cabecalho = []                     # Para tornar o cabeçalho elegível para Excel
    for item in self.colunas_veic:
        self.cabecalho.append(item[0])

    wb = openpyxl.Workbook()                # Criar Excel
    sheet = wb.active
        
    sheet.append(self.cabecalho)            # Introduzir nomes das colunas
    for veiculo in self.dados_veiculos:     # Introduzir os dados
        sheet.append(veiculo)
    wb.save("Relatório dos Veículos.xlsx")  # Salvar a folha num ficheiro
    
    wb.close()      # Fechar ficheiro
    
    showinfo("", "Ficheiro guardado com sucesso")   # Mensagem de ficheiro guardado
        
def exp_veiculo_csv(self):
    self.janela_exportar_veiculos.destroy()
        
    ficheiro_csv = open("Relatório dos Veículos.csv", "w")  # Criar CSV
    
    writer = csv.writer(ficheiro_csv)
    writer.writerow(self.colunas_veic)          # Introduzir nomes das colunas
    writer.writerows(self.dados_veiculos)       # Introduzir os dados
        
    del writer              # excluir objectos
    ficheiro_csv.close()    # Fechar ficheirod
        
    showinfo("", "Ficheiro guardado com sucesso")   # Mensagem de ficheiro guardado
    
    