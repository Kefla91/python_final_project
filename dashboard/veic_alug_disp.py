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

      
def veic_alugados(self):
    
    #Estrutura do gráfico:
    self.espacoLinha = Label(self.relatorio, text='')
    self.espacoLinha.grid(row=14, column=1, columnspan=3)
    
    self.etiqueta_veiculos_alugados = Label(self.relatorio, text="Veículos Alugados + Dias restantes de Reserva")
    self.etiqueta_veiculos_alugados.grid(row=15, column=1)
        
    figura = plt.figure(figsize=(8, 4), dpi=60)
    ax = figura.add_subplot(111) #grafico
        
    canva= FigureCanvasTkAgg(figura, self.etiqueta_veiculos_alugados)
    canva.get_tk_widget().grid(row=1, column=0)
    
    self.get_veic_alugados()
    
    id_veiculoAlug = self.lista_idVeicAlug
    y_pos = np.arange(len(self.lista_veicAlugados))
    dias_restantes = self.lista_diasRest_alug
    error = np.random.rand(len(self.lista_veicAlugados))

    hbars = ax.barh(y_pos, dias_restantes, xerr=error, align='center')
    ax.set_yticks(y_pos, labels=id_veiculoAlug)
    ax.invert_yaxis()  # labels read top-to-bottom
    ax.set_xlabel('Dias restantes de aluguer')
    ax.set_ylabel('ID Veículo')
    ax.set_title('Veículos alugados')

    # Label with specially formatted floats
    ax.bar_label(hbars, fmt='%.2f')
    ax.set_xlim(right=40)  # adjust xlim
    
def get_veic_alugados(self):
    query_reservas = 'SELECT * FROM reservas'
    registos_reservas = self.db_consulta(query_reservas)  # Faz-se a chamada ao método db_consultas
    
    query_veic = 'SELECT * FROM veiculos'
    registos_veic = self.db_consulta(query_veic)
        
    # Listas de dados dos Veiculos alugados:
    self.lista_veicAlugados = []
    # Listas de matrículas e dias rest aluguer dos carros alugados para o gráfico
    self.lista_idVeicAlug = []
    self.lista_diasRest_alug = []
    
    # Listas de dados dos veículos disponíveis:
    self.lista_veicDisp = []
    # Lista tipo Veic Disponivel para o gráfico
    self.lista_tipoVeic_disp = []

    # loop for para preencher lista com veículos alugados e lista c veic disponíveis:
    for reserva in registos_reservas:
        data_inicio_res = datetime.strptime(reserva[3], self.formato_data_num)                    # Formatar string para datetime
        data_fim_res = datetime.strptime(reserva[4], self.formato_data_num)
        diferenca_dias_inicio = data_inicio_res - self.dt                                       # Há quantos dias está reservado
        diferenca_dias_fim = data_fim_res - self.dt                                             # Quantos dias ainda vai estar reservado
        
        id_veicReservado = reserva[2]    # ID Veiculo na query RESERVA
        if diferenca_dias_inicio <= self.dif_data_nula and diferenca_dias_fim >= self.dif_data_nula: # Caso o veiculo esteja alugado no momento
            #print('ID Veic Alugados: ', veiculo_reserva)    
            self.lista_veicAlugados.append({"idVeiculo":id_veicReservado, "diasRest":diferenca_dias_fim.days})
            self.lista_idVeicAlug.append(id_veicReservado)
            self.lista_diasRest_alug.append(diferenca_dias_fim.days)
            #print('No get\ndentro do loop:\nID veic Alugado:', veiculo_id)               # Para debug'''


    for veiculo in registos_veic:                               # Percorre os veiculos
            veiculo_id = veiculo[0]                                 # ID do veiculo na query VEIC
        #for item in self.lista_idVeicAlug:
            if veiculo_id not in self.lista_idVeicAlug and veiculo[11] == "Sim": # Se (ID veiculo do loop é = ao do veiculo disp) E Se (não foi colocado em manutenção)
                self.lista_veicDisp.append({"idVeiculo":veiculo_id, "matricula":veiculo[1], "tipoVeic":veiculo[2], "categVeic":veiculo[5]})
                self.lista_tipoVeic_disp.append(veiculo[2])


def veic_disponiveis(self):
    # Construçºao do label:
    self.etiqueta_quant_veiculos_disp = Label(self.relatorio, text="Quantidade de Veículos Disponíveis")
    self.etiqueta_quant_veiculos_disp.grid(row=15, column=3)

    figura = plt.figure(figsize=(8, 4), dpi=60)
    ax = figura.add_subplot(111) #grafico
        
    canva= FigureCanvasTkAgg(figura, self.etiqueta_quant_veiculos_disp)
    canva.get_tk_widget().grid(row=0, column=0)
    
    # Variáveis para contagem de número de veículos por tipo:
    qtd_carros_disp=0
    qtd_motas_disp=0
    qtd_monoVolume_disp=0
    qtd_carrinhas_disp=0
    qtd_camioes_disp=0
    
    # Variáveis para contagem por tipo de veículo e categoria:
    cont_carro_pequeno=0
    cont_carro_grande=0
    cont_carro_medio=0
    
    cont_mota_pequena=0
    cont_mota_grande=0
    cont_mota_media=0
    
    cont_monoVol_pequeno=0
    cont_monoVol_grande=0
    cont_monoVol_medio=0

    cont_carrinhas_pequena=0
    cont_carrinhas_grande=0
    cont_carrinhas_media=0

    cont_camioes_pequeno=0
    cont_camioes_grande=0
    cont_camioes_medio=0
    
    # Contagem do índice para utilizar dentro do loop for
    contagem_indice=0

    # Loop para contagem de cada tipo de veículo e sua categoria:
    for tipo in self.lista_tipoVeic_disp:
        
        categoria = self.lista_veicDisp[contagem_indice]["categVeic"]
        
        if tipo =='Carro':
            qtd_carros_disp +=1
            if categoria == 'Pequeno':
                cont_carro_pequeno +=1
            elif categoria == 'Médio':
                cont_carro_medio +=1
            elif categoria == 'Grande':
                cont_carro_grande += 1
            
        elif tipo == 'Mota':
            qtd_motas_disp += 1
            if categoria == 'Pequeno':
                cont_mota_pequena +=1
            elif categoria == 'Médio':
                cont_mota_media +=1
            elif categoria == 'Grande':
                cont_mota_grande += 1 

        elif tipo == 'MonoVolume':
            qtd_monoVolume_disp +=1
            if categoria == 'Pequeno':
                cont_monoVol_pequeno +=1
            elif categoria == 'Médio':
                cont_monoVol_medio +=1
            elif categoria == 'Grande':
                cont_monoVol_grande += 1
            
        elif tipo == 'Carrinha':
            qtd_carrinhas_disp += 1
            
            if categoria == 'Pequeno':
                cont_carrinhas_pequena +=1
            elif categoria == 'Médio':
                cont_carrinhas_media +=1
            elif categoria == 'Grande':
                cont_carrinhas_grande += 1 
            
        elif tipo  == 'Camião':
            qtd_camioes_disp += 1

            if categoria == 'Pequeno':
                cont_camioes_pequeno +=1
            elif categoria == 'Médio':
                cont_camioes_medio +=1
            elif categoria == 'Grande':
                cont_camioes_grande += 1 
            
        contagem_indice +=1

    tipos_veiculo = ('Carros', 'Motas','MonoVolume', 'Carrinhas', 'Camiões')
        
    categoria_veiculo = {
        'Pequeno': np.array([cont_carro_pequeno, cont_mota_pequena, cont_monoVol_pequeno, cont_carrinhas_pequena, cont_camioes_pequeno]),
        'Médio': np.array([cont_carro_medio, cont_mota_media, cont_monoVol_medio, cont_carrinhas_media ,cont_camioes_medio]),
        'Grande': np.array([cont_carro_grande, cont_mota_grande, cont_monoVol_grande,cont_carrinhas_grande ,cont_camioes_grande]),
    }
    width = 0.6  # the width of the bars: can also be len(x) sequence

    bottom = np.zeros(5)        # Número de barras (nº de tipos de veículos)

    for veiculo, categoria_veiculo in categoria_veiculo.items():
        p = ax.bar(tipos_veiculo, categoria_veiculo, width, label=veiculo, bottom=bottom)
        bottom += categoria_veiculo

        ax.bar_label(p, label_type='center')

    ax.set_title(f'Número de veículos Disponíveis: {len(self.lista_veicDisp)}')
    ax.legend()
        
