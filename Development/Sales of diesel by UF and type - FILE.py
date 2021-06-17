#!/usr/bin/env python
# coding: utf-8

# Importar as bibliotecas


import os
import urllib
import xlwings as xw
import pandas as pd
import numpy as np
from pandas import DataFrame 


# Nomear o arquivo que será salvo


nome_arquivo = '...\Sales_of_dieadel_by_UF_and_type.xls'


# Download do arquivo que utilizaremos:


url = 'http://www.anp.gov.br/arquivos/dados-estatisticos/vendas-combustiveis/vendas-combustiveis-m3.xls'
file_name, headers = urllib.request.urlretrieve(url)


# Abertura do arquivo de apoio com a macro para revelar os dados da tabela dinamica e abertura do arquivo com as tabelas dinamicas:


wb_vba = xw.Book('...\Double_click_diesel.xlsm')
wb = xw.Book(file_name)


# Executa a macro


macro = wb_vba.macro('click')
macro()


# Se o arquivo já existir no diretório ele deleta:


try:
    os.remove(nome_arquivo)
except:
    print("Sem arquivo")


# Salva o arquivo com as abas da dinamica explicitas


wb.save(nome_arquivo)


# Fecha os arquivos


app = xw.apps.active 
app.quit()


# Importa todas as abas da planilha para um dataframe


df_oil_2 = pd.DataFrame(columns=['COMBUSTÍVEL', 'ANO', 'REGIÃO', 'ESTADO', 'UNIDADE', 'Jan', 'Fev','Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'TOTAL'])
for num in range(0,8):
    df_oil = pd.read_excel('...\Sales_of_dieadel_by_UF_and_type.xls',sheet_name=num)    
    df_oil_2 = df_oil_2.append(df_oil, ignore_index=True)


# Validação dataframe


df_oil_2


# Deleção da coluna total para fazer o pivot


del df_oil_2['TOTAL']


# Pivot a tabela para ficar colunar


df_oil_2 = df_oil_2.melt(id_vars=["COMBUSTÍVEL", "ANO", "REGIÃO", "ESTADO", "UNIDADE"], 
                var_name="MES", 
                value_name="TOTAL")


# Validação do dataframe


df_oil_2


# Validação dos valores


round(df_oil_2.groupby(['ANO', 'MES']).sum(['TOTAL']))


# De para solicitado para o mes (de extenso para numerico) e para o Estado (de nome do estado para sigla UF) e concatenação do ano com o mes


de_para_mes = {'Jan':1, 'Fev':2,'Mar':3, 'Abr':4, 'Mai':5, 'Jun':6, 'Jul':7, 'Ago':8, 'Set':9, 'Out':10, 'Nov':11, 'Dez':12} 

mes_convertido = []
for mes in df_oil_2['MES']:
    mes_convertido.append(de_para_mes.get(mes))

de_para_uf = {'ACRE':'AC', 
               'AMAZONAS':'AM',
               'AMAPÁ':'AP', 
               'ALAGOAS':'AL', 
               'BAHIA':'BA', 
               'CEARÁ':'CE', 
               'ESPÍRITO SANTO':'ES', 
               'GOIÁS':'GO', 
               'MARANHÃO':'MA', 
               'MATO GROSSO':'MT', 
               'MATO GROSSO DO SUL':'MS', 
               'MINAS GERAIS':'MG', 
               'PARÁ':'PA', 
               'PARAÍBA':'PB', 
               'PARANÁ':'PR', 
               'PERNAMBUCO':'PE', 
               'PIAUÍ':'PI', 
               'RIO DE JANEIRO':'RJ', 
               'RIO GRANDE DO NORTE':'RN', 
               'RIO GRANDE DO SUL':'RS', 
               'RONDÔNIA':'RO', 
               'RORAIMA':'RR', 
               'SANTA CATARINA':'SC', 
               'SÃO PAULO':'SP', 
               'SERGIPE':'SE', 
               'TOCANTINS':'TO', 
               'DISTRITO FEDERAL':'DF'}    

unidade_federativa = []
for uf in df_oil_2['ESTADO']:
    unidade_federativa.append(de_para_uf.get(uf))


df_oil_2['MES_NUM'] = np.array(mes_convertido)

df_oil_2['UF'] = np.array(unidade_federativa)

df_oil_2['YEAR_MONTH'] = df_oil_2['ANO'].map(str) + '-' + df_oil_2['MES_NUM'].map(str) + '-' + '01'

df_oil_2


# Validação


final = DataFrame(df_oil_2, columns = ['YEAR_MONTH', 'UF', 'COMBUSTÍVEL', 'UNIDADE', 'TOTAL'])


# Setar o year_moth como index


final.set_index('YEAR_MONTH', inplace=True)


# Exportar o dataframe


#final
final.to_csv('...\final_diesel.csv')


# In[ ]:




