#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import urllib
import xlwings as xw
import pandas as pd
import numpy as np
from pandas import DataFrame 


# In[2]:


nome_arquivo = 'D:\\Raizen\\Sales_of_dieadel_by_UF_and_type.xls'


# In[3]:


url = 'http://www.anp.gov.br/arquivos/dados-estatisticos/vendas-combustiveis/vendas-combustiveis-m3.xls'
file_name, headers = urllib.request.urlretrieve(url)


# In[4]:


wb_vba = xw.Book('D:\\Raizen\\Double_click_diesel.xlsm')
wb = xw.Book(file_name)


# In[5]:


macro = wb_vba.macro('click')
macro()


# In[6]:


try:
    os.remove(nome_arquivo)
except:
    print("Sem arquivo")


# In[7]:


wb.save(nome_arquivo)


# In[8]:


app = xw.apps.active 
app.quit()


# In[9]:


df_oil_2 = pd.DataFrame(columns=['COMBUSTÍVEL', 'ANO', 'REGIÃO', 'ESTADO', 'UNIDADE', 'Jan', 'Fev','Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'TOTAL'])
for num in range(0,8):
    df_oil = pd.read_excel('D:\\Raizen\\Sales_of_dieadel_by_UF_and_type.xls',sheet_name=num)    
    df_oil_2 = df_oil_2.append(df_oil, ignore_index=True)


# In[10]:


df_oil_2


# In[11]:


del df_oil_2['TOTAL']


# In[12]:


df_oil_2 = df_oil_2.melt(id_vars=["COMBUSTÍVEL", "ANO", "REGIÃO", "ESTADO", "UNIDADE"], 
                var_name="MES", 
                value_name="TOTAL")


# In[13]:


df_oil_2


# In[14]:


round(df_oil_2.groupby(['ANO', 'MES']).sum(['TOTAL']))


# In[15]:


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


# In[16]:


final = DataFrame(df_oil_2, columns = ['YEAR_MONTH', 'UF', 'COMBUSTÍVEL', 'UNIDADE', 'TOTAL'])


# In[17]:


final.set_index('YEAR_MONTH', inplace=True)


# In[18]:


#final
final.to_csv('D:\\Raizen\\final_diesel.csv')


# In[ ]:




