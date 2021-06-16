
# coding: utf-8

# Importar as bibliotecas


import pymysql.cursors
import pandas as pd
import mysql.connector
from mysql.connector import MySQLConnection
from sqlalchemy import create_engine

# Conexão com o container mysql

engine = create_engine('mysql+pymysql://oil:secret@172.18.0.5/oil?charset=utf8')


# Leitura do csv


final = pd.read_csv('final.csv')


# Validação


#final.set_index('YEAR_MONTH', inplace=True)
final


# Export do dataframe para o mysql em uma tabel de stage


final.to_sql('stg_oil', con = engine, if_exists = 'replace', index = False)


# Conexao com o container mysql


conn = mysql.connector.connect(host='172.18.0.5',
                                       database='oil',
                                       user='oil',
                                       password='secret')
cursor_create = conn.cursor()
cursor = conn.cursor()


# carga da tabela de stage para a tabela de produção e query com o indicador solicitado no desafio


cursor_create.execute("CREATE TABLE ods_oil PARTITION BY HASH(MONTH(`YEAR_MONTH`)) AS SELECT STR_TO_DATE(`YEAR_MONTH`, '%Y-%m-%d') AS `year_month`, UF AS ´uf´, `COMBUSTÍVEL` AS `product`, `UNIDADE` AS `unit`, `TOTAL` AS `volume`, CURRENT_TIMESTAMP() AS `created_at` FROM stg_oil")
   
cursor.execute("SELECT ´uf´, `product`, SUM(`volume`) FROM ods_oil GROUP BY 1,2 ORDER BY 3 DESC")    


# fecha o cursor das querys


row = cursor.fetchone()


# Mostra o resultado da query


while row is not None:
            print(row)
            row = cursor.fetchone()

