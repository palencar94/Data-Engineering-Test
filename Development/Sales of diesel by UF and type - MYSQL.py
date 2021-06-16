
# coding: utf-8

# In[3]:


import pymysql.cursors
import pandas as pd
import mysql.connector
from mysql.connector import MySQLConnection


# In[4]:


from sqlalchemy import create_engine

engine = create_engine('mysql+pymysql://root:secret@172.18.0.5/oil?charset=utf8')


# In[6]:


final = pd.read_csv('final_diesel.csv')


# In[7]:


#final.set_index('YEAR_MONTH', inplace=True)
final


# In[8]:


final.to_sql('stg_diesel', con = engine, if_exists = 'replace', index = False)


# In[9]:


conn = mysql.connector.connect(host='172.18.0.5',
                                       database='oil',
                                       user='root',
                                       password='secret')
cursor_create = conn.cursor()
cursor = conn.cursor()


# In[18]:


cursor_create.execute("CREATE TABLE ods_diesel PARTITION BY HASH(MONTH(`YEAR_MONTH`)) AS SELECT STR_TO_DATE(`YEAR_MONTH`, '%Y-%m-%d') AS `year_month`, UF AS ´uf´, `COMBUSTÍVEL` AS `product`, `UNIDADE` AS `unit`, `TOTAL` AS `volume`, CURRENT_TIMESTAMP() AS `created_at` FROM stg_diesel")
   
cursor.execute("SELECT ´uf´, `product`, SUM(`volume`) FROM ods_diesel GROUP BY 1, 2 ORDER BY 3 DESC")    


# In[19]:


row = cursor.fetchone()


# In[20]:


while row is not None:
            print(row)
            row = cursor.fetchone()

