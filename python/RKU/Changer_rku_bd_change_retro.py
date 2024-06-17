# -*- coding: utf-8 -*-
"""
Created on Fri Feb  9 11:51:00 2024

@author: kostin ivan

1. Дозаливка инфо из РКУ Маргариты в бд
2. Дозаливка инфо из Номенклатуры в бд
"""
import pandas as pd
import pyodbc


def logging(x):
    log_file = r'\\***\Для всех\***\Дистрибьюторы спец.прайс\***\Python скрипты\Логирование\Лог_файл_changer_rku.txt'
    with open(log_file , 'a') as file:
        file.writelines(x)
        
def file_to_df(path, sheet, skiprows, columns):
    file = pd.read_excel(path, sheet_name = sheet, skiprows=skiprows)
    df = pd.DataFrame(file, columns = columns)
    return df
        
#Переменные
name_dt = ''
counter = 0

#Подключение к бд/загрузка файла
conn = pyodbc.connect(driver='{SQL Server}',
                      server="***", 
                      database="FinanceAndSAP",               
                      trusted_connection="yes")
cur = conn.cursor()

try:
    cur.execute(f"""update [FinanceAndSAP].[segment].[Agreements]
                    set [IndividRB] = [StandartRB]
                    where [TillDate] = '2024-12-31 00:00:00.000' and [SegmentName] = {name_dt} and ([StandartRB] - [IndividRB])<>0 and Brand = 'Пять Озер'
                    """)
except:
    counter+=1

conn.commit()
conn.close()
