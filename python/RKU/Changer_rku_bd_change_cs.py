# -*- coding: utf-8 -*-
"""
Created on Fri Feb  9 11:51:00 2024

@author: kostin ivan

1. Дозаливка инфо из РКУ Маргариты в бд
"""
import pandas as pd
import pyodbc
import datetime
from tqdm import tqdm

def logging(x):
    log_file = r'\\***\Для всех\***\Дистрибьюторы спец.прайс\***\Python скрипты\Логирование\Лог_файл_changer_rku.txt'
    with open(log_file , 'a') as file:
        file.writelines(x)
        
def file_to_df(path, sheet, skiprows, columns):
    file = pd.read_excel(path, sheet_name = sheet, skiprows=skiprows)
    df = pd.DataFrame(file, columns = columns)
    return df
        
#Переменные
path_to_rku_excel = r'\\***\Для всех\***\Pricing\***\РКУ\РКУ 2024.xlsx'
sp_columns = ['Код SKU', 'Код сегмента', 'Бренд', 'BI название ДТ', 'Дивизион', 'Название сегмента', 'Название SKU', 
              'Емкость', 'Срок с', 'Срок по', 'Цена', 'скидка за непринятие', 'скидка для торгового дома', 
              'Partnership discount MT текущий', 'Trade Marketing discount conditional(incl CLAP) текущий', 
              'Logistics discount текущий', 'EDLP текущий', 'Reb', 'Регулярная цена без НДС', 'Регулярная цена с НДС', 
              'самовывоз', 'предоплата', 'ЦЕНА', 'МОЦ', 'ДТ с ндс', 'ТТсндс', 'МТсндс', 'ОТсндс', 'ОТ/ТТ', 'наценка ДТ', 
              'РБ', 'вход с уч РБ', 'наценка ДТ с РД', 'ДТ с ндс.1', 'ТТ с ндс', 'МТ с ндс', 'ОТ с ндс', 'ОТ/ТТ.1', 
              'наценка ДТ.1', 'вход с уч РБ.1', 'наценка ДТ с РД.1', 'ТТ', 'МТ', 'ОТ']
now = datetime.datetime.now()
counter = 0

#Подключение к бд/загрузка файла
conn = pyodbc.connect(driver='{SQL Server}',
                      server="***", 
                      database="FinanceAndSAP",               
                      trusted_connection="yes")
cur = conn.cursor()
df = file_to_df(path_to_rku_excel, 'Загрузчик КУ в 1С', 4, sp_columns)
df.rename(columns = {'ДТ с ндс':'ДТ_с_ндс_стандарт', 'ТТсндс':'ТТ_с_ндс_стандарт', 'МТсндс':'МТ_с_ндс_стандарт',
                      'ОТсндс':'ОТ_с_ндс_стандарт', 'ОТ/ТТ':'ОТ_ТТ_стандарт', 'наценка ДТ':'Наценка_ДТ_стандарт',
                      'РБ':'РБ_стандарт', 'вход с уч РБ':'Вход_с_уч_РБ_стандарт', 'наценка ДТ с РД':'Наценка_ДТ_с_РБ_стандарт',
                      'ДТ с ндс.1':'ДТ_с_ндс_индивид', 'ТТ с ндс':'ТТ_с_ндс_индивид', 'МТ с ндс':'МТ_с_ндс_индивид',
                      'ОТ с ндс':'ОТ_с_ндс_индивид', 'ОТ/ТТ.1':'ОТ_ТТ_индивид', 'наценка ДТ.1':'Наценка_ДТ_индивид',
                      'вход с уч РБ.1':'Вход_с_уч_РБ_индивид', 'наценка ДТ с РД.1':'Наценка_ДТ_с_РБ_индивид',
                      'Код SKU':'Код_SKU', 'Код сегмента':'Код_сегмента', 'Название SKU':'SKU_name'
                      }, inplace = True)
df.drop(index=df.index [0], axis= 0, inplace= True)

#бахаем цски
for row in tqdm(df.itertuples(index=True)):
    try:
        cur.execute(f"""update [FinanceAndSAP].[segment].[Agreements]
                        set [StandartDT] = {row.ДТ_с_ндс_стандарт}, [StandartTT] = {row.ТТ_с_ндс_стандарт}, [StandartMT] = {row.МТ_с_ндс_стандарт},
                            [StandartOT] = {row.ОТ_с_ндс_стандарт}, [StandartOT_TT] = {row.ОТ_ТТ_стандарт}, [StandartExtraDT] = {row.Наценка_ДТ_стандарт}, 
                            [StandartRB] = {row.РБ_стандарт}, [StandartEntryRB] = {row.Вход_с_уч_РБ_стандарт}, [StandartExtraDT_RB] = {row.Наценка_ДТ_с_РБ_стандарт},
                            [IndividDT] = {row.ДТ_с_ндс_индивид}, [IndividTT] = {row.ТТ_с_ндс_индивид}, [IndividMT] = {row.МТ_с_ндс_индивид},
                            [IndividOT] = {row.ОТ_с_ндс_индивид}, [IndividOT_TT] = {row.ОТ_ТТ_индивид}, [IndividExtraDT] = {row.Наценка_ДТ_индивид},
                            [IndividEntryRB] = {row.Вход_с_уч_РБ_индивид}, [IndividExtraDT_RB] = {row.Наценка_ДТ_с_РБ_индивид}
                            
                        where [TillDate] = '2024-12-31 00:00:00.000' and [SKU_ID] = {int(row.Код_SKU)} and SegmentCode = '{row.Код_сегмента}'
                        """)
    except:
        counter+=1

conn.commit()
conn.close()
logging(f"\n{now.strftime('%Y-%m-%d')} - {counter} вылетов")
