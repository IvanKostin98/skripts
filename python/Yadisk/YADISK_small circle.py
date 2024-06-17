# -*- coding: utf-8 -*-
"""
Created on Wed May 10 16:08:16 2023

@author: kostin ivan
Скрипт на бекап файлов в OneDrive
"""

import os, fnmatch
from os.path import getctime
import shutil
import datetime
import sys
from my_functions_only_for_import import logging_on_bd

first = logging_on_bd(username = 'Ivan')

#блок кода для проверки доступа к серверу папки

path_to_txt=r'\\***\Для всех\***\Дистрибьюторы спец.прайс\Костин И.Н\Python скрипты\Используем\ДТ\Бекап файлов на облако\Доступ к серверу.txt'
log_file = r'\\***\Для всех\***\Дистрибьюторы спец.прайс\Костин И.Н\Python скрипты\Логирование\Лог_файл_для_Бекапа.txt'
now = datetime.datetime.now()

try:
    open_txt = open(path_to_txt, 'r')
    check_serv = True
except:
    checkt_serv = False

if check_serv == True:
    print("Есть доступ к серверу msk")
else: #если условие не выполняется
    print("Нет доступа к серверу msk")
    sys.exit(0)

def contracts(otkuda, kuda, priznak_file):
    listOfFiles = os.listdir(otkuda)
    pattern = priznak_file     
    for entry in listOfFiles:
        if fnmatch.fnmatch(entry, pattern):
            print(entry)
            print(datetime.datetime.fromtimestamp(getctime(otkuda+'/'+entry)).strftime('%H:%M:%S'))
            shutil.copy2(otkuda+'/'+entry, kuda+'/'+entry)

def only_one_file(otkuda, kuda):
    shutil.copy2(otkuda, kuda)
    
def DT_reports(otkuda, kuda, priznak_file):
    listOfFiles = os.listdir(otkuda)
    pattern = priznak_file
    for entry_papka in listOfFiles:
        try:
            listik = os.listdir(otkuda+'/'+entry_papka)
            for entry_file in listik:
                if fnmatch.fnmatch(entry_file, pattern):
                    shutil.copy2(otkuda+'/'+entry_papka+'/'+entry_file, kuda+'/'+entry_file)
        except:
            continue

def handleAssetsFile(path, path1, name_create_papka):      
    if os.path.exists(path1):
        print (path1, 'существует, поэтому сначала удаляю')
        shutil.rmtree(path1+'/'+name_create_papka)
        print ('Начинаю копирование папки с файлами ...')
        shutil.copytree(path, path1+'/'+name_create_papka)
        print ('Заканчиваю копирование папки с файлами! \ n')
        
def work_cycle(otkuda, kuda, name_papka_1):
    if __name__ == "__main__":
        handleAssetsFile(otkuda+'/'+name_papka_1, kuda, name_papka_1)

        
sp_contracts = [[r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Северо-Запад\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\Северо-Запад\2024'],
      [r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Дальний Восток\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\Дальний Восток\2024'],
      [r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Москва\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\Москва\2024'],
      [r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Омск\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\Омск\2024'],
      [r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Сибирь\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\Сибирь\2024'],
      [r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\СПБ\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\СПБ\2024'],
      [r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Урал\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\Урал\2024'],
      [r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Центр\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\Центр\2024'],
      [r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Юг\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\ЮГ\2024'],
      [r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Приморье\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Контракты\Приморье\2024'],
      [r'\\***\Для всех\Третьякова М\Pricing\цены\протоколы\2024\Протокол РФ', r'C:\Users\kostin\YandexDisk\New_back_up\Протоколы\2024'],
      [r'\\***\Для всех\Третьякова М\Pricing\цены\цс\2024', r'C:\Users\kostin\YandexDisk\New_back_up\ЦС\2024']]

sp_file = [[r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\!Реестр дополнительных начислений ЧС.xlsm', r'C:\Users\kostin\YandexDisk\New_back_up\Реестр ДН\!Реестр дополнительных начислений ЧС.xlsm'],
           [r'\\***\Для всех\Третьякова М\Pricing\Рита\РКУ\РКУ 2024.xlsx', r'C:\Users\kostin\YandexDisk\New_back_up\Реестр КУ\РКУ 2024.xlsx']]

spisok_kakaya_papka_dlya_skrripts = ['Python скрипты','SQL скрипты','VBA скрипты']
counter=0
# рабочие циклы
for i in range(len(sp_contracts)):
    try:
        contracts(sp_contracts[i][0], sp_contracts[i][1], "*.xl*")
        print(sp_contracts[i][0])
    except:
        counter+=1

for i in range(len(sp_file)):
    try:
        only_one_file(sp_file[i][0], sp_file[i][1])
    except:
        counter+=1
try:
    DT_reports(r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Отчёт по ДТ\2024', r'C:\Users\kostin\YandexDisk\New_back_up\Отчеты ДТ\2024', 'P&L*')   
    print('отчёт дт ОК')
except:
    counter+=1

try:
    for i in range(len(spisok_kakaya_papka_dlya_skrripts)):
        work_cycle(r'\\***\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Костин И.Н', r'C:\Users\kostin\YandexDisk\New_back_up\Codes', spisok_kakaya_papka_dlya_skrripts[i])
except:
    counter+=1

#калькуляторы с сервера
try:
    contracts(r'\\***\Для всех\Calc\LKA', r'C:\Users\kostin\YandexDisk\New_back_up\Калькуляторы', '*Калькулятор*')
except:
    counter+=1

# логирование
if counter>1:
    first.bd_open()
    first.bd_process(filename = 'OneDrive', processname = 'Малый круг', result = '0', comments = 'ошибка')
    first.bd_close()
else:
    first.bd_open()
    first.bd_process(filename = 'OneDrive', processname = 'Малый круг', result = '1', comments = 'загружено')
    first.bd_close()
    
    
    
