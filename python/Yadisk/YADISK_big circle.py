"""
Created on Tue May 16 14:57:47 2023

@author: kostin ivan

первая функция на очистку и копирование
вторая функция пробует найти самую актуальную папку

"""

import os
import shutil
import datetime
import sys
import time
from my_functions_only_for_import import logging_on_bd

first = logging_on_bd(username = 'Ivan')
 
start_price = time.perf_counter()
# Проверка доступа к серверу папки
path_to_txt = r'\\***\Для всех\***\Дистрибьюторы спец.прайс\***\Python скрипты\Используем\ДТ\Бекап файлов на облако\Доступ к серверу.txt'
 
check_serv = os.path.exists(path_to_txt)
 
if check_serv:
    print("Есть доступ к серверу msk")
else:
    print("Нет доступа к серверу msk")
    sys.exit(0)
 
def handle_assets_file(path, path1, name_create_papka):
    if os.path.exists(path1):
        print(path1, 'Существует, удаляем')
        shutil.rmtree(os.path.join(path1, name_create_papka))
        print('Начинаем копирование папки ...')
        shutil.copytree(path, os.path.join(path1, name_create_papka))
        print('Конец копирования папки!\n')
 
def work_cycle(otkuda, kuda, name_papka, create_papka_imya):
    if __name__ == "__main__":
        handle_assets_file(os.path.join(otkuda, name_papka), kuda, create_papka_imya)
 
def try_except(papka):
    try:
        os.mkdir(f'C:\\Users\\kostin\\YandexDisk\\New_back_up\\Шаблоны\\{papka}')
    except FileExistsError:
        print('Папка уже существует')
 
dt_now = datetime.datetime.now()
dt_m = dt_now.month
 
print(dt_m)
 
zaeb_slovar = {
    2: '02_Февраль', 3: '03_Март', 4: '04_Апрель', 5: '05_Май', 6: '06_Июнь',
    7: '07_Июль', 8: '08_Август', 9: '09_Сентябрь',
    10: '10_Октябрь', 11: '11_Ноябрь', 12: '12_Декабрь'
}
 
pzdc_slovar = {
    3: 'RF 03_2024', 4: 'RF 04_2024', 5: 'RF 05_2024', 6: 'RF 06_2024',
    7: 'RF 07_2024', 8: 'RF 08_2024', 9: 'RF 10_2024',
    10: 'RF 10_2024', 11: 'RF 11_2024', 12: 'RF 12_2024'
}
 
papka_name = zaeb_slovar.get(dt_m)
papka_name_2 = zaeb_slovar.get(dt_m - 1)
papka_name_3 = zaeb_slovar.get(dt_m - 2)
papka_sop = pzdc_slovar.get(dt_m)
papka_sop_2 = pzdc_slovar.get(dt_m - 1)
papka_sop_3 = pzdc_slovar.get(dt_m - 2)
 
print(papka_name, 'Закрытие')
print(papka_name_2, 'Закрытие 2')
print(papka_name_3, 'Закрытие 3')
print(papka_sop, 'S&OP')
print(papka_sop_2, 'S&OP 2')
print(papka_sop_3, 'S&OP 3')
 
pyt_k_zakritiu = r'\\***\Для всех\***\Дистрибьюторы спец.прайс\закрытие\2024'
pyt_k_zakritiu_backup = r'C:\\Users\\kostin\\YandexDisk\\New_back_up\\Шаблоны'
pyt_k_sop = r'\\***\Для всех\***\Дистрибьюторы спец.прайс\S&OP\2024'
pyt_k_sop_backup = r'C:\\Users\\kostin\\YandexDisk\\New_back_up\\Шаблоны'

try:
    try_except('Закрытие 1С')
    work_cycle(pyt_k_zakritiu, pyt_k_zakritiu_backup, papka_name, 'Закрытие 1С')
    first.bd_open()
    first.bd_process(filename = 'OneDrive', processname = 'Большой круг', result = '1', comments = f'загружено {papka_name}')
    first.bd_close()
except:
    try:
        try_except('Закрытие 1С')
        work_cycle(pyt_k_zakritiu, pyt_k_zakritiu_backup, papka_name_2, 'Закрытие 1С')
        first.bd_open()
        first.bd_process(filename = 'OneDrive', processname = 'Большой круг', result = '1', comments = f'загружено {papka_name_2}')
        first.bd_close()
    except:
        try_except('Закрытие 1С')
        work_cycle(pyt_k_zakritiu, pyt_k_zakritiu_backup, papka_name_3, 'Закрытие 1С')
        first.bd_open()
        first.bd_process(filename = 'OneDrive', processname = 'Большой круг', result = '1', comments = f'загружено {papka_name_3}')
        first.bd_close()
    
try:
    try_except('S&OP')
    work_cycle(pyt_k_sop, pyt_k_sop_backup, papka_sop, 'S&OP')
    first.bd_open()
    first.bd_process(filename = 'OneDrive', processname = 'Большой круг', result = '1', comments = f'загружено {papka_sop}')
    first.bd_close()
except:
    try:
        try_except('S&OP')
        work_cycle(pyt_k_sop, pyt_k_sop_backup, papka_sop_2, 'S&OP')
        first.bd_open()
        first.bd_process(filename = 'OneDrive', processname = 'Большой круг', result = '1', comments = f'загружено {papka_sop_2}')
        first.bd_close()
    except:
        try_except('S&OP')
        work_cycle(pyt_k_sop, pyt_k_sop_backup, papka_sop_3, 'S&OP')
        first.bd_open()
        first.bd_process(filename = 'OneDrive', processname = 'Большой круг', result = '1', comments = f'загружено {papka_sop_3}')
        first.bd_close()

end_price = time.perf_counter()
print(f'Скрипт отработал за: {round(end_price - start_price)} сек.')       
        
        
        
