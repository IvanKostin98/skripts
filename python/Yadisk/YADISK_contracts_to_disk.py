# -*- coding: utf-8 -*-
"""
Created on Mon Apr  1 10:48:55 2024

@author: kostin
"""
from yadisk import YaDisk
import os, fnmatch
from my_functions_only_for_import import logging_on_bd
from retrying import retry


class Yadisk:
    def __init__(self, file_path, yadisk_path, path_to_RDN, path_to_yadisk_RDN):
        self.file_path = file_path
        self.yadisk_path = yadisk_path
        self.token = YaDisk(token='y0_AgAEA7qkiwzgAADLWwAAAAEAh6GBAACxh5brZjdGx6Kkgt9UrTXKV2Btiw')
        self.token.retries = 5
        self.path_to_RDN = path_to_RDN
        self.path_to_yadisk_RDN = path_to_yadisk_RDN
    
    def download_to_yadisk(self, file_name):
       """Загрузка на Ядиск"""
       with open(f"{self.path_to_RDN}\{file_name}", 'rb') as f:
           self.token.upload(f, f'{self.path_to_yadisk_RDN}{file_name}', timeout=30)
    
    @retry(stop_max_attempt_number = 5)
    def download_to_yadisk_divisions(self, file_name, serv_path, ya_path):
        """Загрузка на Ядиск папок которые разбиты по дивизионам"""
        with open(f"{serv_path}\{file_name}", 'rb') as f:
            self.token.upload(f, f'{ya_path}{file_name}', timeout=30)
    
    def clear_ya_disk(self, clear_path):
        """Очистка папки на Ядиске"""
        files = self.token.listdir(clear_path)
        for file_info in files:
            file_path = file_info['path']
            self.token.remove(file_path)
            print(f'Delete {file_path}')
            
    def contracts(self, priznak_file, papka):
        """Находим список файлов"""
        sp_files = list()
        listOfFiles = os.listdir(papka)
        for entry in listOfFiles:
            if fnmatch.fnmatch(entry, priznak_file):
                sp_files.append(entry)
        return sp_files

    def process_for_contracts(self, name_of_division):
        """В эту болванку добавляем процессы, чтобы не запускать каждый метод по отдельности"""
        self.clear_ya_disk(self.yadisk_path)
        for i in self.contracts('*.xl*', self.file_path):
            self.download_to_yadisk_divisions(i, self.file_path, self.yadisk_path)
            print(i)
        
        self.clear_ya_disk(self.path_to_yadisk_RDN)
        sp_files = self.contracts('*.xl*', self.path_to_RDN)
        for i in sp_files:
            if name_of_division in i:
                self.download_to_yadisk(i)
                print(self.path_to_yadisk_RDN)
            else:
                print(f'Вылет по {name_of_division}')
        

    


class_dv = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Дальний Восток\2024', 'Division. Дальний Восток/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Дальний Восток/Дистрибьюторы/Компенсации ниже МОЦ/')
class_msk = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Москва\2024', 'Division. Москва/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Москва/Дистрибьюторы/Компенсации ниже МОЦ/')
class_omsk = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Омск\2024', 'Division. Омск/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Омск/Дистрибьюторы/Компенсации ниже МОЦ/')
class_prim = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Приморье\2024', 'Division. Приморье/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Приморье/Дистрибьюторы/Компенсации ниже МОЦ/')
class_spb = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\СПБ\2024', 'Division. Санкт-Петербург/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Санкт-Петербург/Дистрибьюторы/Компенсации ниже МОЦ/')
class_szfo = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Северо-Запад\2024', 'Division. Северо-Запад/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Северо-Запад/Дистрибьюторы/Компенсации ниже МОЦ/')
class_sib = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Сибирь\2024', 'Division. Сибирь/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Сибирь/Дистрибьюторы/Компенсации ниже МОЦ/')
class_ural = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Урал\2024', 'Division. Урал/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Урал/Дистрибьюторы/Компенсации ниже МОЦ/')
class_center = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Центр\2024', 'Division. Центр/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Центр/Дистрибьюторы/Компенсации ниже МОЦ/')
class_ug = Yadisk(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Юг\2024', 'Division. Юг/Дистрибьюторы/Согласованные калькуляторы/2024/', r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Реестр ДН нарезка по дивизионам', 'Division. Юг/Дистрибьюторы/Компенсации ниже МОЦ/')


class_dv.process_for_contracts("Дальний Восток")
class_msk.process_for_contracts("Москва")
class_omsk.process_for_contracts("Омск")
class_prim.process_for_contracts("Приморье")
class_spb.process_for_contracts("Санкт-Петербург")
class_szfo.process_for_contracts("Северо-Запад")
class_sib.process_for_contracts("Сибирь")
class_ural.process_for_contracts("Урал")
class_center.process_for_contracts("Центр")
class_ug.process_for_contracts("Юг")


first = logging_on_bd(username = 'Ivan')
first.bd_open()
first.bd_process(filename = 'YaDisk', processname = 'Загрузка контрактов', result = '1', comments = 'загружено')
first.bd_close()












