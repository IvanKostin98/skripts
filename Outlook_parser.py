# -*- coding: utf-8 -*-
"""
Created on Thu May 30 16:25:04 2024

@author: kostin
"""

import win32com.client, os
from my_functions_only_for_import import logging_on_bd

log_class = logging_on_bd('Иван')

log_class.ignore_warnings()         #игнорим предупреждения


def replace_custom(item):
    f = item.replace(':', '')
    f = f.replace('FW ', '')
    f = f.replace('RE ', '')
    f = f.replace('Re ', '')
    f = f.replace('/', ' ')
    return f
 
def extract_emails_by_subject(folder_name, subject):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
   
    # Получаем папку по ее названию
    folder = namespace.Folders.Item("ivan.kostin@asg.ru").Folders.Item(folder_name).Folders.Item("Согласования")
   
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # Сортировка сообщений по времени получения

    for message in messages:
        if subject in str(message):
            try:
                message.SaveAs(os.path.join(r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Костин И.Н\Согласования 2024', f"{replace_custom(message.Subject)}.msg"))
            except:
                print(message)
            
    outlook.Quit()
    
extract_emails_by_subject("Архив", "2024")

log_class.bd_open()
log_class.bd_process('Outlook', 'Парсинг согласований', '1', 'Загружено')
log_class.bd_close()