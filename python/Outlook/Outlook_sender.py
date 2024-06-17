# -*- coding: utf-8 -*-
"""
Created on Mon Apr  3 17:28:58 2023

@author: kostin
"""

import datetime
import os
import shutil
from pathlib import Path
import pandas as pd
import win32com.client as win32
#################################
per_ya = '02'
mes = "февраль"
papka=r'\\*****\Для всех\****\Дистрибьюторы спец.прайс\***\2024\**\Division'

#################################

today_string2 = datetime.datetime.today().strftime('%b %d, %Y')
df = {'CUSTOMER_ID': ['ФИО', 'ФИО', 'ФИО', 'ФИО', 'ФИО', 'ФИО', 'ФИО', 'ФИО', 'ФИО', 'ФИО'],
    'EMAIL': ['111@asg.ru', '222@asg.ru', '333@asg.ru','444@asg.ru', '555@asg.ru', '666@asg.ru', '777@asg.ru', '888@asg.ru', '999@asg.ru', '000@asg.ru'],
    'FILE': [papka+'/'+'P&L по DT '+per_ya+'.2024_Дальний Восток.xlsx', 
             papka+'/'+'P&L по DT '+per_ya+'.2024_Москва.xlsx', 
             papka+'/'+'P&L по DT '+per_ya+'.2024_Омск.xlsx', 
             papka+'/'+'P&L по DT '+per_ya+'.2024_СПБ.xlsx', 
             papka+'/'+'P&L по DT '+per_ya+'.2024_СЗФО.xlsx', 
             papka+'/'+'P&L по DT '+per_ya+'.2024_Сибирь.xlsx', 
             papka+'/'+'P&L по DT '+per_ya+'.2024_Урал.xlsx', 
             papka+'/'+'P&L по DT '+per_ya+'.2024_Центр.xlsx', 
             papka+'/'+'P&L по DT '+per_ya+'.2024_Юг.xlsx',
             papka+'/'+'P&L по DT '+per_ya+'.2024_Приморье.xlsx']}

combined = pd.DataFrame(df)

# Отправка индивидуальных отчетов по электронной почте соответствующим получателям
class EmailsSender:
    def __init__(self):
        self.outlook = win32.Dispatch('outlook.application')
    def send_email(self, to_email_address, attachment_path):
        mail = self.outlook.CreateItem(0)
        mail.To = to_email_address
        mail.Cc = "fff@asg.ru"
        mail.Subject ='Отчет ДТ ' + today_string2
        mail.Body = f"""Добрый день!  Отчет ДТ за {mes} во вложении
Просьба направить всем заинтересованным лицам

  
Если возникнут вопросы - обращайтесь
С уважением, Иван"""
        mail.Attachments.Add(Source=attachment_path)
        # Используйте это, чтобы показать электронную почту
        mail.Display(True)
        # Раскомментировать для отправки
        #mail.Send()

email_sender = EmailsSender()
for index, row in combined.iterrows():
    email_sender.send_email(row['EMAIL'], row['FILE'])