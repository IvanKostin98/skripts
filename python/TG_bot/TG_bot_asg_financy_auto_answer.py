# -*- coding: utf-8 -*-
"""
Created on Tue Jun 27 16:40:27 2023

@author: kostin

Надо привязать бота к кубу и подтягиваь актуальную информацию о ДМ и ТМ
"""

import telebot
from my_functions_only_for_import import logging_on_bd
first = logging_on_bd(username = 'Иван')

first.ignore_warnings()
all_log_table = first.bd_select_loggs_result(priznak = False)[:15]

TOKEN = '599*****************4p07mZ6***********plU'
bot=telebot.TeleBot('5*****28:AAG*****************OrkxMp***')
chat_id = '945394408'
message_text = '\n'.join(all_log_table)
 
bot.send_message(chat_id=chat_id, text=message_text)
