# -*- coding: utf-8 -*-
"""
Created on Tue Jun 27 16:40:27 2023

@author: kostin

Надо привязать бота к кубу и подтягиваь актуальную информацию о ДМ и ТМ
"""

import telebot
from telebot import types
from my_functions_only_for_import import logging_on_bd

first = logging_on_bd(username = 'Иван')

first.ignore_warnings()

bot=telebot.TeleBot('5997206328:AAGc1SFKcTc4p07mZ6ZserZmH93OrkxMplU')

@bot.message_handler(content_types=['text'])   
def startBot(message):
      first_mess = f"<b>{message.from_user.first_name}</b>, привет!\nВыбери кнопку для проверки отработанного скрипта"
      markup = types.InlineKeyboardMarkup()
      button_DV = types.InlineKeyboardButton(text = 'Проверить РДН', callback_data='Проверить РДН')
      button_YR = types.InlineKeyboardButton(text = 'Проверить Бекап', callback_data='Проверить Бекап')
      button_KV = types.InlineKeyboardButton(text = 'Проверить РКУ', callback_data='Проверить РКУ')
      markup.add(button_DV, button_YR, button_KV)
      bot.send_message(message.chat.id, first_mess, parse_mode='html', reply_markup=markup)

@bot.callback_query_handler(func=lambda call:True)
def response(function_call):
    if function_call.message:
        if function_call.data == "Проверить РДН":
            second_mess = '\n'.join(first.bd_select_loggs_result('RDN')[:5])
            bot.send_message(function_call.message.chat.id, second_mess)
            bot.answer_callback_query(function_call.id)
        elif function_call.data == "Проверить Бекап":
            end_mess = '\n'.join(first.bd_select_loggs_result('OneDrive')[:2])
            bot.send_message(function_call.message.chat.id, end_mess)
            bot.answer_callback_query(function_call.id)
        elif function_call.data == "Проверить РКУ":
            end_mess = '\n'.join(first.bd_select_loggs_result('RDN')[:5])
            bot.send_message(function_call.message.chat.id, end_mess)
            bot.answer_callback_query(function_call.id)

bot.polling()