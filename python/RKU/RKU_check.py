# -*- coding: utf-8 -*-
"""
Created on Wed May 22 12:37:12 2024

@author: kostin
"""

from my_functions_only_for_import import logging_on_bd, usefull_function, excel

log_class = logging_on_bd('Иван')
excel_class = excel('Иван')
usefull_class = usefull_function()


log_class.ignore_warnings()         #игнорим предупреждения

                   
path_to_rku_excel = r'\\fs1v-msk01\Для всех\Третьякова М\Pricing\Рита\РКУ\РКУ 2024.xlsx'
sp_columns = ['Код SKU', 'Код сегмента', 'Бренд', 'BI название ДТ', 'Дивизион', 'Название сегмента', 'Название SKU', 
              'Емкость', 'Срок с', 'Срок по', 'Цена', 'скидка за непринятие', 'скидка для торгового дома', 
              'Partnership discount MT текущий', 'Trade Marketing discount conditional(incl CLAP) текущий', 
              'Logistics discount текущий', 'EDLP текущий', 'Reb', 'Регулярная цена без НДС', 'Регулярная цена с НДС', 
              'самовывоз', 'предоплата', 'ЦЕНА', 'МОЦ', 'ДТ с ндс', 'ТТсндс', 'МТсндс', 'ОТсндс', 'ОТ/ТТ', 'наценка ДТ', 
              'РБ', 'вход с уч РБ', 'наценка ДТ с РД', 'ДТ с ндс.1', 'ТТ с ндс', 'МТ с ндс', 'ОТ с ндс', 'ОТ/ТТ.1', 
              'наценка ДТ.1', 'вход с уч РБ.1', 'наценка ДТ с РД.1', 'ТТ', 'МТ', 'ОТ']


RKU_from_bd = log_class.bd_select("""SELECT *
                        FROM [FinanceAndSAP].[segment].[Agreements]
                        WHERE TillDate = '2024-12-31 00:00:00.000'""")

RKU_from_excel = excel_class.loading_excel(path_to_rku_excel, 'Загрузчик КУ в 1С', 4, sp_columns)
RKU_from_excel = RKU_from_excel.dropna(subset=['Код SKU'])

class check_rku:
    def __init__(self, bd, exl):
        self.bd = bd
        self.exl = exl
    
    def check_colvo(self, bd_colname, exl_colname):
        unique_bd = len(set(self.bd[bd_colname].dropna().tolist()))
        unique_exl = len(set(self.exl[exl_colname].dropna().tolist()))
        bool_result = unique_bd == unique_exl
        
        return f"{bool_result} {unique_bd}" if bool_result else f"{bool_result} В BD: {unique_bd} В EXL: {unique_exl}"

    def check_unique(self, bd_colname, exl_colname, types = True):
        """Функция проверяет равенство уникальных значний в бд и эксель-файле"""
        
        bd_unique = set(self.bd[bd_colname].dropna().tolist())
        excel_unique = set(self.exl[exl_colname].dropna().tolist())
        if types == False:
            razn = set([int(i) for i in bd_unique]) - set([int(i) for i in excel_unique])
            if len(razn)!=0:
                return razn
            else:
                return True
        elif bd_unique == excel_unique:
            return True
        else:
            return bd_unique - excel_unique
    
    def check_column(self, bd, exl):
        """Функция проверяет данные по сцепке клиент и L6 на соответствие РКУ"""
        
        import pandas as pd
        
        merge_df = pd.merge(self.bd, self.exl, left_on = ['SegmentCode', 'SKU_ID'], right_on = ['Код сегмента', 'Код SKU'], how = 'left')
        merge_df['diff'] = merge_df[bd] - merge_df[exl]
        col_sum = merge_df['diff'].sum()
        result = True if -1 < col_sum < 1 else f"{False} {col_sum}"
        return result

check = check_rku(RKU_from_bd, RKU_from_excel)

print(">>>Тесты на колличество<<<")
print(f"Test: {check.check_colvo('SegmentCode', 'Код сегмента')} Код сегмента")
print(f"Test: {check.check_colvo('SKU_ID', 'Код SKU')} L6")
print()
print(">>>Тесты на актуальность справочников<<<")
print(f"Test: {check.check_unique('SegmentCode', 'Код сегмента')} Код сегмента")
print(f"Test: {check.check_unique('SKU_ID', 'Код SKU', types = False)} L6")
print(f"Test: {check.check_unique('Division', 'Дивизион')} Дивизион")
print(f"Test: {check.check_unique('Brand', 'Бренд')} Бренд")
print(f"Test: {check.check_unique('BI_NameDT', 'BI название ДТ')} Наименование BI")
print(f"Test: {check.check_unique('SegmentName', 'Название сегмента')} Наименование сегмента")
print(f"Test: {check.check_unique('SKU', 'Название SKU')} Наименование СКЮ")
print()
print(">>>Тесты на корректность данных<<<")
print(f"Test: {check.check_column('GS', 'Цена')} Базовая цена")
print(f"Test: {check.check_column('PriceNoVAT', 'Регулярная цена без НДС')} net-цена")
print(f"Test: {check.check_column('Pickup', 'самовывоз')} Самовывоз")
print(f"Test: {check.check_column('Prepay', 'предоплата')} Предоплата")
print(f"Test: {check.check_column('MOC', 'МОЦ')} МОЦ")
print(f"Test: {check.check_column('StandartTT', 'ТТсндс')} TT стандарт")
print(f"Test: {check.check_column('StandartMT', 'МТсндс')} MT стандарт")
print(f"Test: {check.check_column('StandartOT', 'ОТсндс')} OT стандарт")
print(f"Test: {check.check_column('IndividTT', 'ТТ с ндс')} TT индивид")
print(f"Test: {check.check_column('IndividMT', 'МТ с ндс')} MT индивид")
print(f"Test: {check.check_column('IndividOT', 'ОТ с ндс')} OT индивид")