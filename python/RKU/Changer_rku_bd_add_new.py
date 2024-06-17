# -*- coding: utf-8 -*-
"""
Created on Mon Mar 18 10:32:48 2024

Авто добавление новых клиентов и sku из РКУ в бд
>>>>>>>>Не забывай менять начальную дату ввода новинки!!!!!<<<<<<<<
@author: kostin ivan
"""
import pandas as pd
import pyodbc
import datetime
import warnings
import random

warnings.filterwarnings("ignore")

def file_to_df(path, sheet, skiprows, columns):
    file = pd.read_excel(path, sheet_name = sheet, skiprows=skiprows)
    df = pd.DataFrame(file, columns = columns)
    return df

def func_for_find_difference(query, column, df_exl, type_x, flag = True):
    """функция для нахождения разницы между бд и РКУ эксель
        flag - служит чтобы проверить находим совпадения или различия"""
    sp = list()
    df_sql = pd.read_sql(query, conn)
    df_sql = list(map(lambda x: type_x(x), df_sql[column].unique()))
    df_exl = list(map(lambda x: type_x(x), df_exl[column].unique()))
    for i in df_exl:
        if i not in df_sql and flag == True:        #находит разницу бд и экселя
            sp.append(i)
        if i in df_sql and flag == False:           #находит совпадения в РКУ и бд по топ 5 стандартным ДТ из бд
            sp.append(i) 
    return sp
 
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

#Подключение к бд/загрузка файла
conn = pyodbc.connect(driver='{SQL Server}',
                      server="***", 
                      database="FinanceAndSAP",               
                      trusted_connection="yes")
cur = conn.cursor()
df_exl = file_to_df(path_to_rku_excel, 'Загрузчик КУ в 1С', 4, columns = sp_columns)
df_exl.rename(columns = {'Код SKU':'Код_SKU', 'Код сегмента':'Код_сегмента', 'BI название ДТ':'BI_NameDT', 'Название сегмента':'SegmentName',
                         'ДТ с ндс':'ДТ_с_ндс_стандарт', 'ТТсндс':'ТТ_с_ндс_стандарт', 'МТсндс':'МТ_с_ндс_стандарт',
                         'ОТсндс':'ОТ_с_ндс_стандарт', 'ОТ/ТТ':'ОТ_ТТ_стандарт', 'наценка ДТ':'Наценка_ДТ_стандарт',
                         'РБ':'РБ_стандарт', 'вход с уч РБ':'Вход_с_уч_РБ_стандарт', 'наценка ДТ с РД':'Наценка_ДТ_с_РБ_стандарт',
                         'Название SKU':'Название_SKU', 'Регулярная цена без НДС':'Регулярная_цена_без_НДС', 'Регулярная цена с НДС':'Регулярная_цена_с_НДС', 'Бренд':'Бренд'
                         
                      }, inplace = True)
df_exl.drop(index=df_exl.index [0], axis= 0, inplace= True)

query_sku = """ SELECT SKU_ID as Код_SKU
                FROM [FinanceAndSAP].[segment].[Agreements]
                WHERE TillDate = '2024-12-31 00:00:00.000'
                GROUP BY SKU_ID """

query_client = """  SELECT SegmentCode as Код_сегмента
                    FROM [FinanceAndSAP].[segment].[Agreements]
                    WHERE TillDate = '2024-12-31 00:00:00.000'
                    GROUP BY SegmentCode """
                    
query_example_dt = """  SELECT TOP (5) SegmentCode as Код_сегмента
        				FROM [FinanceAndSAP].[segment].[Agreements]
					    WHERE TillDate = '2024-12-31 00:00:00.000' and EDLP = 0 and StandartTT = IndividTT
					    GROUP BY SegmentCode
					    ORDER BY count(SKU_ID) DESC """
        
diff_sku = func_for_find_difference(query_sku, 'Код_SKU', df_exl, type_x = int)
diff_client = func_for_find_difference(query_client, 'Код_сегмента', df_exl, type_x = str)
one_top_client = random.choice(func_for_find_difference(query_example_dt, 'Код_сегмента', df_exl, type_x = str, flag = False))

print(f'{diff_sku} нет в бд' if len(diff_sku)>0 else 'SKU актуализированы')
print(f'{diff_client} нет в бд' if len(diff_client)>0 else 'Клиенты актуализированы')

for i in diff_client:
    df_exl_filtred = df_exl[df_exl['Код_сегмента'] == i].iloc[:1]
    res = dict(*(df_exl_filtred[['Код_сегмента', 'Дивизион', 'BI_NameDT', 'SegmentName', 'Бренд']].groupby("Код_сегмента").apply(lambda x: x.drop(columns="Код_сегмента").to_dict("records")).to_dict()).get(i))     #Датафрейм из эксель файла фильтранул по столбцам и коду сегмента преобразовал в словарь
    cur.execute(f""" INSERT INTO [FinanceAndSAP].[segment].[Agreements]
                            SELECT [ProcID]
                                  ,[SKU_ID]
                                  ,[SKU]
                                  ,[SegmentCode] = '{i}'
                                  ,[BrandID]
                                  ,[Brand] = '{res.get('Бренд')}'
                                  ,[BI_NameDT] = '{res.get('BI_NameDT')}'
                                  ,[DivisionID] = '{0}'
                                  ,[Division] = '{res.get('Дивизион')}'
                                  ,[SegmentName] = '{res.get('SegmentName')}'
                                  ,[Vol]
                                  ,[FromDate]
                                  ,[TillDate]
                                  ,[GS]
                                  ,[PIAD]
                                  ,[TH]
                                  ,[PD]
                                  ,[TMD]
                                  ,[LD]
                                  ,[EDLP]
                                  ,[Reb]
                                  ,[PriceNoVAT]
                                  ,[Price]
                                  ,[Pickup]
                                  ,[Prepay]
                                  ,[PriceNoPickupAndPrepay]
                                  ,[MOC]
                                  ,[StandartDT]
                                  ,[StandartTT]
                                  ,[StandartMT]
                                  ,[StandartOT]
                                  ,[StandartOT_TT]
                                  ,[StandartExtraDT]
                                  ,[StandartRB]
                                  ,[StandartEntryRB]
                                  ,[StandartExtraDT_RB]
                                  ,[IndividDT]
                                  ,[IndividTT]
                                  ,[IndividMT]
                                  ,[IndividOT]
                                  ,[IndividOT_TT]
                                  ,[IndividExtraDT]
                                  ,[IndividRB]
                                  ,[IndividEntryRB]
                                  ,[IndividExtraDT_RB]
                                  ,[Dal]
                                  ,[SellinBaseGrossSales]
                                  ,[SellinGrossSales]
                                  ,[InclRebate]
                                  ,[InclGuaranteedYield]
                                  ,[InclPartnershipDiscount]
                                  ,[InclPrepaymentDiscount]
                                  ,[InclPickUpDiscount]
                                  ,[InclPartnershipDiscountTT]
                                  ,[InclListingDiscount]
                                  ,[inclMotivationDiscountSPSR]
                                  ,[InclMerchandising]
                                  ,[InclDiscountPromotion]
                                  ,[Excise]
                                  ,[NetSales]
                                  ,[TotalMaterialCosts]
                                  ,[OtherProduction]
                                  ,[WarehousingCosts]
                                  ,[FixClientsLogistic]
                                  ,[TotalSupplyIndirectCosts]
                                  ,[BrandMarketing]
                                  ,[InclVariableProduction]
                                  ,[VariableIntercompanyLogistics]
                                  ,[VariableClientLogistic]
                                  ,[Contribution0]
                                  ,[Contribution1]
                                  ,[PlanVersion]
                                  ,[DescriptionTT]
                                  ,[DescriptionMT]
                                  ,[DescriptionOT]
                                  ,[AttrPrice]
                                  ,[AttrCS]
                                  ,[DM] = 'text'
                                  ,[TM] = 'text'
                                  ,[ClientService] = 'text'
                                  ,[OPK] = 'text'
                                  ,[isDeleted]
                                  ,[isFromDB]
                            FROM [FinanceAndSAP].[segment].[Agreements]
                            WHERE TillDate = '2024-12-31 00:00:00.000' and SegmentCode = '{one_top_client}' """)

for i in diff_sku:
    df_exl_filtred = df_exl[df_exl['Код_SKU'] == i].iloc[:1]
    res = dict(*(df_exl_filtred[['Код_SKU', 'Название_SKU', 'Емкость', 'Цена', 'Регулярная_цена_без_НДС', 'Регулярная_цена_с_НДС', 'ЦЕНА', 'МОЦ',
                                 'Регулярная_цена_без_НДС', 'РБ_стандарт']].groupby("Код_SKU").apply(lambda x: x.drop(columns="Код_SKU").to_dict("records")).to_dict()).get(i))    #Датафрейм из эксель файла фильтранул файл по столбцам и скю преобразовал в словарь
    cur.execute(f"""INSERT INTO [FinanceAndSAP].[segment].[Agreements]
                          ([ProcID]
                                  ,[SKU_ID] 
                                  ,[SKU] 
                                  ,[SegmentCode]
                                  ,[BrandID]
                                  ,[Brand] 
                                  ,[BI_NameDT]
                                  ,[DivisionID]
                                  ,[Division]
                                  ,[SegmentName]
                                  ,[Vol] 
                                  ,[FromDate] 
                                  ,[TillDate] 
                                  ,[GS] 
                                  ,[PIAD]
                                  ,[TH]
                                  ,[PD]
                                  ,[TMD]
                                  ,[LD]
                                  ,[EDLP] 
                                  ,[Reb]
                                  ,[PriceNoVAT] 
                                  ,[Price] 
                                  ,[Pickup]
                                  ,[Prepay]
                                  ,[PriceNoPickupAndPrepay]
                                  ,[MOC] 
                                  ,[StandartDT] 
                                  ,[StandartTT] 
                                  ,[StandartMT] 
                                  ,[StandartOT] 
                                  ,[StandartOT_TT] 
                                  ,[StandartExtraDT] 
                                  ,[StandartRB] 
                                  ,[StandartEntryRB]
                                  ,[StandartExtraDT_RB] 
                                  ,[IndividDT]
                                  ,[IndividTT]
                                  ,[IndividMT]
                                  ,[IndividOT]
                                  ,[IndividOT_TT] 
                                  ,[IndividExtraDT] 
                                  ,[IndividRB] 
                                  ,[IndividEntryRB] 
                                  ,[IndividExtraDT_RB]
                                  ,[Dal] 
                                  ,[SellinBaseGrossSales]
                                  ,[SellinGrossSales] 
                                  ,[InclRebate]
                                  ,[InclGuaranteedYield] 
                                  ,[InclPartnershipDiscount] 
                                  ,[InclPrepaymentDiscount] 
                                  ,[InclPickUpDiscount] 
                                  ,[InclPartnershipDiscountTT]
                                  ,[InclListingDiscount] 
                                  ,[inclMotivationDiscountSPSR] 
                                  ,[InclMerchandising] 
                                  ,[InclDiscountPromotion] 
                                  ,[Excise] 
                                  ,[NetSales] 
                                  ,[TotalMaterialCosts] 
                                  ,[OtherProduction]
                                  ,[WarehousingCosts] 
                                  ,[FixClientsLogistic] 
                                  ,[TotalSupplyIndirectCosts] 
                                  ,[BrandMarketing] 
                                  ,[InclVariableProduction] 
                                  ,[VariableIntercompanyLogistics] 
                                  ,[VariableClientLogistic] 
                                  ,[Contribution0] 
                                  ,[Contribution1] 
                                  ,[PlanVersion] 
                                  ,[DescriptionTT] 
                                  ,[DescriptionMT] 
                                  ,[DescriptionOT] 
                                  ,[AttrPrice] 
                                  ,[AttrCS] 
                                  ,[DM] 
                                  ,[TM] 
                                  ,[ClientService] 
                                  ,[OPK] 
                                  ,[isDeleted]
                                  ,[isFromDB])
                            SELECT [ProcID]
                                  ,[SKU_ID] = '{i}'
                                  ,[SKU] = '{res.get('Название_SKU')}'
                                  ,[SegmentCode]
                                  ,[BrandID] = {0}
                                  ,[Brand]
                                  ,[BI_NameDT]
                                  ,[DivisionID]
                                  ,[Division]
                                  ,[SegmentName]
                                  ,[Vol] = '{res.get('Емкость')}'
                                  ,[FromDate] = '2024-06-10 00:00:00'
                                  ,[TillDate] = '2024-12-31 00:00:00'
                                  ,[GS] = '{0}'
                                  ,[PIAD]
                                  ,[TH]
                                  ,[PD]
                                  ,[TMD]
                                  ,[LD]
                                  ,[EDLP] = '{0}'
                                  ,[Reb]
                                  ,[PriceNoVAT] = '{0}'
                                  ,[Price] = '{0}'
                                  ,[Pickup] = '{0}'
                                  ,[Prepay] = '{0}'
                                  ,[PriceNoPickupAndPrepay] = '{0}'
                                  ,[MOC] = '{0}'
                                  ,[StandartDT] = '{0}'
                                  ,[StandartTT] = '{0}'
                                  ,[StandartMT] = '{0}'
                                  ,[StandartOT] = '{0}'
                                  ,[StandartOT_TT] = '{0}'
                                  ,[StandartExtraDT] = '{0}'
                                  ,[StandartRB] = '{res.get('РБ_стандарт')}'
                                  ,[StandartEntryRB] = '{0}'
                                  ,[StandartExtraDT_RB] = '{0}'
                                  ,[IndividDT] = '{0}'
                                  ,[IndividTT] = '{0}'
                                  ,[IndividMT] = '{0}'
                                  ,[IndividOT] = '{0}'
                                  ,[IndividOT_TT] = '{0}'
                                  ,[IndividExtraDT] = '{0}'
                                  ,[IndividRB] = '{res.get('РБ_стандарт')}'
                                  ,[IndividEntryRB] = '{0}'
                                  ,[IndividExtraDT_RB] = '{0}'
                                  ,[Dal] = '{0}'
                                  ,[SellinBaseGrossSales] = '{0}'
                                  ,[SellinGrossSales] = '{0}'
                                  ,[InclRebate] = '{0}'
                                  ,[InclGuaranteedYield] = '{0}'
                                  ,[InclPartnershipDiscount] = '{0}'
                                  ,[InclPrepaymentDiscount] = '{0}'
                                  ,[InclPickUpDiscount] = '{0}'
                                  ,[InclPartnershipDiscountTT] = '{0}'
                                  ,[InclListingDiscount] = '{0}'
                                  ,[inclMotivationDiscountSPSR] = '{0}'
                                  ,[InclMerchandising] = '{0}'
                                  ,[InclDiscountPromotion] = '{0}'
                                  ,[Excise] = '{0}'
                                  ,[NetSales] = '{0}'
                                  ,[TotalMaterialCosts] = '{0}'
                                  ,[OtherProduction] = '{0}'
                                  ,[WarehousingCosts] = '{0}'
                                  ,[FixClientsLogistic] = '{0}'
                                  ,[TotalSupplyIndirectCosts] = '{0}'
                                  ,[BrandMarketing] = '{0}'
                                  ,[InclVariableProduction] = '{0}'
                                  ,[VariableIntercompanyLogistics] = '{0}'
                                  ,[VariableClientLogistic] = '{0}'
                                  ,[Contribution0] = '{0}'
                                  ,[Contribution1] = '{0}'
                                  ,[PlanVersion] = 'new_sku'
                                  ,[DescriptionTT] = 'РФ'
                                  ,[DescriptionMT] = 'РФ'
                                  ,[DescriptionOT] = 'РФ'
                                  ,[AttrPrice] = 'Стандартный'
                                  ,[AttrCS] = 'Стандартный'
                                  ,[DM] = 'text'
                                  ,[TM] = 'text'
                                  ,[ClientService] = 'text'
                                  ,[OPK] = 'text'
                                  ,[isDeleted]
                                  ,[isFromDB]
                        FROM [FinanceAndSAP].[segment].[Agreements]
                        WHERE TillDate = '2024-12-31 00:00:00' and SKU_ID = '10006000010100' """) # пять озёр 0.5 как основа
conn.commit()
conn.close()
