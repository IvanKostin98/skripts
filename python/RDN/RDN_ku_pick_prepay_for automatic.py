import pandas as pd
import pyodbc
import datetime
from my_functions_only_for_import import logging_on_bd

sp_in_bd = ['Дивизион', 'код_1С', 'Клиент', 'L6', 'sku', 'c', 'по', 'Прайс_ДТ_с_НДС', 'Цена_отгрузки_с_НДС', 'Управленческая_цена_с_НДС', 'КУ_с_НДС', 'КУ_без_НДС', 'самовывоз_процент', 'самовывоз_с_ндс', 'самовывоз_без_ндс', 'предоплата_процент', 'предоплата_с_ндс', 'предоплата_без_ндс', 'самовывоз_и_предоплата_с_НДС', 'самовывоз_и_предоплата_без_НДС', 'Мероприятие', 'Комментарии']                                          
sp_in_table = ['Дивизион', 'код 1С', 'Клиент', 'L6', 'sku', 'c', 'по', 'прайс ДТ', 'Цена отгрузки', 'управленческая цена', 'КУ (статья rebate)', 'КУ (статья rebate).1', 'самовывоз', 'самовывоз.1', 'самовывоз.2', 'предоплата', 'предоплата.1', 'предоплата.2', 'самовывоз и предоплата', 'самовывоз и предоплата.1', 'Мероприятие', 'Комментарии']

dict_rebate = dict(zip(sp_in_bd, sp_in_table))

log_file = r'\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\Костин И.Н\Python скрипты\Логирование\Лог_файл_для_РДН.txt'
df_ku=pd.read_excel(r"\\fs1v-msk01\Для всех\Третьякова М\Дистрибьюторы спец.прайс\!Реестр дополнительных начислений ЧС.xlsm", sheet_name="КУ_самовывоз", skiprows=0)
df_ku=pd.DataFrame(df_ku, columns=sp_in_table)
df_ku.drop(labels = [0,1],axis = 0, inplace=True)

now = datetime.datetime.now()
first = logging_on_bd(username = 'Ivan')

def znach_from_dict(x, nomer, dictionary, df_kakoy):
    k = dictionary.get(x)
    fin = df_kakoy.iloc[nomer][k]
    if pd.isnull(fin)==True or fin == 'None' or fin == None or fin =='NONE':
        fin=None
    return fin

cr__df_ku = df_ku.shape[0]
try:
    # подключение к БД
    conn = pyodbc.connect(driver='{SQL Server}',
                          server="server219", 
                          database="Sandbox",               
                          trusted_connection="yes")
    cur = conn.cursor()
    
    cur.execute("""TRUNCATE TABLE kostin.rdn_ku_pick_prepay""")
    
    for i in range(0, cr__df_ku):
           
        cur.execute(f"""INSERT INTO kostin.rdn_ku_pick_prepay (ID,
                                                        Дивизион,
                                                        код_1С,
                                                        Клиент,
                                                        L6,
                                                        sku,
                                                        c,
                                                        по,
                                                        Прайс_ДТ_с_НДС,
                                                        Цена_отгрузки_с_НДС,
                                                        Управленческая_цена_с_НДС,
                                                        КУ_с_НДС,
                                                        КУ_без_НДС,
                                                        самовывоз_процент,
                                                        самовывоз_с_ндс,
                                                        самовывоз_без_ндс,
                                                        предоплата_процент,
                                                        предоплата_с_ндс,
                                                        предоплата_без_ндс,
                                                        самовывоз_и_предоплата_с_НДС,
                                                        самовывоз_и_предоплата_без_НДС,
                                                        Мероприятие,
                                                        Комментарии,
                                                        Дата_обновления,
                                                        Кто_обновил
                                                        )
                    VALUES ({i+1},
                            case when '{znach_from_dict('Дивизион', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('Дивизион', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('код_1С', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('код_1С', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('Клиент', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('Клиент', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('L6', i, dict_rebate, df_ku)}'='None' then NULL else {int(znach_from_dict('L6', i, dict_rebate, df_ku))} end,
                            case when '{znach_from_dict('sku', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('sku', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('c', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('c', i, dict_rebate, df_ku).strftime('%Y-%m-%d')}' end,
                            case when '{znach_from_dict('по', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('по', i, dict_rebate, df_ku).strftime('%Y-%m-%d')}' end,
                            case when '{znach_from_dict('Прайс_ДТ_с_НДС', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('Прайс_ДТ_с_НДС', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('Цена_отгрузки_с_НДС', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('Цена_отгрузки_с_НДС', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('Управленческая_цена_с_НДС', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('Управленческая_цена_с_НДС', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('КУ_с_НДС', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('КУ_с_НДС', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('КУ_без_НДС', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('КУ_без_НДС', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('самовывоз_процент', i, dict_rebate, df_ku)}' = 'None' then NULL else '{znach_from_dict('самовывоз_процент', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('самовывоз_с_ндс', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('самовывоз_с_ндс', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('самовывоз_без_ндс', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('самовывоз_без_ндс', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('предоплата_процент', i, dict_rebate, df_ku)}' = 'None' then NULL else '{znach_from_dict('предоплата_процент', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('предоплата_с_ндс', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('предоплата_с_ндс', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('предоплата_без_ндс', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('предоплата_без_ндс', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('самовывоз_и_предоплата_с_НДС', i, dict_rebate, df_ku)}' = 'None' then NULL else '{znach_from_dict('самовывоз_и_предоплата_с_НДС', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('самовывоз_и_предоплата_без_НДС', i, dict_rebate, df_ku)}' = 'None' then NULL else '{znach_from_dict('самовывоз_и_предоплата_без_НДС', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('Мероприятие', i, dict_rebate, df_ku)}' = 'None' then NULL else '{znach_from_dict('Мероприятие', i, dict_rebate, df_ku)}' end,
                            case when '{znach_from_dict('Комментарии', i, dict_rebate, df_ku)}'='None' then NULL else '{znach_from_dict('Комментарии', i, dict_rebate, df_ku)}' end,
                            '{now.strftime('%Y-%m-%d')}',
                            '{'Pyth_kostin'}')
                            """)
    
    conn.commit()
    conn.close()
    print("всё окей")
    first.bd_open()
    first.bd_process(filename = 'RDN', processname = 'КУ компенсации', result = '1', comments = 'загружено')
    first.bd_close()
except:
    print("чёт не то, проверь логи")
    first.bd_open()
    first.bd_process(filename = 'RDN', processname = 'КУ компенсации', result = '0', comments = 'ошибка')
    first.bd_close()
