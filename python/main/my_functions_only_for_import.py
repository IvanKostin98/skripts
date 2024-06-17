# -*- coding: utf-8 -*-
"""
Created on Wed Apr 24 11:44:04 2024

@author: kostin
Вызвать инфо о классе или функции можно дандер методом __doc__
"""

class logging_on_bd:
    """Класс создан для подключения к БД и логирования или проведения манипуляций с данными"""
    def __init__(self, username):
        self.username = username
        
    @staticmethod
    def ignore_warnings():
        """Игнорирование предупреждений"""
        import warnings
        warnings.filterwarnings("ignore")
        
    def bd_open(self):
        """Соединение с сервером"""
        import pyodbc

        self.conn = pyodbc.connect(driver='{SQL Server}',
                              server="***", 
                              database="Sandbox",               
                              trusted_connection="yes")
        self.cur = self.conn.cursor()
        
    def bd_close(self):
        """Разрыв соединения с сервером""" 
        self.conn.close()
                
    def bd_process(self, filename, processname, result, comments):
        """Запускает процесс добавления строк в табл"""
        import datetime
        
        now = datetime.datetime.now()
        self.cur.execute(f"""INSERT INTO Sandbox.kostin.logging_table (filename
                                                                    ,processname
                                                                    ,date_of_loading
                                                                    ,time_of_loading
                                                                    ,result
                                                                    ,comments
                                                                    ,username
                                                                     )
                                                                VALUES ('{filename}'
                                                                    ,'{processname}'
                                                                    ,'{now}'
                                                                    ,'{now.strftime('%H:%M:%S')}'
                                                                    ,{result}
                                                                    ,'{comments}'
                                                                    ,'{self.username}'
                                                                    )
                                                                """)
        self.conn.commit()
    
    def bd_truncate(self, shema):
        """Очистка табл логирования"""
        self.cur.execute("""TRUNCATE TABLE {shema}""")
        self.conn.commit()
        
    def bd_select(self, query):
        """select из бд, возвращает датафрейм"""
        import pandas as pd
        
        self.bd_open()
        res = pd.read_sql(query, self.conn)
        self.bd_close()
        
        return res
    
    def bd_select_loggs_table(self):
        """select из бд, возвращает логги в виде датафрейма"""
        
        query = 'SELECT top(50) [date_of_loading], [comments], [processname] FROM [Sandbox].[kostin].[logging_table] ORDER BY [date_of_loading] DESC'
        
        return self.bd_select(query)
    
    def bd_select_loggs_result(self, priznak=False):
        """select из бд, возвращает логги в виде отфильтрованного списка"""
        
        dict_loggs = self.bd_select_loggs_table().to_dict('records')
        if priznak is False:
            sp = [' '.join(i.values()) for i in dict_loggs]
        else:
            filtred_dict_loggs = filter(lambda x: x['filename'] == priznak, dict_loggs)
            sp = [' '.join(i.values()) for i in filtred_dict_loggs]
        
        return sp

class usefull_function:
    """Класс создан для упрощения работы, здесь есть декораторы и полезные функции"""
    
    @staticmethod
    def ignore_warnings():
        """Игнорирование предупреждений"""
        import warnings
        warnings.filterwarnings("ignore")
    
    @staticmethod
    def timer(iters=1):
        '''Декоратор на подсчет времени выполнения функции
           iters - кол-во тестов функции'''
           
        import functools
        import time
           
        def decorator(func):
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                total = 0
                for i in range(iters):
                    start = time.perf_counter()
                    value = func(*args, **kwargs)
                    end = time.perf_counter()
                    total += end - start
                print(f'Время выполнения {func.__name__}: {round(total/iters, 4)} сек.')
                return value
            return wrapper
        return decorator
    
    @staticmethod
    def check_server():
        '''Проверка есть ли доступ к папкам с сервера'''
        
        import sys

        try:
            with open(r'\\***\Для всех\***\Дистрибьюторы спец.прайс\***\Python скрипты\Используем\ДТ\Бекап файлов на облако\Доступ к серверу.txt', 'r') as file:
                file.close()
            print("Есть доступ к серверу msk")

        except:
            print("Нет доступа к серверу msk")
            sys.exit(0)

class visualization:
    """Класс создан для визуализации датафрейма"""
    
    def __init__(self, tablename, dataframe):
        import seaborn as sns
        self.tablename = tablename
        self.dataframe = dataframe
        self.sns = sns
    
    def lineplot(self, x, y):
        self.sns.lineplot(data=self.dataframe, x=x, y=y)
    
    def barplot(self, x, y):
        self.sns.barplot(data=self.dataframe, x=x, y=y)
        
class excel:
    """Класс создан для работы эксель файлами, преобразованием их в датафрейм"""

    def __init__(self, username):
        self.username = username
    
    def loading_excel(self, path, sheet, skiprows, columns):
        """Загрузка эксель файла в датафрейм"""
        import pandas as pd
        
        file = pd.read_excel(path, sheet_name = sheet, skiprows=skiprows)
        df = pd.DataFrame(file, columns = columns)
        return df
    
    def upload_to_excel(self, df, path, index=False):
       """Загрузка датафрейма в эксель файл"""
       df.to_excel(r'{path}', index=index)
        

    
    
    
    
    
    
