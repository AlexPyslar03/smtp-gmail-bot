import os # Библиотека для работы с ос
import pandas as pd # Библиотека для работы с массивами
import time # Библиотека для работы со временем
from messenger import Messenger # Интегрирование отправщика сообщений

class XlsxParser(): # Класс парсера xlsx файлов
    def __init__(self, xlsx_file_name, xlsx_sheet_name): # Конструктор класса
        try:
            self.xlsx = pd.read_excel(xlsx_file_name, sheet_name=xlsx_sheet_name)  # Открытие Excel файла
            print("Таблица: " + xlsx_file_name + " открыта успешно")
        except:
            print("Ошибка открытия таблицы")
    def send_msg(self, msg, time_sleep): # Отправка писем на каждую почту из БД
        for i in range(0, self.xlsx.values.size):  # Прохождения по ячейкам таблицы
            if (str(self.xlsx.values.item(i)) != "nan"):  # Обход пустых ячеек
                msg.send_msg(str(self.xlsx.values.item(i))) # Отправка сообщения
                time.sleep(time_sleep) # Ожидания против спама