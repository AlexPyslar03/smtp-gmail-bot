import os # Библиотека для работы с ос
from config import * # Файл конфигурации пользовательских данных
import pandas as pd # Библиотека для работы с массивами
import time # Библиотека для работы со временем

def send_email(to_email): # Функция отправки сообщения
    try: # Процесс отправки
        for filename in os.listdir(file_directory):  # Прохождение по файлам директории
            f = os.path.join(file_directory, filename)  # Открытие файла
            if os.path.isfile(f):  # Проверка на существование файла
                with open(f, 'rb') as f:  # Открытие свойств файла
                    image_data = f.read()  # Считывание данных файла
                    image_name = f.name  # Считывание имени файла
                if (image_name.endswith('.jpg')):
                    image_name = image_name[:-4][(len(file_directory) + 1):]
                    msg.add_attachment(image_data, maintype='image', subtype='jpg', filename=image_name)  # Добавление вложений
        return "Письмо отправленно на почту " + to_email # Возвращение сообщения об успехе
    except Exception as _ex: # Обработка исключений
        return "!Ошибка при отправки на почту " + to_email + " проверьте логин или пароль"

# Отправка писем на каждую почту из БД
xlsx = pd.read_excel(xlsx_file_name, sheet_name=xlsx_sheet_name) # Открытие Excel файла
for i in range(0, xlsx.values.size): # Прохождения по ячейкам таблицы
    if (str(xlsx.values.item(i)) != "nan"): # Обход пустых ячеек
        print(str(i + 1) + "-ый элемент в таблицы: " + send_email(to_email=xlsx.values.item(i))) # Отправка
        time.sleep(time_sleep) # Ожидания против спама