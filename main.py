import smtplib # Библиотека SMTP запросов
import os # Библиотека для работы с ос
from email.message import EmailMessage
from config import * # Файл конфигурации пользовательских данных
import pandas as pd # Библиотека для работы с массивами
import time # Библиотека для работы со временем

# Создание server'а обработки
server = smtplib.SMTP(host, port)  # Создание объекта загрузки
server.starttls()  # Cоединение с SMTP-сервером в режим TLS
server.login(from_email, password) # Вход в систему для отправки

# Формирования письма
msg = EmailMessage()  # Создание объекта письма
msg['From'] = from_email # Указание почты откуда идёт отправка
msg['Subject'] = msg_subject # Добавление заголовка письма
msg.set_content(msg_text) # Добавление текста письма
for filename in os.listdir(file_directory): # Прохождение по файлам директории
    f = os.path.join(file_directory, filename) # Открытие файла
    if os.path.isfile(f): # Проверка на существование файла
        with open(f, 'rb') as f: # Открытие свойств файла
            image_data = f.read() # Считывание данных файла
            image_name = f.name # Считывание имени файла
        msg.add_attachment(image_data, maintype='image', subtype='jpg', filename=image_name) # Добавление вложений

def send_email(to_email): # Функция отправки сообщения
    try: # Процесс отправки
        msg['To'] = to_email  # Указание почты куда идёт отправка
        server.sendmail(from_email, to_email, msg.as_string()) # Отправка письма
        return "Письмо отправленно на почту " + to_email # Возвращение сообщения об успехе
    except Exception as _ex: # Обработка исключений
        return "!Ошибка при отправки на почту " + to_email + " проверьте логин или пароль"

# Отправка писем на каждую почту из БД
xlsx = pd.read_excel(xlsx_file_name, sheet_name=xlsx_sheet_name) # Открытие Excel файла
for i in range(0, xlsx.values.size): # Прохождения по ячейкам таблицы
    if (str(xlsx.values.item(i)) != "nan"): # Обход пустых ячеек
        print(str(i + 1) + "-ый элемент в таблицы: " + send_email(to_email=xlsx.values.item(i))) # Отправка
        time.sleep(time_sleep) # Ожидания против спама

server.quit() # Закрытия сервера запросов