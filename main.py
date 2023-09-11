import smtplib # Библиотека SMTP запросов
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from config import * # Файл конфигурации пользовательских данных
import pandas as pd

# Создание server'а обработки
server = smtplib.SMTP(host, port)  # Создание объекта загрузки
server.starttls()  # Cоединение с SMTP-сервером в режим TLS

def send_email(to_email): # Функция отправки сообщения
    # Процесс отправки
    try:
        server.login(from_email, password) # Вход в систему для отправки
        msg = MIMEMultipart() # Формирование письма
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = msg_subject # Добавление заголовка письма
        #msg.attach(MIMEText(msg_text, 'plain')) # Добавление текста письма
        msg.attach(MIMEImage(file("thumbnails1.jpg").read()))
        server.sendmail(from_email, to_email, msg.as_string()) # Отправка письма
        return "Письмо отправленно на почту " + to_email # Возвращение сообщения об успехе
    except Exception as _ex: # Обработка исключений
        return "!Ошибка при отправки на почту " + to_email + " проверьте логин или пароль"

xlsx = pd.read_excel(xlsx_file_name, sheet_name=xlsx_sheet_name)
print(send_email('alexpyslar03@yandex.ru'))
#for i in range(0, xlsx.values.size):
    #if (str(xlsx.values.item(i)) != "nan"):
        #print(str(i + 2) + "-ый элемент в таблицы: " + send_email(to_email=xlsx.values.item(i)))

server.quit()