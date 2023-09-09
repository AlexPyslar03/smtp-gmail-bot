import smtplib # Библиотека SMTP запросов
from email.mime.text import MIMEText
from config import * # Файл конфигурации пользовательских данных
import pandas as pd

# Создание server'а обработки
server = smtplib.SMTP(host, port)  # Создание объекта загрузки
server.starttls()  # Cоединение с SMTP-сервером в режим TLS

def send_email(message, to_email): # Функция отправки сообщения
    # Процесс отправки
    try:
        server.login(from_email, password) # Вход в систему для отправки
        msg = MIMEText(message) # Формирование письма
        msg["Subject"] = msg_subject # Добавление заголовка письма
        #msg["Message"] = message  # Добавление текста письма
        server.sendmail(from_email, to_email, msg.as_string()) # Отправка письма
        return "Письмо отправленно на почту " + to_email # Возвращение сообщения об успехе
    except Exception as _ex: # Обработка исключений
        return "!Ошибка при отправки на почту " + to_email + " проверьте логин или пароль"

def main():
    xlsx = pd.read_excel(xlsx_file_name, sheet_name=xlsx_sheet_name)
    #for i in range(0, xlsx.values.itemsize - 1):
        #if (xlsx.values.item(i) != ""):
            #print(i)
            #print(send_email(message=msg_text, to_email=xlsx.values.item(i)))
    send_email(message=msg_text, to_email="alexpyslar03@yandex.ru")

if __name__ == "__main__":
    main()