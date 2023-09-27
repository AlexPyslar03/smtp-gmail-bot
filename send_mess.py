import smtplib # SMTP запросы
from email.message import EmailMessage # Формирование писмьма

class Messenger(): # Класс отправщика
    def __init__(self, host, port, from_email, password): # Конструктор класса
        # Создание server'а обработки
        self.server = smtplib.SMTP(host, port)  # Создание объекта загрузки
        self.server.starttls()  # Cоединение с SMTP-сервером в режим TLS
        self.server.login(from_email, password)  # Вход в систему для отправки
        self.msg['From'] = from_email  # Указание почты откуда идёт отправка
    def create_msg(self, msg_subject, msg_text): # Функция создания письма
        self.msg = EmailMessage()  # Создание объекта письма
        self.msg['Subject'] = msg_subject  # Добавление заголовка письма
        self.msg.set_content(msg_text)  # Добавление текста письма
    def send_msg(self, to_email): # Функция отправки письма
        self.msg['To'] = to_email  # Указание почты куда идёт отправка
        self.server.sendmail(self.from_email, to_email, self.msg.as_string())  # Отправка письма
    def __del__(self): # Деструктор класса
        self.server.quit() # Закрытия сервера запросов
