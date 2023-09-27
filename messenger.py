import os # Библиотека для работы с ос
import smtplib # SMTP запросы
from email.message import EmailMessage # Формирование писмьма

class Messenger(): # Класс отправщика
    def __init__(self, host, port, from_email, password): # Конструктор класса
        self.host = host # Сохранение хоста
        self.port = port # Сохранение порта
        self.password = password # сохранение пароля
        # Создание server'а обработки
        try:
            self.server = smtplib.SMTP(host, port)  # Создание объекта загрузки
            print("Сервер создан на хосте: " + str(host) + " и порте: " + str(port))
        except: print("Ошибка: порта: " + str(host) + " или хоста: " + str(port))
        try:
            self.server.starttls()  # Cоединение с SMTP-сервером в режим TLS
            print("Cоединение с SMTP-сервером в режим TLS прошло успешно")
        except: print("Ошибка: cоединение с SMTP-сервером в режим TLS")
        try:
            self.server.login(from_email, password)  # Вход в систему для отправки
            print("Вход в систему на почту " + from_email + " с паролем " + password + " прошёл успешно")
        except: print("Ошибка входа в систему, проверьте почту: " + str(from_email) + " или пароль: " + str(password))
        self.from_email = from_email # Сохранение почты откуда идёт отправка
    def create_msg(self, msg_subject, msg_text, file_directory): # Функция создания письма
        self.msg = EmailMessage()  # Создание объекта письма
        self.msg['From'] = self.from_email  # Указание почты откуда идёт отправка
        self.msg['Subject'] = msg_subject  # Добавление заголовка письма
        self.msg_subject = msg_subject
        self.msg.set_content(msg_text)  # Добавление текста письма
        self.msg_text = msg_text
        for filename in os.listdir(file_directory):  # Прохождение по файлам директории
            f = os.path.join(file_directory, filename)  # Открытие файла
            if os.path.isfile(f):  # Проверка на существование файла
                with open(f, 'rb') as f:  # Открытие свойств файла
                    image_data = f.read()  # Считывание данных файла
                    image_name = f.name  # Считывание имени файла
                if (image_name.endswith('.jpg')):
                    image_name = image_name[(len(file_directory) + 1):]
                    self.msg.add_attachment(image_data, maintype='image', subtype='jpg', filename=image_name)  # Добавление вложений
        self.file_directory = file_directory
    def send_msg(self, to_email): # Функция отправки письма
        try:
            self.create_msg(self.msg_subject, self.msg_text, self.file_directory)
            self.msg['To'] = to_email  # Указание почты куда идёт отправка
            self.server.sendmail(self.from_email, to_email, self.msg.as_string())  # Отправка письма
            print("Письмо на почту: " + to_email + " отправленно успешно")
        except:
            print("Ошибка отправки письма на почту: " + to_email)
    def __del__(self): # Деструктор класса
        try:
            self.server.quit() # Закрытия сервера запросов
            print("Сервер выключен")
        except:
            print("Ошибка выключения сервера")
