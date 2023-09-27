from config import * # Интегрирование конфигурационного файла
from messenger import Messenger # Интегрирование отправщика сообщений
from xlsx_parser import XlsxParser # Интегрирование парсера xlsx

msg = Messenger(host, port, from_email, password)
msg.create_msg(msg_subject, msg_text, file_directory)

xlsx = XlsxParser(xlsx_file_name, xlsx_sheet_name)
xlsx.send_msg(msg, time_sleep)

del msg