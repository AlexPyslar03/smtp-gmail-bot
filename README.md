# Бот для отправки электроных писем
## Написан [Пысларь Александром](https://t.me/alex_pyslar), студентов Московского Авиационого Института
## Бот позволяет делать массовую рассылку на почты из базы данных в Excel файле
## Инструкции по запуску:
### 1. Установите [python](https://www.python.org) с оффициального сайта
### 2. Установите билиотеки через pip:
`pip install smtplib pandas email`
### 2. Сгенерируйте пароль приложение своей почты и вставьте его в конфигурационный файл
[Инструкция по получения пароля](https://support.google.com/mail/answer/185833?hl=ru)

`from_email = "" # Почта откуда отправляется письмо` - сюда ставим почту откуда отправлять
`password = "" # Ключь пароль почты` - сюда полученный пароль
### 3. В конфигурационом файле находятся все параметры, меня их можно менять письмо
### 4. Создаём Excel таблицу с именем который указан в конфигурационом файле: `xlsx_file_name`
!Внимание - название листа должно совпадать с `xlsx_sheet_name`
### 5. В Excel файле ставим почты (первая строка должна быть пустая)
### 6. В папку `files` поместите все файлы, которые должны вкладываться в письмо 
### 7. Запускаем файл main.py
