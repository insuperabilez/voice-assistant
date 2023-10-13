
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import configparser

def send_message(number,text):
    smtp_server = 'smtp.yandex.ru'
    smtp_port = 587
    config = configparser.ConfigParser()
    config.read('config.ini')
    smtp_username = config.get('MAIL', 'smtp_username')
    smtp_password = config.get('MAIL', 'smtp_password')
    smtp_address = config.get('MAIL', 'smtp_address')
    # Создание объекта сообщения
    message = MIMEMultipart()
    message['From'] = smtp_username
    message['To'] = smtp_address
    message['Subject'] = number

    # Добавление текстовой части письма
    body = text
    message.attach(MIMEText(body, 'plain'))

    # Отправка письма
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.send_message(message)