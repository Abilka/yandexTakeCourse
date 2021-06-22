import smtplib
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders

import time
import configparser


def mail_send(filename: str, row_amount: int) -> None:
    '''
    :param email: gmail почта
    :param password: пароль от почты выше
    :param filename: (необязательно) файл, вложенный в письмо
    '''

    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    part.add_header(
        "Content-Disposition",
        "attachment", filename=filename
    )
    encoders.encode_base64(part)
    login_data = take_data_login_gmail()

    message = MIMEMultipart()
    message["From"] = login_data[0]
    message["To"] = login_data[0]
    message['Subject'] = "Курсы валют на " + str(time.asctime())

    body = "В таблице " + str(row_amount) + " " + get_num_to_str(row_amount)

    message.attach(MIMEText(body, "plain"))
    message.attach(part)
    text = message.as_string()
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(login_data[0], login_data[1])
    server.sendmail(login_data[0], login_data[0], text.encode('utf-8'))
    server.quit()


def get_num_to_str(num: int) -> str:
    lastNum = num % 10
    if lastNum == 1:
        return "строка"
    elif lastNum in [2, 4, 5]:
        return "строки"
    else:
        return "строк"


def take_data_login_gmail() -> list:
    setting = configparser.ConfigParser()
    setting.read('mail.cfg')
    return [setting.get('setting', "email"), setting.get('setting', "password")]
