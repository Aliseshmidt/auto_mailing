import win32com.client as win32
from pathlib import Path
<<<<<<< HEAD
import schedule
import time
from datetime import date
from path import find_file_by_partial_name

def send_email_with_attachment():
=======

def send_email_with_attachment(to_address, subject, body, attachment_path):
>>>>>>> parent of 7cd1086 (Version for standart outlook)
    # Инициализация Outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

<<<<<<< HEAD
    to_address = '#IBSO_GROUP <list.ibso_group@autofinancebank.ru>'
    cc_address_1 = 'BOCHAROV Anatoly <anatoly.bocharov@autofinancebank.ru>'
    cc_address_2 = 'KOSTIN Oleg <oleg.kostin@autofinancebank.ru>'
    subject = 'Актуальные заявки в ots.cft'
    body = f'Коллеги, здравствуйте!\nПрикрепляю выгрузку реестра активных заявок OTS за {date.today().strftime("%d/%m/%Y")}'
    attachment_path = find_file_by_partial_name()

    # Настройка письма
    mail.To = to_address
    mail.CC = cc_address_1 + ';' + cc_address_2  # Добавление адресов в копию
=======
    # Настройка письма
    mail.To = to_address
>>>>>>> parent of 7cd1086 (Version for standart outlook)
    mail.Subject = subject
    mail.Body = body

    # Добавление вложения
    attachment = Path(attachment_path)
    if attachment.exists():
        mail.Attachments.Add(str(attachment))
    else:
        print(f"Файл {attachment_path} не найден.")
        return

    # Отправка письма
<<<<<<< HEAD
    try:
        mail.Send()
        print(f"Письмо отправлено на адрес: {to_address}")
    except Exception as e:
        print(f"Ошибка отправки письма: {e}")


=======
    mail.Send()
    print(f"Письмо отправлено на адрес: {to_address}")

# Укажите параметры письма
to_address = 'recipient@example.com'
subject = 'Тема письма'
body = 'Текст письма'
attachment_path = '/path/to/your/folder/filename.xlsx'

# Отправка письма
send_email_with_attachment(to_address, subject, body, attachment_path)
import win32com.client as win32
from pathlib import Path

def send_email_with_attachment(to_address, subject, body, attachment_path):
    # Инициализация Outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    # Настройка письма
    mail.To = to_address
    mail.Subject = subject
    mail.Body = body

    # Добавление вложения
    attachment = Path(attachment_path)
    if attachment.exists():
        mail.Attachments.Add(str(attachment))
    else:
        print(f"Файл {attachment_path} не найден.")
        return

    # Отправка письма
    mail.Send()
    print(f"Письмо отправлено на адрес: {to_address}")
>>>>>>> parent of 7cd1086 (Version for standart outlook)

# Укажите параметры письма
to_address = 'recipient@example.com'
subject = 'Тема письма'
body = 'Текст письма'
attachment_path = '/path/to/your/folder/filename.xlsx'

# Отправка письма
send_email_with_attachment(to_address, subject, body, attachment_path)
