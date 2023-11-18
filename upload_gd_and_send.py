import os
import gspread
import datetime
import re
import logging
from oauth2client.service_account import ServiceAccountCredentials
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from urllib.request import urlopen
from urllib.parse import quote
from smsaero import SmsAero
from PyPDF2 import PdfReader

logging.basicConfig(level=logging.INFO,
                    filename='main.log',
                    format='%(levelname)s %(asctime)s : %(message)s (Line: %(lineno)d) [%(filename)s]',
                    datefmt='%d.%m.%Y %I:%M:%S',
                    encoding='UTF-8',
                    filemode='w')

dirr = 'test'  # Указание папки в которой хранятся файлы

BASE_DIR = os.path.join(f'C:/Users/Desktop/{dirr}')  # директория для хранения файлов

items = os.listdir(BASE_DIR)  # список файлов в папке

SMSAERO_API_KEY = 'Api-key'  # API ключ
SMSAERO_EMAIL = 'login'  # логин
SMSAERO_PASS = 'password'  # пароль

if items:
    logging.info(f'Список файлов {items}')
    try:
        scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive "]
        if not os.path.exists('example.json'):
            logging.info('Создание json файла для доступа к реестру')
            import json  # данные для service_account
            data = {
                "type": "service_account",
                "project_id": "<project_id>",
                "private_key_id": "<private_key_id>",
                "private_key": "<private_key>",
                "client_email": "<client_email>",
                "client_id": "<client_id>",
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_x509_cert_url": "<client_x509_cert_url>",
                "universe_domain": "googleapis.com"
            }
            with open('example.json', 'w') as f:
                json.dump(data, f)
        credentials = ServiceAccountCredentials.from_json_keyfile_name(os.path.join('example.json'), scope)
        client = gspread.authorize(credentials)
        sheet = client.open('<google sheets name>').sheet1
        values = sheet.get_all_values()  # таблица google
        now = datetime.datetime.now().strftime("%d.%m.%Y")  # сохраняем сегодняшнюю дату в переменную
    except Exception as ex:
        logging.exception(ex)

    try:
        logging.info('Подключение к google drive')
        gauth = GoogleAuth()  # аутентификация
        if not os.path.exists('key.json'):  # проверка на наличие ключа
            gauth.LocalWebserverAuth()  # авторизация через браузер
            gauth.SaveCredentialsFile('key.json')  # сохранение токена для входа
        gauth.LoadCredentialsFile('key.json')  # использование токена для входа
        gauth.Refresh()
    except Exception as e_go:
        logging.exception(e_go)

def extract_data(name: str) -> list:
    try:
        logging.info('Чтение файла, получение номера клиента')
        pdf = os.path.join(BASE_DIR, name)  # конкретный файл в итерации цикла
        reader = PdfReader(pdf)  # чтение
        page = reader.pages[2]  # выбираем нужную
        text = page.extract_text()  # получаем текст
        phone = int(re.findall(r'\+\d{1,2}\s?\d{3}\s?\d{3}\s?\d{2}\s?\d{2}\b', text)[1])  # находим номер телефона
        logging.info(f'В {name} Найден номер - {phone}')

        # Ищем сигнатуру для указания отправителя
        last_page = reader.pages[-1]
        cont = last_page.extract_text()
        heads = ['...перечень организаций...']
        for head in heads:
            if head in cont:
                if head in [heads[0], heads[2]]: # [..., ...]
                    signature = '<signature1>'
                    break
                elif head == heads[1]:  # ...
                    signature = '<signature2>'
                    break
                elif head == heads[3]:  # ...
                    signature = '<signature3>'
                    break
        logging.info(f'Найдена организация {signature}')
        return [phone, signature]
    except Exception as _ex:
        logging.exception(_ex)

def upload_dir(file: str, name: str) -> str:
    try:
        logging.info(f'Загрузка {name} в Google Drive')
        drive = GoogleDrive(gauth)  # объект класса Google Drive с доступом
        context = {'title': f'{name}'}  # контекст для загрузки файла
        my_file = drive.CreateFile(context)  # создание файла на google drive
        my_file.SetContentFile(file)

        my_file['parents'] = [{'id': '<folder_id>'}]
        my_file.Upload()
        link = my_file['webContentLink']  # ссылка на скачивание
        logging.info(f'Ссылка для скачивания {name} - {link}')
        return link
    except Exception as _exc:
        logging.exception(_exc)


def main():
    for item in items:
        try:
            file = os.path.join(BASE_DIR, item)
            file_name = item
            pdf_data = extract_data(item)
            phone = pdf_data[0]
            signature = pdf_data[1]
            logging.info('Данные для ссылки')
            url = short_link(url=upload_dir(file, file_name))
        except Exception as exc:
            logging.exception(exc)

        try:
            logging.info(f'Отправка URL-{url} по номеру {phone} - {signature}')
            send_sms(phone, f'Ссылка на полис {url}', sender=signature)
        except Exception as ec:
            logging.exception(ec)

        try:
            os.remove(os.path.join(file))  # удаление файла после загрузки
        except Exception as exc_rem:
            logging.exception(exc_rem)

        try:
            logging.info('Фиксация данных в google sheets')
            sheet.update(f'A{len(values)+1}', now)
            sheet.update(f'B{len(values)+1}', file_name)
            sheet.update(f'C{len(values)+1}', url)
        except Exception as _e:
            logging.exception(_e)


def short_link(url: str) -> str:  # сокращение ссылки
    try:
        logging.info(f'Отправка {url} в сервис clck.ru')

        clickr_a = urlopen('https://clck.ru/--?url=' + quote(url))
        clickr = clickr_a.read()
        clickr_clear = clickr.decode("utf-8")

        logging.info(f'Результат сокращенной ссылки {clickr_clear}')
        return clickr_clear
    except Exception as _ec:
        logging.exception(_ec)


def send_sms(number: int, message: str, sender='<Дефолтное имя отправителя в ЛК>') -> dict:
    try:
        try:
            api = SmsAero(SMSAERO_EMAIL, SMSAERO_PASS, signature=sender)
        except:
            api = SmsAero(SMSAERO_EMAIL, SMSAERO_API_KEY, signature=sender)

        logging.info(f'+{number} - Текст СМС {message}')
        res = api.send(number, message)
        assert res.get('success'), res.get('message')
        logging.info(f'Ответ SMSAERO {res.get("data")}')
        return res.get('data')
    except Exception as e_sms:
        logging.exception(e_sms)


if __name__ == '__main__':
    main()
