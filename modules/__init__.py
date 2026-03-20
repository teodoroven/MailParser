import base64
import logging
import re
import json
import platform
import eel
import requests
import sys
import keyring
from keyring.errors import KeyringError
from os import startfile
from os import getcwd
from os import _exit as EXIT
from os.path import isfile
from os.path import join
from os import system
from threading import Thread
from time import sleep
from datetime import datetime, timedelta, timezone

from ssl import SSLError
from oauth2client.service_account import ServiceAccountCredentials

from socket import gaierror
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import askyesno
from tkinter.messagebox import showerror

import psutil
import pytz
import email
import imaplib
import smtplib
from imaplib import IMAP4
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import parsedate_to_datetime

import httplib2 
from googlediscovery import discovery as discovery
import googleapiclient
from oauth2client.service_account import ServiceAccountCredentials	
    
from Crypto.Cipher import AES
from Crypto.Random import get_random_bytes
from Crypto.Util.Padding import pad, unpad

from warnings import warn

from typing import Literal
from typing import Iterable
from typing import Callable
from typing import Any

ROW_COUNT: int = 1000
COLUMN_COUNT: int = 1000

TZ: Literal["Asia/Yekaterinburg"] = "Asia/Yekaterinburg"
DATE_FORMAT_1: Literal["%a, %d %b %Y %H:%M:%S %z"] = "%a, %d %b %Y %H:%M:%S %z"
DATE_FORMAT_2: Literal["%Y-%m-%d"] = "%Y-%m-%d"
DATE_FORMAT_3: Literal["%d.%m.%Y"] = "%d.%m.%Y"
DATE_FORMAT_BASED: Literal["%d.%m.%Y %H:%M:%S"] = "%d.%m.%Y %H:%M:%S"


APP_TITLE: str = "Мониторинг отправок"
APP_STOPPED: str =  "Приложение остановлено"
GETING_MAILS: str =  "Получение писем..."
CONNECTION_LOST: str =  "Нет подключения к интернету"
RATE_LIMIT: str = "Слишком много синхронизаций"
WAIT_LIMIT: str = "Ожидание обнуления лимита"
WAIT_CONNECTION: str = "Ожидание подключения"
MAIL_REQUIRED: str =  "Не указана почта. Проверьте настройки"
MAIL_WRONG: str =  "Проблемы с почтой. Проверьте настройки"
MAIL_ERROR: str = "Неправильная почта или пароль: необходим пароль для внешних приложений"
MAIL_ATTACHING: str =  "Подключение к почте..."
SPREADSHEET_ATTACHING: str =  "Подключение к таблице..."
SPREADSHEET_SAVING: str = "Сохранение изменений..."
SPREADSHEET_REQUIRED: str =  "Сначала нужно создать таблицу. Проверьте настройки"
SPREADSHEET_WRONG: str =  "Проблемы с таблицей. Проверьте настройки"
SYNC_PROGRESS: str =  "Синхронизация"
TRIGGERS_REQUIRED: str =  "Укажите хотя бы один триггер в настройках"
NAME_REQUIRED: str =  "Название организации не указано в письме"
NAME_WRONG: str =  "Название организации {} не соответствует {}"
HEADER_WRONG: str =  "Постороннее значение в ячейке {} - ожидалось {}"
COLUMN_WRONG: str =  "Не найден подходящий для даты столбец"
DATE_REPEATED: str =  "Для этой организации существуют письма с другими датами в том же месяце:"
DATE_WRONG: str = "Внимание! Дата из будущего"
CREDENTIALS_REQUIRED: str = "Сначала выберите файл credentials"
SHEET_REQUIRED: str = "Укажите название листа таблицы"
SPREADSHEET_WRONG: str = "Сначала создайте Google-таблицу"
SPREADSHEETS_UNAVAILABLE: str = "Ошибка при подключении к Google-таблицам"
CREDENTIALS_WRONG: str = "Неверный файл credentials"
GRANT_FAILED: str = "Ошибка при выдаче доступа"
CREDENTIALS_NOFILE: str = "Файл не выбран"
CREDENTIALS_NOTEXIST: str = "Файл не существует"
CREDENTIALS_WRONGFILE: str = "Неверный файл credentials"
SPREADSHEET_FAILED: str = "Ошибка при создании таблицы"

NOSYNC: str = "nosync"
MANUAL: str = "manual"
IN_PROGRESS: str = "in_progress"

KEY_SIZE = 32  # AES-256
IV_SIZE = 16   # AES.block_size
SERVICE: Literal["Parser_app"] = "Parser_app"
KEY_KEY: Literal["Parser_key"] = "Parser_key"
IV_KEY: Literal["Parser_iv"] = "Parser_iv"

def save_token(key: str, token: str) -> None:
    try:
        keyring.set_password(SERVICE, key, token)
    except KeyringError as e:
        raise RuntimeError(f"Keyring недоступен: {e}") from e

def load_token(key: str) -> str:
    try:
        return keyring.get_password(SERVICE, key) or ""
    except KeyringError as e:
        raise RuntimeError(f"Keyring недоступен: {e}") from e


def _b64e(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode("ascii")


def _b64d(s: str) -> bytes:
    return base64.urlsafe_b64decode(s.encode("ascii"))


def save_bytes(key: str, value: bytes) -> None:
    save_token(key, _b64e(value))

def load_bytes(key: str) -> bytes:
    token = load_token(key)
    if not token:
        return b""
    try:
        return _b64d(token)
    except Exception:
        # токен битый/не base64 — считаем, что его нет
        return b""

def load_keys() -> tuple[bytes, bytes]:
    key = load_bytes(KEY_KEY)
    iv = load_bytes(IV_KEY)

    if len(key) not in (16, 24, 32) or len(iv) != IV_SIZE:
        return b"", b""
    return key, iv

def create_keys() -> tuple[bytes, bytes]:
    key = get_random_bytes(KEY_SIZE)
    iv = get_random_bytes(IV_SIZE)
    save_bytes(KEY_KEY, key)
    save_bytes(IV_KEY, iv)
    return key, iv

def encrypt(data: bytes, key: bytes, iv: bytes) -> bytes:
    cipher = AES.new(key, AES.MODE_CBC, iv)
    return cipher.encrypt(pad(data, AES.block_size))

def decrypt(data: bytes, key: bytes, iv: bytes) -> bytes:
    decipher = AES.new(key, AES.MODE_CBC, iv)
    return unpad(decipher.decrypt(data), AES.block_size)


def is_process_running(process_name):
    """Проверяет, запущен ли процесс с указанным именем."""
    try:
        for proc in psutil.process_iter(attrs=['name']):
            if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                return True
    except:
        pass
    return False

def count_processes(process_name):
    count = 0
    try:
        for proc in psutil.process_iter(attrs=['name']):
            if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                count += 1
    except:
        pass
    return count

def kill_process(process_name):
    """Завершает процесс по его имени."""
    return system(f"taskkill /f /im {process_name}")
    try:
        for proc in psutil.process_iter(attrs=['name', 'pid']):
            if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                # print(f"Завершаем процесс: {proc.info['name']} (PID {proc.info['pid']})")
                proc.kill()  # Принудительное завершение процесса
    except:
        pass

def destroy():
    EXIT(0)
    system("exit")
    system("taskkill /f /im parser.exe")

def strptime(date_str: str) -> datetime:
    date_str = re.sub(r"\s*\(.*\)", "", date_str)
    date_format = "%a, %d %b %Y %H:%M:%S %z"
    dt_utc = datetime.strptime(date_str, date_format).astimezone(pytz.utc)
    tz = pytz.timezone(TZ)
    return dt_utc.astimezone(tz).replace(tzinfo=None)
    
def strftime(date: datetime) -> str:
    """
    ```
    DATE_FORMAT_1: Literal['%a, %d %b %Y %H:%M:%S %z']
    ```
    На выходе значение не будет работать в strptime: используйте .stftime(DATE_FORMAT_BASED).
    Значение предназначено для отправки в js.
    """
    return date.strftime(DATE_FORMAT_1)

def check_connection():
    try:
        requests.get("https://www.google.com", timeout=3)
        # FIXME:
        # requests.get("https://www.googleapis.com", timeout=3)
        # requests.get("https://mail.ru", timeout=3)
    except Exception as err:
        print(type(err),err)
        return False
    else:
        return True


__print__ = print


class Global:
    PRINT_ENABLED: bool = False  # FIXME: -> False before build


def print(*args, **kwargs):
    if Global.PRINT_ENABLED:
        __print__(*args, **kwargs)
    
    # log(" ".join([str(elem) for elem in args]))

def log(*args, **kwargs):
    return print(*args, **kwargs)

def process_start_date(self, dates: list[str]):
    """
    # FIXME: Обработать изменение last_date - если изменено пользователем или данные в таблице отличаются
    # FIXME: надо сдвинуть все данные в таблице и добавить недостающие столбцы
    :param dates: Например, "01.01.2025"
    """
    def get_date() -> str:
        """
        Возвращает первую валидную дату
        """
        for date in dates:
            date: str = str(date)
            
            if date.count(".") != 2:
                continue
            
            d: list[str] = date.split(".")
            m: str = date[1]
            y: str = date[2]
            
            try:
                month: int = int(m)
                year: int = int(y)
            except:
                continue
            
            if month < 1 or month > 12:
                continue
            
            if year < 1:
                continue
            
            return date
        return ""
    
    if len(dates) < 2:
        return
    
    date: str = dates[1]
    
    if not date:
        # Всё в порядке, пустые столбцы будут заполнены автоматически
        return
        
    spreadsheet_month: int = int(date.split(".")[1])
    spreadsheet_year: int = int(date.split(".")[2])
    required_date: datetime = self.data.get_last_date()
    required_month: int = required_date.month
    required_year: int = required_date.year
        
def find_column(self, sheet_title: str, date: str) -> int:
    """
    :param date: Например, "01.01.2025"
    
    :rtype: int
    :return: column_index (from 1)
    """
    
    months: list[str] = self.try_spreadsheet(lambda: self.spreadsheet.get_row(1))
    
    if isinstance(months, Exception):
        return -1
    
    expected_date: str = get_last_date(date)
    expected_month: int = int(expected_date.split(".")[1])
    expected_year: int = int(expected_date.split(".")[2])
    
    columns: list[str] = get_columns(len(months) + 1)
    
    print(f"find column {date=}")
    for i in range(1, len(months)):  # В первом столбце названия организаций
        # Если дата в заголовке столбца соответствует, возвращаем номер столбца
        if months[i].strip() == expected_date:
            print(f"Найдено совпадение: {columns[i]}1")
            return i + 1
        
        # Если в ячейке не дата или пусто, смотрим какие даты в ячейках
        if months[i].count(".") != 2:
            column: list[str] = self.try_spreadsheet(lambda: self.spreadsheet.get_column_index(sheet_title, i + 1))
            
            print(f"{columns[i]}:{columns[i]} {column=}")
            if isinstance(column, Exception):
                return -1
            
            # Найдена ли дата с таким же месяцем и годом
            valid_date: bool = False
            
            # Найдена ли дата с другим месяцем или годом
            invalid_date: bool = False
            
            for cell in column:
                if cell.count(".") == 2:
                    m: str = cell.split(".")[1]
                    y: str = cell.split(".")[2]
                    
                    try:
                        month: int = int(m)
                        year: int = int(y)
                    except TypeError:
                        continue
                    except ValueError:
                        continue
                    except OverflowError:
                        continue
                    # TODO: убрать
                    except Exception as err:
                        logging.getLogger(__name__).error("Непредвиденная ошибка при парсинге даты",type(err),err)
                        continue
                    
                    if month == expected_month and year == expected_year:
                        valid_date = True
                    else:
                        invalid_date = True
                        
            if valid_date and not invalid_date:
                print(f"Найден столбец с подходящими датами: {columns[i]}")
                return i + 1
            else:
                print(f"Столбец {columns[i]} не подходит")
    
    # Не хватило ячеек
    print(f"Создаём новый столбец {columns[len(months)]}1 = {expected_date} ->", len(months) + 1)
    if isinstance(self.try_spreadsheet(lambda: self.spreadsheet.set(
            sheet_title, f"{columns[len(months)]}1", [[expected_date]]
        )), Exception):
        return -1
    
    return len(months) + 1

def cut(obj: Iterable, max_length: int = float("inf"), dots: bool = False) -> Iterable:
    """
    :rtype: obj[:max_length]
    :return: Обрезанную копию индексируемого объекта, длина которого не превышает указанное значение max_length
    - Если не указывать max_length, возвращает изначальный объект без изменений
    - Если isinstance(obj,str) и dots=True, то в конец строки будет добавлено многоточие, при этом длина строки с точками не будет превышать max_length

    ### Примеры вызова:
    ```
    cut("1234567890",5) # "12345"
    cut("12",5) # "12"
    cut("1234567890") # "1234567890"
    cut("1234567890",5,dots=True) # "12..." # Точки влезли
    cut("1234567890",3,dots=True) # "123" # Точки не влезли
    cut((1,2,3,4,5),2) # (1,2)
    cut([1,2,3,4,5],2) # [1,2]
    cut([1,2,3,4,5],2,dots=True) # [1,2] # Ошибка вызвана не будет, точек не будет
    cut(1) # raise TypeError()
    ```

    ### Пример использования:
    ```
    if len(text) > MAX_BUTTON_LENGTH:
        text = cut(text, MAX_BUTTON_LENGTH, dots=True)
    ```

    ### Аргументы:
    :param obj: Индексируемый объект, который необходимо обрезать
    :param max_length: Целое число - максимальная длина индексируемой копии
    :param dots: Добавлять ли точки в конец строки при обрезании
    
    ### Вызываемые исключения:
    :raises TypeError: if obj is not iterable
    """
    if len(obj) > max_length:
        if isinstance(obj, str):
            max_length: int = int(max_length) - len("..." if dots and max_length > 3 else "")
            return obj[:max_length] + ("..." if dots else "")
        else:
            return obj[:max_length]
    else:
        return obj
