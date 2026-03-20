from __future__ import annotations
from logging.handlers import RotatingFileHandler
from pathlib import Path 
from modules import *
import pystray
from PIL import Image
import base64
from io import BytesIO
from mail import IMAP, Letter
from spreadsheets import Spreadsheet
from os import listdir
from os import rename
from os.path import isfile
from os.path import isdir

# Максимальное количество дней/часов/минут/секунд в поле
MAX_NUMBER: Literal[10000] = 10000

# Минимальное время между синхронизациями (в секундах)
MIN_SYNC_TIME: Literal[30] = 30

ACCURATE_DEFAULT: Literal[False] = False
SYNC_NUMBER_DEFAULT: Literal[20] = 20
SYNC_DATE_DEFAULT: Literal["2025-01-01"] = "2025-01-01"

COLORS: dict[int, tuple[int]] = {
    # >2 -> red
    3: (1, 0, 0),
    # >1 -> yellow
    2: (1, 1, 0)
}

import os
import certifi
os.environ['SSL_CERT_FILE'] = certifi.where()  # Используем сертификаты из certifi

def setup_logging() -> None:
    """Настраивает логирование в консоль и файл.

    :rtype: None
    :return: None

    Пример:
        >>> setup_logging()
    """
    logs_dir: Path = Path("logs")
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_file: Path = logs_dir / "parser.log"

    formatter: logging.Formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    )

    file_handler: RotatingFileHandler = RotatingFileHandler(
        log_file,
        maxBytes=2 * 1024 * 1024,
        backupCount=2,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)

    console_handler: logging.StreamHandler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logging.basicConfig(level=logging.INFO, handlers=[file_handler, console_handler])

def create_tray_image() -> Image:
    # Вставьте ваш base64 код вместо этой строки
    icon_base64 = """
AAABAAEAYGAAAAEAIAColAAAFgAAACgAAABgAAAAwAAAAAEAIAAAAAAAAJAAAHQSAAB0EgAAAAAAAAAAAAD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AOGabGvtroXx9buV//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/5aiA/+Ghd/Pbk2Vy////AP///wD///8A////AP///wD///8A////AP///wD///8A4Zpsa/e9mP/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//mqYP/25Nlcv///wD///8A////AP///wD///8A////AP///wD///8A7a6F8fe+mf/3vpn/976X//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//oqYH/56qC/+eqg//nqoP/4aF38////wD///8A////AP///wD///8A////AP///wD///8A9buV//e+mf/3vpn/976Z//a5kf/2vJb/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+emfP/opn3/56qD/+eqg//nqoP/5aiA/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3toz/9riQ//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/6aBy/+ikef/nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/9rWK//azif/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+epgv/qmWb/6KR5/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a1jP/2rX//976Y//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56h//+ySWv/npnz/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3to7/9ad1//a8lv/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//po3f/7Y9T/+imff/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/9riQ//Sibv/2uJD/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+qdbP/tjFD/56iA/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e6kv/0n2r/9rOI//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqYL/7JRe/+2MT//oqIH/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3u5T/9J9q//Wrff/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eogP/tjE//7Y9T/+epgv/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/9ryW//ShbP/0pHH/972X//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/6KR5/++FRf/skVn/56mC/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e9mP/1o2//9J5o//a5kv/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//qnm7/74E8/+yVXv/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpj/9KZz//OaYv/2tIr/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+uWYP/wfzj/65hj/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Y//Sod//zmF7/9a1+//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqID/7Y1S//B/OP/rmmj/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/1q3z/85dd//Smc//3vZf/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+ikev/uhkb/7385/+qdbP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/9a2A//OXXv/zn2r/97qT//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/6Z9w/++BPf/wgDv/6aBw/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//awhP/zmF7/85pj//a1jP/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//ql2L/8H84/++CPf/ponT/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/1sof/85lf//OXX//1roH/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/6KmB/+2OVP/wfjf/74NA/+mjd//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/9rWK//SZYf/zl13/9ad1//e9mP/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//opXv/7odH//B+N//uhUP/6KV6/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e2jv/zmmL/85dd//Sgav/3u5P/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+mgcf/vgj7/8H43/+6GRv/opn3/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/2uJD/85xl//OXXf/zm2T/97aM//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/65hk//B/Of/wfjf/7ohK/+eof//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/97qS//SeZ//zl13/85hf//awg//3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+ipgf/sj1X/8H43//B+N//ti07/56iA/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e7lP/zn2n/85dd//OXXf/1qHj/972Y//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/6KZ8/+6ISP/wfjf/8H43/+yOUv/nqYL/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/2vJb/9KFs//OXXf/zl13/9KFt//e7lf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//poHP/8II+//B+N//wfjf/7JBX/+epgv/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/972Y//Wjb//zl13/85dd//ScZP/2t47/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+qZZv/vgDn/8H43//B+N//rlFz/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mP/0pnP/85dd//OXXf/zmF//9rGF//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a+mP/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqYL/7JFX//B+N//wfjf/8H84/+uXYf/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpj/9Kh3//OXXf/zl13/85dd//Wqef/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z/+ypf//gnXP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+imff/uiUn/8H43//B+N//wfzj/6pln/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//WrfP/zl13/85dd//OXXf/0om7/97uV//e+mf/3vpn/976Z//e+mf/yto//2IhV/8VqMf/FajH/04NR/+OlfP/nqoP/56qD/+eqg//nqoP/6aJ0/++DQP/wfjf/8H43/++AOf/qnGr/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/1rYD/85de//OXXf/zl13/9J1m//a4j//3vpn/9ryX/+GXZ//Ibzf/xWox/8VqMf/FajH/xWox/8duN//Zj2D/5qmC/+eqg//rmmj/8IA6//B+N//wfjf/8IA6/+meb//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/9rCE//OYXv/zl13/85dd//SZYP/qpHb/zXdB/8VqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH/y3ZA/+KPXP/wfjf/8H43//B+N//vgTz/6aFz/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//Wyh//zmV//75hi/9aGUv/FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMf/SekT/64A+//CCPv/po3b/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a8lv/glGL/yG83/8VqMf/FajH/xWox/8VqMf/FajH/zG0y/+R4Nf/keDX/zG0y/8VqMf/FajH/xWox/8VqMf/FajH/yG02/9iKWf/mqYH/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/6aR4/8x1Pv/FajH/xWox/8VqMf/FajH/xWox/8hrMv/cdTT/7343//B+N//wfjf/7343/9x1NP/IazL/xWox/8VqMf/FajH/xWox/8VqMf/KdD3/3Zht/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//Gzi//Vgk7/xWox/8VqMf/FajH/xWox/8VqMf/FajH/1HEz/+x9N//wfjf/8H43//B+N//wfjf/8H43//B+N//sfTf/1HEz/8VqMf/FajH/xWox/8VqMf/FajH/xWox/9B/TP/ko3r/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/1u5X/35Ji/8dsNP/FajH/xWox/8VqMf/FajH/xWox/81uMv/meTX/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+Z5Nf/NbjL/xWox/8VqMf/FajH/xWox/8VqMf/GbDT/14tb/+aogf/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z/+ihdf/MdD3/xWox/8VqMf/FajH/xWox/8VqMf/IazL/3nY0/+9+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//vfjf/3nY0/8hrMv/FajH/xWox/8VqMf/FajH/xWox/8pyPP/dl2v/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/wsYj/1IBL/8VqMf/FajH/xWox/8VqMf/FajH/xWox/9ZyM//sfDb/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+x8Nv/WcjP/xWox/8VqMf/FajH/xWox/8VqMf/FajH/0H1K/+Ohef/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/9bqU/9yOXf/GbDP/xWox/8VqMf/FajH/xWox/8VqMf/ObjL/6Ho2//B+N//wfjf/8H43//B+N//ghUj/upRy/6Wch/+aoZT/kKSe/4Wpq/+HqKn/k6Ob/56fkP+pmoP/vZJt/+eCQf/wfjf/8H43//B+N//wfjf/6Ho2/85uMv/FajH/xWox/8VqMf/FajH/xWox/8ZrM//ViFj/5ad//+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/nn3L/y3M7/8VqMf/FajH/xWox/8VqMf/FajH/yWwy/992Nf/wfjf/8H43/+yAO/+zl3n/fK20/1O94P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/0DF9f9lt87/k6Ob/9SJVf/wfjf/8H43//B+N//fdjX/yWwy/8VqMf/FajH/xWox/8VqMf/FajH/ynI7/9uVaf/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/7q6E/9J9SP/FajH/xWox/8VqMf/FajH/xWox/8ZqMf/YczT/7X02//B+N//ehUr/laOZ/06/5v89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z/F9v94r7j/zYxd//B+N//wfjf/7X02/9hzNP/GajH/xWox/8VqMf/FajH/xWox/8VqMf/Pe0b/4aB2/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//S4kf/ai1r/xmsz/8VqMf/FajH/xWox/8VqMf/FajH/z28z/+l7Nv/wfjf/8H43/6qbgv9Gwu7/Pcb4/z3G+P89xvj/Pcb4/z3G+P9Pv+X/fK20/42lov+doJH/rpl//7yTbv+/kWz/uZRz/66Zf/+Vo5n/ea23/03A6P89xvj/Pcb4/z3G+P89xvj/Pcb4/4Kqrf/pgkD/8H43//B+N//pezb/z28z/8VqMf/FajH/xWox/8VqMf/FajH/xmsz/9SHVv/lpn7/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/5Jxt/8txOf/FajH/xWox/8VqMf/FajH/xWox/8psMv/hdzX/8H43//B+N//ghUj/cbLA/z3G+P89xvj/Pcb4/z3G+P9lt87/maGV/86LXP/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+9/OP/MjV7/h6ip/0XD7/89xvj/Pcb4/z3G+P9Yu9v/zYxd//B+N//wfjf/8H43/+F3Nf/KbDL/xWox/8VqMf/FajH/xWox/8VqMf/KcTn/2pJk/+aqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z/+2rgf/QekX/xWox/8VqMf/FajH/xWox/8VqMf/GajH/2nQ0/+59N//wfjf/8H43/+OERf9butj/Pcb4/z3G+P89xvj/bbPF/8mOYP/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+CER/98rLP/Pcb4/z3G+P89xvj/Q8Px/8KQaf/wfjf/8H43//B+N//ufTf/2nQ0/8ZqMf/FajH/xWox/8VqMf/FajH/xWox/854RP/gnnP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/ytpD/2YlW/8VrMv/FajH/xWox/8VqMf/FajH/xWox/9FwM//pezb/8H43//B+N//wfjf/54JB/2K40P89xvj/Pcb4/0PD8f+mnIb/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/wJJr/1S83/89xvj/Pcb4/0TD8P/EkGb/8H43//B+N//wfjf/8H43/+l7Nv/RcDP/xWox/8VqMf/FajH/xWox/8VqMf/FazL/1IVT/+Olff/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/972Y/+KYav/JcDj/xWox/8VqMf/FajH/xWox/8VqMf/LbTL/43g2//B+N//wfjf/8H43//B+N//wfjf/dq+6/z3G+P89xvj/RMPw/8OQZ//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/96FSv9WvN7/Pcb4/z3G+P9Hwez/3YZL//B+N//wfjf/8H43//B+N//wfjf/43g2/8ttMv/FajH/xWox/8VqMf/FajH/xWox/8dvOP/ZkWL/5qmD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/rqH7/znhC/8VqMf/FajH/xWox/8VqMf/FajH/x2sx/9t0NP/ufTf/8H43//B+N//wfjf/8H43//B+N/+8k27/Pcb4/z3G+P8+xff/vJJv//B+N//wfjf/8H43//B+N//wfjf/7384/9eIUf/Akmv/zI1e/+mBP//wfjf/8H43//B+N//wfjf/8H43/+9/OP/XiFH/vpJt/82MXf/kg0T/8H43//B+N//wfjf/8H43//B+N//ah0//UL7j/z3G+P89xvj/c7C+//B+N//wfjf/8H43//B+N//wfjf/8H43/+59N//bdDT/x2sx/8VqMf/FajH/xWox/8VqMf/FajH/zHZB/+Cccv/nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//e+mf/3vpn/8bSN/9iFU//FajH/xWox/8VqMf/FajH/xWox/8VqMf/TcTP/6nw2//B+N//wfjf/8H43//B+N//wfjf/8H43/+uAPP9Xu9z/Pcb4/z3G+P+PpZ//8H43//B+N//wfjf/8H43//B+N/+zl3n/W7rX/z3G+P89xvj/Pcb4/0fC7f+Lp6X/44NF//B+N//wfjf/x45i/0vA6f89xvj/Pcb4/z3G+P8/xfb/cbHB/7qTcf/wfjf/8H43//B+N//wfjf/5oND/9qHT//ah0//2odP/+5/Of/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/6nw2/9NxM//FajH/xWox/8VqMf/FajH/xWox/8VqMf/Sgk//46R8/+eqg//nqoP/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/976Z//a8lv/hlWb/yG82/8VqMf/FajH/xWox/8VqMf/FajH/zG0y/+R4Nf/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/7eVdf89xvj/Pcb4/0+/5f/qgD3/8H43//B+N//wfjf/7n85/4umo/89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/T7/l/8OQZ//qgD3/RcPv/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P92r7r/34VJ//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//keDX/zG0y/8VqMf/FajH/xWox/8VqMf/FajH/x241/9mOX//mqYH/56qD/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z//e+mf/3vpn/36B5/8p1Pv/FajH/xWox/8VqMf/FajH/xWox/8drMf/bdDT/7343//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/32ss/89xvj/Pcb4/6KejP/wfjf/8H43//B+N//wfjf/nKCS/z3G+P89xvj/Pcb4/z/F9v9jt8//f6ux/1O94P89xvj/Pcb4/0PD8f+Wopj/Pcb4/z3G+P89xvj/Pcb4/z/F9v89xvj/Pcb4/z3G+P89xvj/U73h/9eIUf/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/7343/9t0NP/HazH/xWox/8VqMf/FajH/xWox/8VqMf/IdD7/1ZVs/+eqg//nqoP/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/3vpn/976Z/+yxi/+8d0v/o1os/8FoMP/FajH/xWox/8VqMf/FajH/1HEz/+x9N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/64A8/0fB7P89xvj/QMX1/+ODRf/wfjf/8H43//B+N//ViVT/QsTz/z3G+P89xvj/ScHr/8CRav/wfjf/8H43//B+N//Pi1v/W7rY/z3G+P89xvj/Pcb4/z3G+P89xvj/arTH/+SDRP/dhkz/jaWi/0DF9f89xvj/Pcb4/0+/5f/WiVP/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//sfTf/1HEz/8VqMf/FajH/xWox/8VqMf/BaDD/o1os/7p1Sf/foXr/56qD/+eqg//nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z//e+mf/1u5X/zIlf/6ZdMP+jWiz/o1os/8FoMP/FajH/xWox/8xtMv/leTX/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/v5Fs/z3G+P89xvj/brPD//B+N//wfjf/8H43//B+N/+Kp6T/Pcb4/z3G+P89xvj/vpJt//B+N//wfjf/8H43//B+N//wfjf/5oJC/3Gxwf89xvj/Pcb4/z3G+P89xvj/np+Q//B+N//wfjf/8H43/8iNYf9KwOr/Pcb4/z3G+P9cudb/6YE///B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+V5Nf/MbTL/xWox/8VqMf/BaDD/o1os/6NaLP+lXTD/x4Ra/+Wogf/nqoP/56qD/////wD///8A////AP///wD///8A////AP///wD///8A976Z/9yddf+tZjj/o1os/6NaLP+jWiz/o1os/8FoMP/IazL/3XU0/+9+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/o52L/z3G+P89xvj/l6KX//B+N//wfjf/8H43/+9/OP9Pv+X/Pcb4/z3G+P9qtMj/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+5/Of9qtMj/Pcb4/z3G+P89xvj/fqyz//B+N//wfjf/8H43//B+N//UiVX/Ur7h/z3G+P89xvj/f6uw//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//vfjf/3XU0/8hrMv/BaDD/o1os/6NaLP+jWiz/o1os/6tlN//TlGr/56qD/////wD///8A////AP///wD///8A////AP///wD///8AvXdL7qNaLP+jWiz/o1os/6NaLP+jWiz/o1os/9FvMv/sfDb/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/kKSe/z3G+P89xvj/upRx//B+N//wfjf/8H43/9iHUf89xvj/Pcb4/z3G+P+ZoZT/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//ah0//Q8Px/z3G+P89xvj/WrrZ//B+N//wfjf/8H43//B+N//wfjf/zYxd/0DF9f89xvj/Pcb4/8aPY//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+x8Nv/RbzL/o1os/6NaLP+jWiz/o1os/6NaLP+jWiz/u3VJ7////wD///8A////AP///wD///8A////AP///wD///8Ap1oqG6VbLcmjWiz/o1os/6NaLP+jWiz/o1os/+Z6Nv/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/f6yy/z3G+P89xvj/zYxd//B+N//wfjf/8H43/8OQZ/89xvj/Pcb4/z3G+P+xmHz/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/h6ip/z3G+P89xvj/P8X2/+eCQf/wfjf/8H43//B+N//wfjf/8H43/5Gknf89xvj/Pcb4/22zxf/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//mejb/o1os/6NaLP+jWiz/o1os/6NaLP+lWy3Jp1oqG////wD///8A////AP///wD///8A////AP///wD///8A////AJ1VJwWmXC2Zo1os/6NaLP+jWiz/o1os/+Z6Nv/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/bbPF/z3G+P89xvj/14hS//B+N//wfjf/8H43/7WVd/89xvj/Pcb4/z3G+P+7k3D/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/0opX/z3G+P89xvj/Pcb4/8OQZ//wfjf/8H43//B+N//wfjf/8H43/+qAPf9LwOn/Pcb4/z3G+P/Qi1r/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//mejb/o1os/6NaLP+jWiz/o1os/6ZcLZmdVScF////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8AqV4uYKRaLfWjWiz/o1os/+Z6Nv/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/bLPF/z3G+P89xvj/2odP//B+N//wfjf/8H43/8iOYv89xvj/Pcb4/z3G+P+tmn//8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/1G94v89xvj/Pcb4/5yfkf/wfjf/8H43//B+N//wfjf/8H43//B+N/+Jp6b/Pcb4/z3G+P+ZoZX/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//mejb/o1os/6NaLP+kWi31qV4uYP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AKleLjCjWizeo1os/+Z6Nv/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/dq+6/z3G+P89xvj/0YtZ//B+N//wfjf/8H43/+CFSP89xvj/Pcb4/z3G+P+NpqH/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/3yttP89xvj/Pcb4/3avuv/wfjf/8H43//B+N//wfjf/8H43//B+N//LjWD/Pcb4/z3G+P9lts3/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//mejb/o1os/6NaLN6pXi4w////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wCrXCoPpVost+Z6Nv/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/hamq/z3G+P89xvj/w5Bo//B+N//wfjf/8H43//B+N/9Jwev/Pcb4/z3G+P9iuND/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/4ymov89xvj/Pcb4/1C+5P/wfjf/8H43//B+N//wfjf/8H43//B+N//ufzn/QcT0/z3G+P9Hwu3/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//mejb/pVost6tcKg////8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AOx8N+7wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/o52L/z3G+P89xvj/op2K//B+N//wfjf/8H43//B+N/+IqKj/Pcb4/z3G+P8+xff/2IdR//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/5Ojm/89xvj/Pcb4/z3G+P/ehUr/8H43//B+N//wfjf/8H43//B+N//wfjf/U73g/z3G+P89xvj/6oE+//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//sfDfu////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/w5Bo/z3G+P89xvj/f6ux//B+N//wfjf/8H43//B+N//MjV7/Pcb4/z3G+P89xvj/kaSd//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/36ss/89xvj/Pcb4/z3G+P+3lXX/8H43//B+N//wfjf/8H43//B+N//wfjf/ZrbN/z3G+P89xvj/3YZM//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/6oE+/0LE8/89xvj/TcDo/+5/Of/wfjf/8H43//B+N//wfjf/YbfR/z3G+P89xvj/RMPw/9yHTf/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/2O3z/89xvj/Pcb4/z3G+P+RpJ3/8H43//B+N//wfjf/8H43//B+N//wfjf/Wbva/z3G+P8+xff/6oA9//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/2yzxf89xvj/Pcb4/8aPZf/wfjf/8H43//B+N//wfjf/y41g/z/F9v89xvj/Pcb4/2q0yP/pgT//8H43//B+N//wfjf/8H43//B+N//wfjf/xo9l/z3G+P89xvj/Pcb4/z3G+P9qtMf/8H43//B+N//wfjf/8H43//B+N//vfzj/Q8Ty/z3G+P9Nv+f/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/6uagf89xvj/Pcb4/3uttf/wfjf/8H43//B+N//wfjf/8H43/4+ln/89xvj/Pcb4/z3G+P9ntcv/5oJC//B+N//wfjf/8H43//B+N//ihEb/XrnU/z3G+P89xvj/Pcb4/z3G+P9Gwu7/7n85//B+N//wfjf/8H43//B+N//ehUr/Pcb4/z3G+P9kts//8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+iBQP9Fw+//Pcb4/0DF9f/ah07/8H43//B+N//wfjf/8H43/+qAPf9ptMn/Pcb4/z3G+P89xvj/W7rX/5yfkf/NjF3/1IlV/66Zfv9QvuP/Pcb4/z3G+P9rtMb/Pcb4/z3G+P89xvj/0opY//B+N//wfjf/8H43//B+N/+jnYn/Pcb4/z3G+P+Cqq7/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N/+RpZ7/Pcb4/z3G+P94r7j/8H43//B+N//wfjf/8H43//B+N//rgDz/eK+4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/3avuv/ehUr/Pcb4/z3G+P89xvj/rJmA//B+N//wfjf/8H43/+9/OP9Yu9v/Pcb4/z3G+P/Gj2X/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//kg0T/SMLr/z3G+P89xvj/vJNu//B+N//wfjf/8H43//B+N//wfjf/7n85/6Gejf9NwOj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/i6aj/+1/Ov/wfjf/U73h/z3G+P89xvj/hamq//B+N//wfjf/8H43/7qTcf89xvj/Pcb4/1q62f/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/o52L/z3G+P89xvj/ScHr/9aJU//wfjf/8H43//B+N//wfjf/8H43//B+N//pgkD/pJ2J/3+ssv9it8//arTI/4qnpP/Siln/8H43//B+N//wfjf/y41g/72Sbf+9km3/y4xf//B+N//wfjf/3YZL/0+/5f89xvj/Pcb4/6Odi//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/7384/22zxf89xvj/Pcb4/1a83f/ah0//8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//qgD3/YrfP/z3G+P89xvj/V7vc/+qBPv/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/96FSv9Pv+X/Pcb4/z3G+P9Pv+X/y41g//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+GERv9xssD/Pcb4/z3G+P8/xfb/xo9k//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//UiVX/Tb/n/z3G+P89xvj/QMX1/5eil//rgDz/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/uZRz/03A6P89xvj/Pcb4/z3G+P+nnIb/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/0opX/1W93v89xvj/Pcb4/z3G+P9QvuT/o52L/+WCRP/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/7IA7/7mUc/90sLz/Pcb4/z3G+P89xvj/PsX3/5yfkf/wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+aDQ/92r7r/Pcb4/z3G+P89xvj/Pcb4/0HE9P9wssL/naCR/8CSa//Yh1H/4oRG/+aCQv/Tilb/v5Fs/6qbgv+GqKn/SsDq/z3G+P89xvj/Pcb4/z3G+P9Hwez/t5V1//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/upRy/1u61v89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/PsX3/3uttf/fhUn/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+5/Of+6lHL/eq22/0TD8P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/z3G+P89xvj/Pcb4/1C+5P+MpqL/1olT//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43/+mBP//Gj2T/o52J/4Oprf91sLv/arTI/2G30f9ntcr/cbLA/4ioqP+lnIf/zoxc//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+N+Dwfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjf/8H43//B+N//wfjfg////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////APB+Nyrwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw7Xw3NNFwM6fNbjL/zW4y/81uMv/NbjL/zW4y/81uMv/NbjL/zW4y/81uMv/NbjL/zW4y/81uMv/NbjL/zW4y/81uMv/NbjL/zW4y/81uMv/RcDOn7Xw3NPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcw8H43MPB+NzDwfjcq////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wDFajFaxWox9cVqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH/xWox9cVqMVr///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8AxWoxLcVqMdvFajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajHbxWoxLf///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AMVqMQ/FajGxxWox/8VqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH/xWox/8VqMbHFajEP////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wDFajEBxWoxe8VqMf3FajH/xWox/8VqMf/FajH/xWox/8VqMf/FajH9xWoxe8VqMQH///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AMVqMUbFajHsxWox/8VqMf/FajH/xWox/8VqMezFajFG////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wDFajEfxWoxy8VqMf/FajH/xWoxy8VqMR////8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8AxWoxB8VqMZvFajGbxWoxB////wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD////////////////////////////////////////////////////////////////4AAAAAAAAAAAAAB/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/wAAAAAAAAAAAAAA/4AAAAAAAAAAAAAB/+AAAAAAAAAAAAAH//AAAAAAAAAAAAAP//gAAAAAAAAAAAAf//4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB///4AAAAAAAAAAAB////////gAAf////////////wAA/////////////4AB/////////////8AD//////////////AP//////////////gf//////////////w///////////////////////////////////////////////////////////////////////8=
"""
    
    # Декодируем и создаем изображение
    icon_data = base64.b64decode(icon_base64)
    return Image.open(BytesIO(icon_data))

def parse_date(date: str | datetime.datetime, is_genitive: bool = False, short_form: bool = False, remove_milliseconds: bool = True, max_units: int = 2) -> str:
    """
    Преобразует разницу между текущей датой и переданной датой в удобочитаемый формат.

    :rtype: str
    :return: Строка, представляющая разницу во времени (например, "2 дня 3 часа" или "5 мин 30 сек")

    ### Примеры вызова:
    ```
    parse_date("2025-08-01")  # "5 дней 12 часов" (если сейчас 2025-08-06)
    parse_date("2025-08-01", short_form=True)  # "5 дн 12 ч"
    parse_date("2025-08-01", is_genitive=True)  # "5 дня 12 часов"
    parse_date("2025-08-01", max_units=1)  # "5 дней"
    parse_date(datetime.datetime.now())  # "0 секунд"
    ```

    ### Аргументы:
    :param date: Дата для сравнения (строка в формате ISO или объект datetime)
    :param is_genitive: Использовать ли родительный падеж (по умолчанию False)
    :param short_form: Использовать сокращённые обозначения (по умолчанию False)
    :param remove_milliseconds: Исключать ли миллисекунды из результата (по умолчанию True)
    :param max_units: Максимальное количество единиц времени в результате (по умолчанию 2)

    ### Вызываемые исключения:
    :raises ValueError: Если дата имеет неверный формат
    """
    import datetime

    parsed_date: datetime.datetime
    if isinstance(date, str):
        try:
            # Пробуем стандартный ISO формат
            parsed_date = datetime.datetime.fromisoformat(date)
        except ValueError:
            try:
                # Пробуем распарсить RFC 2822 формат (например "Wed, 30 Jul 2025 13:42:26")
                parsed_date = parsedate_to_datetime(date)
            except (TypeError, ValueError):
                # Если оба варианта не сработали, пробуем другие варианты
                try:
                    parsed_date = datetime.datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    return ""
    else:
        parsed_date = date

    diff_ms: float = abs((datetime.datetime.now() - parsed_date).total_seconds() * 1000)
    
    days: int = int(diff_ms // (1000 * 60 * 60 * 24))
    hours: int = int((diff_ms % (1000 * 60 * 60 * 24)) // (1000 * 60 * 60))
    minutes: int = int((diff_ms % (1000 * 60 * 60)) // (1000 * 60))
    seconds: int = int((diff_ms % (1000 * 60)) // 1000)
    milliseconds: int = int(diff_ms % 1000)

    def pluralize(number: int, words: list[str], genitive: bool = False) -> str:
        if genitive:
            if number % 10 == 1 and number % 100 != 11:
                return words[3]
            elif 2 <= number % 10 <= 4 and not 12 <= number % 100 <= 14:
                return words[1]
            else:
                return words[2]
        else:
            if number % 10 == 1 and number % 100 != 11:
                return words[0]
            elif 2 <= number % 10 <= 4 and not 12 <= number % 100 <= 14:
                return words[1]
            else:
                return words[2]

    days_words: list[str] = ["день", "дня", "дней", "дня"] if not short_form else ["дн", "дн", "дн", "дн"]
    hours_words: list[str] = ["час", "часа", "часов", "часа"] if not short_form else ["ч", "ч", "ч", "ч"]
    minutes_words: list[str] = ["минута", "минуты", "минут", "минуту"] if not short_form else ["мин", "мин", "мин", "мин"]
    seconds_words: list[str] = ["секунда", "секунды", "секунд", "секунду"] if not short_form else ["сек", "сек", "сек", "сек"]
    milliseconds_words: list[str] = ["миллисекунда", "миллисекунды", "миллисекунд", "миллисекунду"] if not short_form else ["мс", "мс", "мс", "мс"]

    result: list[str] = []
    remaining_units: int = max_units

    if days > 0 and remaining_units > 0:
        result.append(f"{days} {pluralize(days, days_words, is_genitive)}")
        remaining_units -= 1
    if hours > 0 and remaining_units > 0:
        result.append(f"{hours} {pluralize(hours, hours_words, is_genitive)}")
        remaining_units -= 1
    if minutes > 0 and remaining_units > 0:
        result.append(f"{minutes} {pluralize(minutes, minutes_words, is_genitive)}")
        remaining_units -= 1
    if (seconds > 0 or (diff_ms == 0 and remove_milliseconds)) and remaining_units > 0:
        result.append(f"{seconds} {pluralize(seconds, seconds_words, is_genitive)}")
        remaining_units -= 1
    if (not remove_milliseconds and milliseconds > 0 or diff_ms == 0) and remaining_units > 0:
        result.append(f"{milliseconds} {pluralize(milliseconds, milliseconds_words, is_genitive)}")

    return " ".join(result)

def get_accurate(string: str, must_accurate: bool) -> str:
    if must_accurate:
        return string
    else:
        result: str = ""
        
        # lower, symbols
        for char in string.lower():
            if char not in """«»"',""":
                result += char
            
        return result
    
def get_columns(num_columns: int):
    """
    Генерирует список названий столбцов Excel.

    :param num_columns: Количество столбцов.
    :return: Список названий столбцов, например ["A", "B", ..., "Z", "AA", "AB", ...].
    """
    column_names = []
    for i in range(num_columns):
        name = ""
        while i >= 0:
            i, remainder = divmod(i, 26)
            name = chr(ord("A") + remainder) + name
            i -= 1
        column_names.append(name)
    return column_names

def get_max_date(dates: list[str]) -> str:
    """
    Находит максимальную дату из списка дат в формате "%d.%m.%Y".

    :param dates: Список дат в формате "%d.%m.%Y".
    :return: Объект datetime с максимальной датой.
    """
    return max(datetime.strptime(date, DATE_FORMAT_3) for date in dates)

def count_month(date: datetime, start_date: datetime = datetime(2025, 1, 1)) -> int:
    """
    Возвращает какой это по счёту месяц из даты в 2025 году.

    :param date: Дата в формате "dd.mm.yyyy".
    :return: Порядковый номер месяца, где январь 2025 это 1, февраль это 2, декабрь 2025 это 12, январь 2026 это 13, декабрь 2024 это 0, ноябрь 2024 это -1.
    """
    return (date.year - start_date.year) * 12 + (date.month - start_date.month) + 1

def get_month(month_name: str) -> int:
    try:
        return datetime.strptime(month_name, "%B").month
    except ValueError:
        try:
            return datetime.strptime(month_name, "%b").month
        except ValueError:
            month_name = str(month_name).strip().capitalize()
            months_list: tuple[str] = (
                "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
            )
            if month_name in months_list:
                return months_list.index(month_name) + 1
            else:
                return 0

def parse_column(input_text: str, result_date: str) -> str:
    letter_date: datetime = datetime.strptime(result_date, DATE_FORMAT_3)
    text: str = input_text.strip()

    if not text:
        return ""

    # 📌 Формат: "<Месяц> <Год>"
    match = re.match(r"^([А-Яа-я]+)\s(\d{4})(.*)$", text)
    if match:
        month_name: str = match.group(1)
        year: str = match.group(2)
        remaining: str = match.group(3).strip()
        month_num: int = get_month(month_name)
        if month_num < 1:
            return ""
        return f"01.{month_num:02d}.{year}" + (f" {remaining}" if remaining else "")

    words: list[str] = text.split()
    word: str = words[0]

    # первое встреченное слово - порядковый номер месяца
    periods: dict[tuple[str], int] = {
        # за полугодие
        ("полугодие",): 6,
        # за год
        ("год",): 12
    }

    # <N> квартал/месяц/месяцев/месяца - порядковый номер месяца * значение из словаря
    n_periods: dict[tuple[str], int] = {
        # за 1 квартал
        ("квартал",): 3,
        # за 9 месяцев 
        ("месяц", "месяца", "месяцев"): 1
    }

    max_periods: dict[tuple[str], int] = {
        ("квартал",): 4,
        ("месяц", "месяца", "месяцев"): 12
    }

    for conditions, month in periods.items():
        if word.lower() in conditions:
            year: int = letter_date.year if letter_date.month >= month else letter_date.year - 1
            remaining: str = " ".join(words[1:])
            return f"01.{month:02d}.{year} {remaining}".strip()
    
    if len(words) < 2:
        return ""

    try:
        month: int = int(word)
    except:
        return ""
    else:
        word = words[1]
    
    for conditions, month_power in n_periods.items():
        if word.lower() in conditions:
            max_month: int = max_periods[conditions]

            if month > max_month or month < 1:
                return ""
            else:
                month *= month_power

            year: int = letter_date.year if letter_date.month >= month else letter_date.year - 1
            remaining: str = " ".join(words[2:])
            return f"01.{month:02d}.{year} {remaining}".strip()
    
    return ""

def get_row(sheet: list[list[str]], value: str, must_accurate: bool) -> int:
    """
    Возвращает индекс в списке значений, где находится указанное значение.
    """
    values: str = [row[0] if row else "" for row in sheet]
    accurate_value: str = get_accurate(value, must_accurate)
    accurate_values: list[str] = [get_accurate(v, must_accurate) for v in values]
    
    if accurate_value in accurate_values:
        return accurate_values.index(accurate_value)
    else:
        return -1

def get_row_counterparty(sheet: list[list[str]], value1: str, value2: str, must_accurate: bool) -> int:
    """
    Возвращает индекс строки, где:
    - в первой ячейке значение = value1
    - во второй ячейке значение = value2
    Если не найдено, возвращает -1.
    """
    accurate_value1: str = get_accurate(value1, must_accurate)
    accurate_value2: str = get_accurate(value2, must_accurate)

    for idx, row in enumerate(sheet):
        if len(row) < 1:
            if idx != 0:
                return idx
        elif get_accurate(row[0], must_accurate) == accurate_value1 and (
            len(row) < 2 or row[1] == "" or get_accurate(row[1], must_accurate) == accurate_value2):
            return idx

    return -1

def get_date(month_index: int, from_year: int = 2025, from_month: int = 1) -> str:
    """
    Возвращает последний день месяца по его порядковому номеру, начиная с указанного года и месяца.
    
    :param month_index: Порядковый номер месяца, начиная с указанного года и месяца.
    :param from_year: Год, с которого начинается отсчёт (по умолчанию 2025).
    :param from_month: Месяц, с которого начинается отсчёт (по умолчанию 1 — январь).
    :return: Дата в формате "dd.mm.yyyy".
    """
    # Вычисляем год и месяц на основе month_index
    total_months = from_month - 1 + month_index
    year = from_year + (total_months - 1) // 12  # Корректируем год
    month = (total_months - 1) % 12 + 1  # Корректируем месяц

    # Определяем первый день следующего месяца
    if month == 12:
        next_month = datetime(year + 1, 1, 1)
    else:
        next_month = datetime(year, month + 1, 1)

    # Вычитаем один день, чтобы получить последний день текущего месяца
    last_day = next_month - timedelta(days=1)

    # Возвращаем дату в формате "dd.mm.yyyy"
    return last_day.strftime("%d.%m.%Y")

def get_last_date(date_str: str) -> str:
    """
    Возвращает последний день месяца для указанной даты.

    :param date_str: Дата в формате "dd.mm.yyyy".
    :return: Последний день месяца в формате "dd.mm.yyyy".
    """
    # Преобразуем строку в объект datetime
    date = datetime.strptime(date_str, "%d.%m.%Y")

    # Определяем первый день следующего месяца
    if date.month == 12:
        next_month = datetime(date.year + 1, 1, 1)
    else:
        next_month = datetime(date.year, date.month + 1, 1)

    # Вычитаем один день, чтобы получить последний день текущего месяца
    last_day = next_month - timedelta(days=1)

    # Возвращаем дату в формате "dd.mm.yyyy"
    return last_day.strftime("%d.%m.%Y")

def get_next_date(task: "Task", from_date: datetime):
    """
    Ищет следующую актуальную дату для задачи.
    """
    now_date: datetime = datetime.now()
    new_date: datetime = from_date
    
    while new_date <= now_date:
        new_date = task.get_date(new_date)
    
    return new_date


class Trigger():
    
    def __init__(self, mail_part: Literal["subject", "address", "body", "date"], condition_type: Literal["includes", "startswith", "endswith"], condition_string: str, sheet_title: str, column_type: Literal["date_in_subject", "date_in_date", "date_in_square", "date_in_exact", "counterparty"], column_value: str, must_accurate: bool):
        if str(mail_part).lower() not in ("subject", "address", "body", "date"):
            raise TypeError(f"invalid {mail_part=}")
        
        if str(column_type).lower() not in ("date_in_subject", "date_in_date", "date_in_square", "date_in_exact", "counterparty"):
            raise TypeError(f"invalid {column_type=}")
        
        if str(condition_type).lower() not in ("includes", "startswith", "endswith"):
            raise TypeError(f"invalid {condition_type=}")
            
        self.mail_part: Literal["subject", "address", "body", "date"] = str(mail_part).lower()
        self.condition_type: Literal["includes", "startswith", "endswith"] = str(condition_type).lower()
        self.condition_string: str = str(condition_string)
        self.sheet_title: str = str(sheet_title)
        self.column_type: Literal["date_in_subject", "date_in_date", "date_in_square", "date_in_exact", "counterparty"] = str(column_type).lower()
        self.column_value: str = str(column_value).lower()
        self.must_accurate: bool = not not must_accurate
        
    def check(self, mail: Letter) -> bool:
        condition: str = get_accurate(self.condition_string, self.must_accurate)
        mail_part: str = get_accurate(mail.get(self.mail_part), self.must_accurate)
        
        match self.condition_type:
            case "includes":
                return condition in mail_part
            case "startswith":
                return mail_part.startswith(condition)
            case "endswith":
                return mail_part.endswith(condition)
            case _:
                raise TypeError(f"invalid {self.condition_type=}")
                
        return False
    
    
class Task():
    
    def __init__(self, number: float, unit: Literal["days", "hours", "minutes", "seconds"], time: str = "", last_sync: datetime | None = None):
        self.number: float = float(number)
        self.unit: str = str(unit).lower()
        self.last_sync: datetime | None = None
        
        self.seconds: int = -1
        self.days: int = -1
        self.time: str = time
        self.update(number, unit, time, last_sync or datetime.now())
    
    def update_date(self):
        if self.last_sync and (self.seconds > 0) and (datetime.now() - self.last_sync).total_seconds() > self.seconds:
            # print("last_sync updated:",(datetime.now() - self.last_sync).total_seconds(),self.seconds)
            self.last_sync = datetime.now()
        
    def update(self, number: float, unit: Literal["days", "hours", "minutes", "seconds"], time: str = "", last_sync: datetime | None = None):
        match unit:
            case "hours":
                self.days = -1
                self.seconds = number * 60 * 60
            case "minutes":
                self.days = -1
                self.seconds = number * 60
            case "seconds":
                self.days = -1
                self.seconds = number
            case "days":
                self.days = number
                self.time = str(time)
        
        if last_sync:
            self.last_sync = last_sync
    
    def get_date(self, from_date: datetime) -> datetime:
        """
        Возвращает `from_date` + период.
        Если days=1, то синхронизация выполняется ежедневно:
        - Для 16.03.2025 16:47 вернётся 17.03.2025 12:00 (на следующий день).
        - Для 16.03.2025 11:47 вернётся 16.03.2025 12:00 (в этот же день).
        - Для 17.03.2025 12:00 вернётся 18.03.2025 12:00 (на следующий день).
        
        :param from_date: Дата, от которой ведётся отсчёт.
        
        :rtype: datetime
        :return: Дата следующей синхронизации (спустя период времени).
        """
        if self.days < 0:
            return from_date + timedelta(seconds=self.seconds)
        else:
            # Разбираем строку времени на часы, минуты и секунды
            time_parts = list(map(int, self.time.split(":")))
            hours: int = time_parts[0]
            minutes: int = time_parts[1]
            seconds: int = time_parts[2] if len(time_parts) > 2 else 0

            # Создаём объект datetime с указанным временем на текущий день
            target_time_today = from_date.replace(hour=hours, minute=minutes, second=seconds, microsecond=0)

            # Если текущее время уже позже указанного времени, добавляем дни
            if from_date >= target_time_today:
                target_time_today += timedelta(days=self.days)

            return target_time_today


class Schedule():
    
    def __init__(self, tasks: dict[Task, datetime | None]):
        self.tasks: dict[Task, datetime] = tasks
        self.miss_tasks: dict[Task, datetime] = {}
        self.last_sync: datetime | None = None
        
        self.update_tasks()
        
    def update_task(self, task: Task, last_date: datetime | None):
        date: datetime = datetime.now()
        
        # FIXME: Пересчитывать даты при изменении часового пояса
        if last_date is None or last_date < date:
            # if last_date is not None:
            #     self.add_miss_task(task, last_date)
            next_date: datetime = get_next_date(task, last_date or date)
            self.tasks[task] = next_date
    
    def update_tasks(self):
        for task, last_date in self.tasks.items():
            self.update_task(task, last_date)
            
    def get_dates(self) -> list[datetime]:
        return sorted(list(self.tasks.values()))
    
    def get_date(self, dates: list[datetime | None], index: int):
        dates_list: list[datetime] = []
        
        for date in dates:
            if date:
                dates_list.append(date)
        
        dates_list.sort()
        
        if dates_list:
            return dates_list[index]
        else:
            return None
        
    def get_first_date(self, dates: list[datetime | None]) -> datetime | None:
        return self.get_date(dates, 0)
    
    def get_last_date(self, dates: list[datetime | None]) -> datetime | None:
        return self.get_date(dates, -1)
            
    def get_miss_date(self) -> datetime | None:
        return self.get_last_date(self.miss_tasks.values())
        
    def add_miss_task(self, task: Task, date: datetime):
        self.miss_tasks[task] = date
    
    def is_now(self, date: datetime):
        return date.replace(microsecond=0) == datetime.now().replace(microsecond=0)
        

class Data():
    
    def __init__(self, update_tray: Callable):
        self.update_tray: Callable = update_tray
        self.filenames: tuple[str] = ("DATA1", "DATA2")
        
        self.data: dict[str, ] = {
            "mails": [],
            "spreadsheets": [],
            "schedule": [],
            "triggers": [],
            "settings": {
                "accurate": False,
            },
            "sync": {
                "last_sync": "",
                "sync_number": 20,
                "sync_date": "2025-01-01"
            }
        }
        
        self.gui: dict [str, ] = {
            "done": {}, # int: dict[str, ]
            "lost": 0,
            "working": True,
            "last_sync": "",
            "sync_miss": "",
            "next_sync": "",
            "sync_status": "",
            "sync_progress": "",
            "sync_finished": False,
            "sync_error": "",
            "nodata": (not isfile(self.filenames[0])) and not isfile(self.filenames[1]),
            "menu": "init",
        }
        
        nodata: bool = self.gui["nodata"]
        print(f"{nodata=}")
        self._key: bytes = bytes(0)
        self._iv: bytes = bytes(0)
        
    def update(self, data: dict[str, ], gui: dict[str, ]):
        for key in data:
            self.data[key] = data[key]
            
        for key in gui:
            self.gui[key] = gui[key]
            
        self.save()
    
    def create_filename(self, filename: str) -> str:
        new_filename: str = filename

        i: int = 0
        while isfile(new_filename) or isdir(new_filename):
            filenames: list[str] = listdir()

            while new_filename in filenames:
                i += 1
                new_filename = f"{filename} ({i})"

                if i > 100:
                    return ""
            
            if i > 100:
                return ""
        
        return new_filename
    
    def backup_data(self) -> bool:
        backup_saved: bool = False

        for filename in self.filenames:
            if filename and isfile(filename):
                backup_saved = True

                new_filename: str = self.create_filename(filename)

                if new_filename:
                    try:
                        rename(filename, new_filename)
                    except Exception as err:
                        print("TODO: Логировать ошибку backup")
                else:
                    print("TODO: Логировать ошибку backup")
        
        return backup_saved
        
    def check_keys(self) -> bool:
        # TODO: len
        return len(self._key) > 0 and len(self._iv) > 0
    
    def load_keys(self):
        key, iv = load_keys()

        if (not key) or not iv:
            print(f"created key {len(key)} and iv {len(iv)}")
            key, iv = create_keys()

            if self.backup_data():
                self.show_error()

        print(f"accept key {len(key)} and iv {len(iv)}")
        self._key, self._iv = key, iv

    def show_error(self):
        window = tk.Tk()
        window.title("Parser")
        window.lift()
        window.attributes("-topmost", True)
        window.after(1, lambda: window.focus_force())

        if showerror("Parser", "Не удалось прочитать файлы с сохранёнными настройками. Настройки будут сброшены") or True:
            window.destroy()
    
    def load(self):
        if not self.check_keys():
            self.load_keys()

        for filename in self.filenames:
            print(f"load {filename=}")
            if filename and isfile(filename):
                try:
                    with open(filename, "rb") as file:
                        data: bytes = file.read()
                except Exception as err:
                    print("TODO: Логировать ошибку чтения")
                
                try:
                    content: str = decrypt(data, self._key, self._iv).decode()
                except ValueError as err:
                    print("TODO: Логировать ошибку дешифрования")
                    self.backup_data()
                    self.show_error()
                    return

                try:
                    d: dict = json.loads(content)
                except Exception as err:
                    # FIXME: with catch_json
                    print("TODO: Логировать ошибку json")
                    self.backup_data()
                    self.show_error()
                    return
                else:
                    print(f"loaded {d=}")
                    return self.update(d, {})
                
    def save(self):       
        if not self.check_keys():
            self.load_keys()
         
        for filename in self.filenames:
            if filename:
                try:
                    string: str = json.dumps(self.data)
                    data: bytes = encrypt(string.encode(), self._key, self._iv)
                # TODO: Убрать
                except Exception as err:
                    pass
                
                with open(filename, "wb") as file:
                    file.write(data)
        
        self.gui["nodata"] = (not isfile(self.filenames[0])) and not isfile(self.filenames[1]),
    
    def exists_mail(self) -> bool:
        mails: list = self.data.get("mails", [])
        
        if mails:
            mail: dict = mails[-1]
            
            mail_server: str = mail.get("mail_server", "imap.mail.ru")
            mail_address: str = mail.get("mail_address", "")
            mail_password: str = mail.get("mail_password", "")
            
            return bool(mail_server == "imap.mail.ru" and mail_address and mail_password)
        
        return False
    
    def get_mail(self) -> IMAP | None:
        mails: list = self.data.get("mails", [])
        
        if mails:
            mail: dict = mails[-1]
            if "get" in dir(mail):
                mail_server: str = mail.get("mail_server", "imap.mail.ru")
                mail_address: str = mail.get("mail_address", "")
                mail_password: str = mail.get("mail_password", "")
                
                if mail_server and mail_address and mail_password:
                    return IMAP(mail_server, mail_address, mail_password)
            
        return None
    
    def exists_spreadsheet(self) -> bool:
        spreadsheets: list = self.data.get("spreadsheets", [])
        
        if spreadsheets:
            spreadsheet: dict = spreadsheets[-1]
            if "get" in dir(spreadsheet):
                spreadsheet_credentials: str = spreadsheet.get("spreadsheet_credentials", "")
                spreadsheet_id: str = spreadsheet.get("spreadsheet_id", "")
                return bool(spreadsheet_credentials and isfile(spreadsheet_credentials) and spreadsheet_id)
        
        return False
                    
    
    def get_spreadsheet(self) -> Spreadsheet | None:
        spreadsheets: list = self.data.get("spreadsheets", [])
        
        if spreadsheets:
            spreadsheet: dict = spreadsheets[-1]
            if "get" in dir(spreadsheet):
                spreadsheet_credentials: str = spreadsheet.get("spreadsheet_credentials", "")
                spreadsheet_id: str = spreadsheet.get("spreadsheet_id", "")

                if spreadsheet_credentials and isfile(spreadsheet_credentials):
                    return Spreadsheet("", spreadsheet_credentials, spreadsheet_id)

        return None
    
    def get_schedule(self) -> Schedule:
        schedule: list = self.data.get("schedule", [])
        
        tasks: dict[Task, datetime | None] = {}
        
        for s in schedule:
            schedule_number: float = s.get("schedule_number", 0)
            schedule_unit: Literal["days", "hours", "minutes", "seconds"] = str(s.get("schedule_unit", "")).lower()
            schedule_time: str = str(s.get("schedule_time", ""))
            
            try:
                schedule_number = float(schedule_number)
            except TypeError:
                schedule_number = 0
            except ValueError:
                schedule_number = 0
            except OverflowError:
                schedule_number = 0
            # TODO: Убрать
            except Exception as err:
                self.logger.error("Непредвиденная ошибка при парсинге числа",type(err),err)
                schedule_number = 0
            
            if schedule_number <= 0:
                continue
            
            if not schedule_unit in ("days", "hours", "minutes", "seconds"):
                continue
                
            if schedule_unit == "days":
                if not schedule_time:
                    continue
                
                if len(schedule_time) < 5 or len(schedule_time) > 8:
                    continue
                
                if not str(schedule_time).count(":") in (1, 2):
                    continue
                
            last_sync: str = str(s.get("last_sync", ""))
            
            # FIXME: Надо хранить часовой пояс на момент создания Task пользователем
            date: datetime = None
            
            if last_sync:
                try:
                    date = datetime.strptime(last_sync, DATE_FORMAT_BASED)
                except ValueError:
                    pass
                # TODO: Убрать
                except Exception as err:
                    self.logger.error("Непредвиденная ошибка при парсинге даты",type(err),err)
                    pass
            
            task: Task = Task(schedule_number, schedule_unit, schedule_time, date)
            # task.update_date()
            tasks[task] = date
            
        return Schedule(tasks)
        
    def get_accurate(self) -> bool:
        return not not self.data.get("settings", {}).get("accurate", ACCURATE_DEFAULT)
    
    def get_last_messages(self) -> int:
        return self.data.get("sync", {}).get("sync_number", 0)
    
    def get_last_date(self) -> datetime:
        date: str = str(self.data.get("sync", {}).get("sync_date", ""))
        try:
            return datetime.strptime(date, DATE_FORMAT_2).replace(tzinfo=None)
        except ValueError:
            return datetime.now().replace(tzinfo=None)
        # TODO: Убрать
        except Exception as err:
            self.logger.error("Непредвиденная ошибка при парсинге даты",type(err),err)
            return datetime.now().replace(tzinfo=None)
        
    def set_last_messages(self, number: int):
        self.data["sync"]["sync_number"] = int(number)
        
    def set_last_date(self, date: str):
        self.data["sync"]["sync_date"] = str(date)
    
    def set_sync_progress(self, progress: str = ""):
        if progress == "":
            self.gui["sync_finished"] = True
            self.gui["sync_in_progress"] = False
            self.gui["sync_progress"] = ""
        else:
            self.set_sync_status("loader")
            self.gui["sync_finished"] = False
            self.gui["sync_in_progress"] = True
            self.gui["sync_progress"] = progress
            
        self.update_tray()
    
    def set_success(self, index: int, result: dict):
        result["success"] = True
        self.gui["done"][index] = result
        self.update_tray()
        
    def set_failed(self, index: int, error: str, result: dict):
        result["success"] = False
        result["error"] = error
        self.gui["done"][index] = result
        self.update_tray()
        
    def set_menu(self, menu: str):
        self.gui["menu"] = menu
        
    def set_last_sync(self, date: str, miss: bool):
        if miss:
            self.gui["sync_miss"] = date
        else:
            self.gui["last_sync"] = date
            self.gui["sync_miss"] = ""
        
        # assert not date.startswith("Mon, 24 Mar 2025 00:29")
        
        if date != self.data["sync"]["last_sync"]:
            self.data["sync"]["last_sync"] = date
            self.save()
            
        self.update_tray()
        
    def set_next_sync(self, date: str | Literal["in_progress", "manual"]):
        self.gui["next_sync"] = date
        self.update_tray()
    
    def set_sync_status(self, status: str):
        self.gui["sync_status"] = status
        self.update_tray()
        
    def set_sync_error(self, error: str):
        self.gui["sync_error"] = error
        self.update_tray()
    
class App():
    
    def __init__(self, html_filename: str = "", cmdline_args: list[str] = []):
        self.working: bool = True
        self.threads: Thread = []
        self.in_progress: list[Task] = []
        self.sync_now: bool = False
        self.sync_errors: list[str] = []
        self.sync_status: bool = False  # True - синхронизация выполняется прямо сейчас, False - она остановится
        self.updating: bool = False
        
        self.window_port: int = 8080
        self.windows: list[int] = []
        self.used_ports: list[int] = []
        self.window_showing: bool = False
        
        # Устанавливается в App.create_tray_icon, принимает один аргумент title: str
        self._update_tray: Callable = lambda title: None
        
        self.data: Data = Data(self.update_tray)
        self.mail: IMAP | None = None
        self.schedule: Schedule | None = None
        self.spreadsheet: Spreadsheet | None = None
        
        self.html_filename = html_filename
        self.cmdline_args = cmdline_args
        
        self.lang_strings: dict[str, dict[str, dict[str, str]]] = {}
        self.load_lang()

        self.logger: logging.Logger = logging.getLogger(__name__)
        
    def load_lang(self):
        filename: str = "lang.js"
        
        if not isfile(filename):
            warn(f"Нет файла с языковыми строками {filename} в папке {getcwd()}")
        
        content: str = "{}"
        
        try:
            with open(filename, "r", encoding="utf-8") as file:
                content = file.read()
        except:
            pass
        
        data: dict = {}
        
        if content and "{" in content and "}" in content and content.find("{") < content.find("}"):
            string: str = content[content.find("{"):content.rfind("}")+1]
            
            try:
                data = json.loads(string)
            except:
                pass
        
        if not data:
            data = {"textContent":{"app":{"ru":"Parser"},"init":{"ru":"Загрузка..."},"no_js":{"ru":"Проверьте, включен ли JavaScript в вашем браузере"},"sync_header":{"ru":"Синхронизация почты и Google-таблицы"},"sync_last":{"ru":"Загрузка..."},"last_sync_placeholder":{"ru":"Загрузка..."},"last_sync_miss":{"ru":"Пропущена синхронизация ${} назад"},"last_sync_was":{"ru":"Последняя синхронизация была ${} назад"},"last_sync_nosync":{"ru":"Ещё не было синхронизаций"},"sync_next":{"ru":"Загрузка..."},"next_sync_placeholder":{"ru":"Загрузка..."},"next_sync_finished":{"ru":"Синхронизация 100%"},"next_sync_in_progress":{"ru":"Выполняется синхронизация..."},"next_sync_manual":{"ru":"Синхронизация вручную"},"next_sync_will":{"ru":"Следующая синхронизация будет через "},"sync_number":{"ru":"Последние"},"sync_submit":{"ru":"Синхронизировать"},"sync_cancel":{"ru":"Отмена"},"sync_date":{"ru":"Игнорировать сообщения ранее"},"open_spreadsheet":{"ru":"Открыть Google-таблицу"},"open_settings":{"ru":"Перейти к настройкам"},"settings_submit":{"ru":"Настройки"},"settings_save":{"ru":"Сохранить изменения"},"mail":{"ru":"Почта Mail.ru"},"mail_address":{"ru":"Адрес"},"mail_password":{"ru":"Пароль"},"spreadsheet":{"ru":"Google-таблица"},"spreadsheet_open":{"ru":"Открыть"},"spreadsheet_creating":{"ru":"Создать новую таблицу"},"spreadsheet_create":{"ru":"Создать"},"spreadsheet_grant":{"ru":"Получить"},"spreadsheet_credentials":{"ru":"Файл credentials"},"spreadsheet_askopenfilename":{"ru":"Выбрать"},"spreadsheet_granting":{"ru":"Получение доступа к таблице"},"spreadsheet_gmail":{"ru":"Адрес Gmail"},"trigger":{"ru":"Триггеры"},"trigger_header":{"ru":"Триггер"},"trigger_condition":{"ru":"Условие"},"subject":{"ru":"Тема"},"address":{"ru":"Адрес"},"body":{"ru":"Текст"},"date":{"ru":"Дата"},"includes":{"ru":"содержит"},"startswith":{"ru":"начинается"},"endswith":{"ru":"заканчивается"},"trigger_sheet":{"ru":"Название листа"},"trigger_column":{"ru":"Название столбца"},"date_in_subject":{"ru":"найти дату в теме письма"},"date_in_date":{"ru":"взять дату получения письма"},"date_in_square":{"ru":"взять из квадратных скобок"},"date_in_exact":{"ru":"скопировать из квадратных скобок"},"date_in_particular":{"ru":"указать другое"},"counterparty":{"ru":"найти в теме название контрагента"},"counterparty_occupied":{"ru":"Один и тот же лист не может использоваться одновременно, и для триггеров с датами, и для триггеров с контрагентами"},"trigger_required":{"ru":"Создайте хотя бы один триггер"},"sheet_required":{"ru":"Укажите название листа таблицы"},"trigger_create":{"ru":"Создать триггер"},"schedule":{"ru":"Расписание"},"schedule_every":{"ru":"Кажд."},"seconds":{"ru":"сек."},"minutes":{"ru":"мин."},"hours":{"ru":"час."},"days":{"ru":"дн."},"schedule_time":{"ru":"в"},"schedule_create":{"ru":"Создать ещё"},"welcome":{"ru":"Добро пожаловать"},"app_lost":{"ru":"Перезапустите приложение"},"app_closing":{"ru":"Выход..."},"force_close":{"ru":"Выйти"},"connection_lost":{"ru":"Нет подключения к интернету"},"first_start":{"ru":"Сохранить и начать"},"check_failed":{"ru":"Не удалось проверить почту"},"check_wrong":{"ru":"Неправильная почта или пароль: необходим пароль для внешних приложений"},"check_processing":{"ru":"Дождитесь проверки почты"},"address_incorrect":{"ru":"Некорректный адрес электронной почты"},"password_wrong":{"ru":"Укажите пароль для внешних приложений"},"password_incorrect":{"ru":"Слишком короткий пароль"},"credentials_required":{"ru":"Необходимо выбрать файл credentials для доступа к Google-таблице"},"credentials_failed":{"ru":"Не удалось открыть файл"},"not_created":{"ru":"Необходимо создать Google-таблицу"},"grant_failed":{"ru":"Не удалось выдать доступ"},"grant_required":{"ru":"Укажите адрес электронной почты Gmail"},"grant_incorrect":{"ru":"Для предоставления доступа нужна почта Gmail"},"part_wrong":{"ru":"Некорректная часть сообщения в триггере"},"type_wrong":{"ru":"Некорректный тип проверки в триггере"},"column_wrong":{"ru":"Некорректное название столбца в триггере"},"value_required":{"ru":"Укажите значение - название столбца"},"text_required":{"ru":"Укажите текст триггера"},"number_low":{"ru":"Слишком маленькое число"},"number_large":{"ru":"Слишком большое число"},"unit_wrong":{"ru":"Некорректный период"},"time_incorrect":{"ru":"Некорректное время"},"save_failed":{"ru":"Не удалось сохранить настройки. Перезапустите приложение."}},"tipText":{"mail_address_1":{"ru":"Укажите адрес электронной почты Mail.ru, откуда нужно читать письма"},"mail_password_1":{"ru":"Укажите пароль для внешних приложений"},"settings_submit":{"ru":"Всё в порядке"}}}
        
        for key, value in data.items():
            self.lang_strings[key] = value
    
    def get_lang_phrase(self, id: str, key: str = "textContent", lang: str = "ru") -> str:
        return self.lang_strings.get(key, {}).get(id, {}).get(lang, str(id)).replace("${", "{")
    
    def stop(self):
        if self.working:
            self.working = False
            self.save_schedule()
            
        self.working = False
        self.window_showing = False
        self.window_port = -1
        
        # for thread in self.threads:
        #     try:
        #         thread.join()
        #     except RuntimeError as err:
        #         print("RuntimeError:", err)
    
    def start(self, eel: bool = True):
        self.load_data()
        self.start_cycle()
        if eel:
            self.run_eel()
        self.show_tray()
            
    def run_eel(self):
        self.window_showing = True
        self.init_eel()
        self.start_eel()
        
    def check_working(self) -> bool:
        return not not self.working
    
    def load_data(self):
        self.data.load()
        self.load_mail()
        self.load_spreadsheet()
        
        last_sync: str = self.data.data["sync"]["last_sync"]
        if last_sync:
            self.data.set_last_sync(last_sync, miss = False)
        else:
            self.data.data["sync"]["last_sync"] = NOSYNC
        
        mail: bool = self.data.exists_mail()
        spreadsheet: bool = self.data.exists_spreadsheet()
        must_accurate: bool = self.data.get_accurate()
        triggers: list[Trigger] = self.get_triggers(must_accurate)
        
        if (not check_connection()) and mail and spreadsheet and triggers:
            self.data.set_menu("sync")
            self.data.set_sync_status("failed")
            return
        elif self.mail and self.spreadsheet and triggers:
            self.data.set_menu("sync")
            self.data.set_sync_status("checked")
            return
        
        if not triggers:
            print("triggers failed")
        
        if not mail:
            print("mail failed")
            if self.data.data["mails"]:
                if self.data.data["mails"][0]:
                    self.data.data["mails"][0]["mail_password"] = ""
                else:
                    self.data.data["mails"] = []
        
        if not spreadsheet:
            print("spreadsheet failed")
            if self.data.data["spreadsheets"]:
                if self.data.data["spreadsheets"][0]:
                    self.data.data["spreadsheets"][0]["spreadsheet_credentials"] = ""
                else:
                    self.data.data["spreadsheets"] = []
            
        self.data.set_menu("settings")
        self.data.set_sync_status("failed")
        
    def load_mail(self):
        mail: IMAP | None = self.data.get_mail()
        
        if mail:
            try:
                mail.start()
            except gaierror:
                mail = None
            except IMAP4.error:
                mail = None
            except ConnectionRefusedError:
                mail = None
            # TODO: Убрать
            except Exception as err:
                self.logger.error("Непредвиденная ошибка при обращении к почте",type(err),err)
                mail = None
                
        self.mail = mail
    
    def try_spreadsheet(self, target: Callable) -> FileNotFoundError | googleapiclient.errors.HttpError | gaierror | httplib2.error.ServerNotFoundError | Any:
        try:
            return target()
        except FileNotFoundError as err:
            return err
        except googleapiclient.errors.HttpError as err:
            return err
        except gaierror as err:
            return err
        except httplib2.error.ServerNotFoundError as err:
            return err
        except SSLError as err:
            return err
        except ConnectionResetError as err:
            return err
        except ConnectionAbortedError as err:
            return err
        # TODO: Убрать
        except Exception as err:
            self.logger.error("Непредвиденная ошибка при обращении к таблице",type(err),err)
            return err
        
    
    def wait_limit(self):
        self.data.set_sync_progress(WAIT_LIMIT)
        # print("Ожидание обнуления лимита...")
        sleep(10)
    
    def wait_connection(self, timeout = 60, interval = 1):
        self.data.set_sync_progress(WAIT_CONNECTION)
        # print("Ожидание подключения...")
        while timeout > 0 and not check_connection():
            sleep(interval)
            timeout -= interval
    
    def try_mail(self, target: Callable) -> Exception | Any:
        try:
            return target()
        except gaierror as err:
            return err
        except ConnectionRefusedError as err:
            return err
        except ConnectionAbortedError as err:
            return err
        except IMAP4.error as err:
            return err
        # TODO: Убрать
        except Exception as err:
            self.logger.error("Непредвиденная ошибка при обращении к почте",type(err),err)
            return err
        
        
    def load_spreadsheet(self):
        spreadsheet: Spreadsheet | None = self.data.get_spreadsheet()
        
        if spreadsheet:
            if self.try_spreadsheet(lambda: spreadsheet.start()):
                spreadsheet = None
                
        self.spreadsheet = spreadsheet

    def load_schedule(self):
        self.schedule = self.data.get_schedule()
    
    def save_schedule(self):
        if self.updating:
            return
        
        schedule: Schedule | None = self.schedule
        
        if schedule:
            tasks: list[dict] = []
            
            for task, date in schedule.tasks.items():
                tasks.append({
                    "schedule_number": float(task.number),
                    "schedule_unit": str(task.unit),
                    "schedule_time": str(task.time),
                    "last_sync": task.last_sync.strftime(DATE_FORMAT_BASED) if task.last_sync else "",
                })
                
            self.data.data["schedule"] = tasks
            self.data.save()
            print("save_schedule", tasks)
    
    # Статические методы
    

    def get_url(self) -> str:
        if self.spreadsheet and self.spreadsheet.spreadsheetId:
            return "https://docs.google.com/spreadsheets/d/" + self.spreadsheet.spreadsheetId
        else:
            return ""
    
    def get_triggers(self, must_accurate: bool) -> list[Trigger]:
        result: list[Trigger] = []
        
        for trigger in self.data.data["triggers"]:
            mail_part: str = str(trigger.get("trigger_part", "subject"))
            condition_type: str = str(trigger.get("trigger_type", "includes"))
            condition_string: str = str(trigger["trigger_text"])
            sheet_title: str = str(trigger["trigger_sheet"])
            column_type: str = str(trigger.get("trigger_column", "date_in_subject"))
            column_value: str = ""  # TODO: str(trigger["trigger_column_value"] or "")
            
            try:
                result.append(Trigger(mail_part, condition_type, condition_string, sheet_title, column_type, column_value, must_accurate))
            except TypeError as err:
                print(f"{err.__class__.__name__}: {err=}")
        
        return result
        
    
        
    # def get_messages(self, must_accurate: bool, progress: list = None) -> dict[Letter, Trigger | None]:
    #     quantity: int = self.data.get_last_messages()
    #     date: datetime = self.data.get_last_date()
        
    #     messages: list[Letter] = self.mail.get_letters(quantity, date, progress)
    #     triggers: list[Trigger] = self.get_triggers(must_accurate)
    #     result: dict[Letter, Trigger | None] = {}
        
    #     if triggers:
    #         for message in messages:
    #             trigger: Trigger | None = self.check_triggers(triggers, message)
    #             if trigger:
    #                 result[message] = trigger
    #     else:
    #         for message in messages:
    #             result[message] = None
        
    #     return result
    
    def get_update(self, letter: Letter, trigger: Trigger, add_error: Callable) -> dict[Literal["name", "date", "period", "counterparty"], str]:
        counterparty_mode: bool = trigger.column_type == "counterparty"
        period_mode: bool = trigger.column_type in ("date_in_square", "date_in_exact")
        date_from_subject: bool = True
        
        next_counterparty: bool = False
        next_period: bool = False
        next_date: bool = False
        next_name: bool = False
        
        counterparty: list[str] = []
        period: list[str] = []
        date: str = ""
        name: list[str] = []
        
        for word in letter.subject.split(" "):
            w: str = word.lower()
            
            if counterparty_mode or trigger.column_type == "date_in_exact":
                match w:
                    case "для":
                        next_counterparty = True
                        continue
                    case "за":
                        next_period = True
                        next_counterparty = False
                        continue
            
            if w == "от":
                next_date = True
                next_period = False
                next_counterparty = False
                continue
            
            if period_mode:
                if w.startswith("["):
                    next_period = True
                    period.append(w[1:])
                    continue
                elif w.endswith("]"):
                    next_period = False
                    period.append(w[:-1])
                    continue
            
            if counterparty_mode:
                if next_counterparty:
                    counterparty.append(word)
                    continue
            
            if next_period:
                period.append(word)
            elif next_date:
                date = word
                next_date = False
                next_name = True
            elif next_name:
                name.append(word)
        
        result_counterparty: str = " ".join(counterparty) if counterparty_mode else ""
        result_period: str = " ".join(period)
        result_date: str = letter.date.strftime(DATE_FORMAT_3)
        result_name: str = " ".join(name)
        
        if counterparty_mode and not result_counterparty:
            add_error(f"Некорректная дата в теме письма от {letter.date}")
            print(f"Некорректная дата в теме письма от {letter.date}: {repr(date)}\n\t{repr(letter.subject)}")
        
        if date:
            try:
                datetime.strptime(date, DATE_FORMAT_3)
            except ValueError as err:
                if date_from_subject:
                    add_error(f"Некорректная дата в теме письма от {letter.date}")
                    print(f"Некорректная дата в теме письма от {letter.date}: {repr(date)}\n\t{repr(letter.subject)}", type(err), err)
            else:
                result_date = date
        else:
            add_error(f"Триггер 'найти в теме название контрагента', но в теме письма от {letter.date} нет названия контрагента")
            print(f"Триггер 'найти в теме название контрагента', но в теме письма от {letter.date} нет названия контрагент: {repr(letter.subject)}")
        
        match trigger.column_type:
            case "date_in_square" | "counterparty":
                if result_period:
                    result_period = parse_column(result_period, result_date)

                    if not result_period:
                        if counterparty_mode:
                            add_error(f"Некорректный текст периода в теме письма от {letter.date}")
                            print(f"Некорректный текст периода в теме письма от {letter.date}: {repr(date)}\n\t{repr(letter.subject)}")
                        else:
                            add_error(f"Некорректный текст даты в квадратных скобках в теме письма от {letter.date}")
                            print(f"Некорректный текст даты в квадратных скобках в теме письма от {letter.date}: {repr(date)}\n\t{repr(letter.subject)}")
                elif counterparty_mode:
                    add_error(f"Триггер 'найти в теме название контрагента', но в теме письма от {letter.date} нет периода")
                    print(f"Триггер 'найти в теме название контрагента', но в теме письма от {letter.date} нет периода: {repr(letter.subject)}")
                else:
                    add_error(f"Триггер 'взять из квадратных скобок', но в теме письма от {letter.date} нет квадратных скобок или они не закрыты")
                    print(f"Триггер 'взять из квадратных скобок', но в теме письма от {letter.date} нет квадратных скобок или они не закрыты: {repr(letter.subject)}")
            case "date_in_exact":
                if result_period:
                    # Если возвращает неполный ДД.ММ.ГГГГ без года, то год добавляется через parse_year
                    result_period = parse_column(result_period, result_date)

                    if " " in result_period:
                        result_name = result_period[result_period.find(" ")+1:]
                        result_period = result_period[:result_period.find(" ")]

                    try:
                        datetime.strptime(result_period, DATE_FORMAT_3)
                    except ValueError as err:
                        add_error(f"Некорректная дата в теме письма от {letter.date}")
                        print(f"Некорректная дата в теме письма от {letter.date}: {repr(date)}\n\t{repr(letter.subject)}", type(err), err)
                        result_period = ""
                else:
                    add_error(f"Триггер 'скопировать из квадратных скобок', но в теме письма от {letter.date} нет квадратных скобок или они не закрыты")
                    print(f"Триггер 'скопировать из квадратных скобок', но в теме письма от {letter.date} нет квадратных скобок или они не закрыты: {repr(letter.subject)}")

        if " " in result_period:
            add_error(f"Некорректная дата в теме письма от {letter.date}")
            print(f"Некорректная дата в теме письма от {letter.date}: {repr(date)}\n\t{repr(letter.subject)} {result_period=}")
            result_period = ""

        # FIXME: if result_counterparty and not counterparty_mode add_error и другие ошибки
        return {
            "counterparty": result_counterparty if counterparty_mode else "",
            "period": get_last_date(result_period or result_date),
            "date": result_date,
            "name": result_name
        }
    
    def get_schedule(self) -> Schedule:
        return self.data.get_schedule()
    
    def get_sheets(self, triggers: list[Trigger]) -> tuple[str]:
        sheets: set[str] = set()
        
        for trigger in triggers:
            sheets.add(trigger.sheet_title)
            
        return tuple(sheets)
    
    def get_color(self, repeated: int) -> tuple[float, float, float] | None:
        keys: list[int] = [key for key in COLORS.keys() if isinstance(key, int)]
        keys.sort(reverse=True)
        
        for threshold in keys:
            if repeated >= threshold:
                return COLORS[threshold]
        
        return None
        
    
    def stop_sync(self):
        self.sync_status = False
        
    def sync(self, quantity: int = 0, min_date: datetime = datetime(2025, 1, 1), show_error: bool = True):
        self.sync_errors.clear()
        self.sync_status = True
        self.update_tray()
        
        self.logger.info(f"Инициирована синхронизация")

        def check_sync() -> bool:
            """
            Возвращает True, если синхронизация отменена.
            """
            if not self.check_working():
                print("not working")
                self.disable_sync()
                self.sync_finished(APP_STOPPED)
                return True
            elif not self.sync_status:
                print("sync cancelled")
                self.sync_finished()
                return True
                
        if check_sync():
            return
        
        if not check_connection():
            print("connection err")
            self.logger.error(f"Ошибка в начале синхронизации: {get_connection_err()}")
            return self.sync_finished(CONNECTION_LOST)

        self.logger.info("Первоначальная проверка подключения пройдена")
        self.logger.info("Проверка почты")
        
        if not self.mail:
            self.load_mail()
                
        if check_sync():
            return
        
        if not self.mail:
            print("mail err")
            self.disable_sync()
            return self.sync_finished(MAIL_REQUIRED)
        
        self.logger.info("Проверка почты пройдена")
        self.logger.info("Проверка таблицы")
        
        if not self.spreadsheet:
            self.load_spreadsheet()
                
        if check_sync():
            return
            
        if not self.spreadsheet:
            print("spreadsheet:", "no spreadsheet")
            self.disable_sync()
            return self.sync_finished(SPREADSHEET_REQUIRED)
        
        self.logger.info("Проверка таблицы пройдена")
        
        spreadsheet: Spreadsheet = self.spreadsheet
        
        from time import time
        # FIXME: comment
        start_time: int = time()
        
        must_accurate: bool = self.data.get_accurate()
        # FIXME: must_accurate сделать настраиваемым
        triggers: list[Trigger] = self.get_triggers(must_accurate)
        err: Exception = Exception()
        
        if not triggers:
            print("triggers:", "no triggers")
            self.disable_sync()
            self.sync_finished(TRIGGERS_REQUIRED)
            return err
        
        self.logger.info("Триггеры получены")
                
        if check_sync():
            return
        
        print("sync in progress!")
        self.logger.info("Выполняется синхронизация")
        
        def try_mail(target: Callable) -> Exception | Any:
            if not self.check_working():
                print("not working")
                self.disable_sync()
                self.sync_finished(APP_STOPPED)
                return err
        
            if not check_connection():
                print("connection err")
                self.logger.error(f"Ошибка при обращении к почте {target}: {get_connection_err()}")
                self.sync_finished(CONNECTION_LOST)
                return err
            
            if not self.mail:
                print("mail err")
                self.disable_sync()
                self.sync_finished(MAIL_REQUIRED)
                return err
            
            result: Exception | Any = self.try_mail(target)
            
            if isinstance(result, IMAP4.error) and check_connection():
                print("IMAP4 err")
                self.disable_sync()
                self.sync_finished("Проблемы с почтой. Проверьте настройки")
                return err
            elif isinstance(result, Exception):
                print("MAIL:", result)
                # self.wait_connection()
                # return try_mail(target, attempt + 1)
                self.logger.error(f"Ошибка при обращении к почте {target}: {result}")
                self.sync_finished(CONNECTION_LOST)
                return err
            
            return result
        
        def try_spreadsheet(target: Callable, attempt: int = 0) -> Any | Exception:
            if not self.check_working():
                print("not working!")
                self.disable_sync()
                self.sync_finished(APP_STOPPED)
                return err
        
            if not check_connection():
                print("spreadsheet: connection err")
                if attempt > 2:
                    self.logger.error(f"Ошибка при обращении к таблице {target}: {get_connection_err()}")
                    self.sync_finished(CONNECTION_LOST)
                    return err
                else:
                    self.wait_connection()
                    return try_spreadsheet(target, attempt = attempt + 1)
            
            if not spreadsheet:
                print("spreadsheet:", "no spreadsheet")
                self.disable_sync()
                self.sync_finished(SPREADSHEET_REQUIRED)
                return err
            
            result: Exception | Any = self.try_spreadsheet(target)
            
            if isinstance(result, FileNotFoundError):
                print("spreadsheet:", type(result), result)
                self.disable_sync()
                self.sync_finished(SPREADSHEET_WRONG)
                return err
            elif isinstance(result, gaierror):
                print("spreadsheet:", type(result), result)
                if attempt > 2:
                    self.logger.error(f"Ошибка при обращении к таблице {target}: {result}")
                    self.sync_finished(CONNECTION_LOST)
                    return err
                else:
                    self.wait_connection()
                    return try_spreadsheet(target, attempt = attempt + 1)
            elif isinstance(result, ConnectionAbortedError):
                print("spreadsheet:", type(result), result)
                if attempt > 2:
                    self.logger.error(f"Ошибка при обращении к таблице {target}: {result}")
                    self.sync_finished(CONNECTION_LOST)
                    return err
                else:
                    self.wait_connection()
                    return try_spreadsheet(target, attempt = attempt + 1)
            elif isinstance(result, ConnectionResetError):
                print("spreadsheet:", type(result), result)
                if attempt > 2:
                    self.logger.error(f"Ошибка при обращении к таблице {target}: {result}")
                    self.sync_finished(CONNECTION_LOST)
                    return err
                else:
                    self.wait_connection()
                    return try_spreadsheet(target, attempt = attempt + 1)
            elif isinstance(result, googleapiclient.errors.HttpError):
                print("spreadsheet:", type(result), result)
                if "RATE_LIMIT_EXCEEDED" in str(result):
                    if attempt > 15:
                        self.disable_sync()
                        self.sync_finished(RATE_LIMIT)
                        return err
                    else:
                        self.wait_limit()
                        return try_spreadsheet(target, attempt = attempt + 1)
                else:
                    self.logger.error(f"Ошибка при обращении к таблице {target}: {result}")
                    self.sync_finished(CONNECTION_LOST)
                return err
            elif isinstance(result, SSLError):
                print("spreadsheet:", type(result), result)
                if attempt > 15:
                    self.disable_sync()
                    self.sync_finished(RATE_LIMIT)
                    return err
                else:
                    self.wait_limit()
                    return try_spreadsheet(target, attempt = attempt + 1)
            elif isinstance(result, Exception):
                print("spreadsheet:", type(result), result)
                self.sync_finished(SPREADSHEET_WRONG)
                return err
            
            return result
        
        def get_sheets():
            self.data.set_sync_progress(SPREADSHEET_ATTACHING)
            return spreadsheet.get_sheets()
        
        sheets_items = try_spreadsheet(get_sheets) or err
        
        if sheets_items is err:
            self.sync_finished(SPREADSHEET_FAILED)
            return print("err1")
        
        sheets_ids: tuple[int] = tuple(sheet[0] for sheet in sheets_items)
        sheets_titles: tuple[str] = tuple(sheet[1] for sheet in sheets_items)
                
        if check_sync():
            return
        
        def add_error(error: str):
            if show_error:
                self.sync_errors.append(error)
        
        class Sheet():
            
            def __init__(self, sheet_title: str):
                self.id: int = -1
                self.title: str = str(sheet_title)
                self.cells: list[list[str]] = []
                self.comments: list[list[str]] = []
        
                self.cells_changes: dict[str, list[list[str]]] = {}
                self.comments_changes: dict[str, str] = {}
                self.colors_changes: dict[str, str] = {}
                self.max_row_length: int = 0
                
                self.counterparty: bool = False
                
                # if exists
                if sheet_title in sheets_titles:
                    self.id = int(sheets_ids[sheets_titles.index(sheet_title)])
                    self.cells = try_spreadsheet(lambda: spreadsheet.get_sheet(sheet_title)) or [[]]
                    self.comments = try_spreadsheet(lambda: spreadsheet.get_all_comments(sheet_title)) or []
                    
                    while len(self.comments) < len(self.cells):
                        self.comments.append([])
                else:
                    self.id = try_spreadsheet(lambda: spreadsheet.create_sheet(sheet_title)) or -1
        
            def check(self) -> bool:
                return self.cells is not err and self.comments is not err and self.id >= 0
            
            def update_max(self, row: int):
                length: int = len(self.cells[row])
                if length > self.max_row_length:
                    self.max_row_length = length
        
            def create_row(self) -> int:
                """lock"""
                while creating and not check_sync():
                    sleep(0.01)
                    
                if check_sync():
                    return    
                
                creating.append(True)
                
                row: int = 0
                while row < 1:
                    self.cells.append([])
                    row = len(self.cells) - 1
                
                while len(self.comments) < len(self.cells):
                    self.comments.append([])
                
                creating.remove(True)
                return row
            
            def set_months(self):
                end: int = self.max_row_length
                
                if not self.cells:
                    return
                
                if not self.cells[0]:
                    self.cells[0] = [""]
                    
                start: int = len(self.cells[0])
                counterparty: int = int(self.counterparty)
                
                if counterparty:
                    # FIXME: Костыль!
                    if len(self.cells[0]) >= 2 and self.cells[0][1] == "31.01.2025":
                        self.cells[0][1] = ""
                        self.cells_changes["B1"] = [[""]]
                        add_error(f"Лист '{self.title}' используется одновременно для контрагентов и других триггеров")
                
                start_index = start - counterparty
                
                if counterparty:
                    start_index = max(1, start_index)
                    
                self.cells[0][start_index:end+counterparty] = [get_date(i) for i in range(len(self.cells[0]), end)]
                # sheet_changes[f"{columns[start]}1:{columns[end - 1]}1"] = sheet[0][start:end]
                # FIXME: here Протестировать int(counterparty)!
                for i in range(start_index, end-counterparty):
                    self.cells_changes[f"{columns[i+counterparty]}1"] = [[get_date(i)]]
                # print(f"{columns[start]}1:{columns[end - 1]}1", sheet[0][start:end])
            
            def get_coords(self, update: dict[Literal["name", "date", "period", "counterparty"], str], show_error: bool = True):
                """
                Возвращает координаты ячейки по дате и ключу. При необходимости создаёт строку и расширяет таблицу.
                """
                name: str = update["name"]
                key: str = get_accurate(name, must_accurate)
                period: str = update["period"]
                counterparty: str = update["counterparty"]
                counterparty_key: str = get_accurate(counterparty, must_accurate) if counterparty else ""
                
                row: int
                if counterparty:
                    self.counterparty = True
                    row = get_row_counterparty(self.cells, key, counterparty_key, must_accurate)
                else:
                    row = get_row(self.cells, key, must_accurate)
                    
                    if self.counterparty:
                        add_error(f"Лист '{self.title}' используется одновременно для контрагентов и других триггеров")
                    
                column: int = count_month(datetime.strptime(period, DATE_FORMAT_3))
                
                if column < 1:
                    return -1, -1
                elif self.counterparty:
                    # Потому что название контрагента занимает столбец B (где должен быть январь)
                    column += 1
                
                # FIXME: Можно объединить cells_changes
                if row < 1:
                    # Создание нового ряда
                    row = self.create_row()
                    self.cells[row].append(name)
                    
                    if counterparty:
                        self.cells[row].append(counterparty)
                        self.cells_changes[f"B{row+1}"] = [[counterparty]]
                    self.cells_changes[f"A{row+1}"] = [[name]]
                elif counterparty:
                    # Использование существующего ряда
                    if len(self.cells[row]) < 1:
                        self.cells[row].append(name)
                        self.cells_changes[f"A{row+1}"] = [[name]]
                    
                    if len(self.cells[row]) < 2:
                        self.cells[row].append(counterparty)
                        self.cells_changes[f"B{row+1}"] = [[counterparty]]
                    elif self.cells[row][1] == "":
                        self.cells[row][1] = counterparty
                        self.cells_changes[f"B{row+1}"] = [[counterparty]]
                    
                while len(self.cells[row]) <= column:
                    self.cells[row].append("")
                    
                self.update_max(row)
                    
                while len(self.comments[row]) <= column:
                    self.comments[row].append("")
                
                return row, column
        
        sheets: dict[str, Sheet] = {}

        for sheet_title in self.get_sheets(triggers):
            if check_sync():
                return
            else:
                sheets[sheet_title] = Sheet(sheet_title)
                
                if not sheets[sheet_title].check():
                    self.sync_finished(SPREADSHEET_FAILED)
                    return print("err2")
            
        date_now: datetime = datetime.now()
        
        if quantity < 1:
            quantity = self.data.get_last_messages()
            
        min_date = max(min_date, self.data.get_last_date())
        
        if min_date < datetime(2025, 1, 1):
            min_date = datetime(2025, 1, 1)
            
        columns: list[str] = get_columns(COLUMN_COUNT)
        
        messages_processed: set[int] = set()
        sync_status: list[bool] = [True]
        
        processing: list[str] = []
        creating: list[bool] = []
        
        # FIXME: return self.sync_finished(HEADER_WRONG.format(f"{columns[column_index - 1]}1", expected))
        
        def update_progress(done: int):
            percent: float = 0
            if quantity > 0:
                percent = int(done / quantity * 1000) / 10
            self.data.set_sync_progress(f"{SYNC_PROGRESS} {percent}%...")
            # print(f"{SYNC_PROGRESS} {done} из {quantity} {percent}%...", end = "     \r")
            
        def processed(mail_id: bytes):
            messages_processed.add(mail_id)
            self.logger.info(f"Обработано {len(messages_processed)}/{quantity} сообщений")
            update_progress(len(messages_processed))
        
        def mail_start() -> list:
            self.data.set_sync_progress(MAIL_ATTACHING)
            self.mail.start()
            _, data = self.mail.mail.search(None, "ALL")
            mail_ids = data[0].split()
            return mail_ids
        
        def get_triggers(letter: Letter) -> list[Trigger]:
            result: list[Trigger] = []
            
            for trigger in triggers:
                if trigger.check(letter):
                    result.append(trigger)
                
            return result
        
        messages: list = try_mail(mail_start) or []
        
        # if type(messages) is list:
        #     print(f"messages=[{len(messages)}]")
        # else:
        #     print(f"{messages=}")
        
        if check_sync():
            return
        
        if messages is err:
            return print("err3")
        
        quantity = min(quantity, len(messages))
        messages = messages[-quantity:]
        print(f"{quantity=} {min_date=}")
        
        start_comment: str = DATE_REPEATED
        processed_letters: dict[str, set[bytes]] = {}
        processed_dates: dict[str, dict[str, dict[str, list[str]]]] = {}
        
        def add_letter(letter_id: bytes, sheet_title: str):
            if sheet_title in processed_letters:
                processed_letters[sheet_title].add(letter_id)
            else:
                processed_letters[sheet_title] = set([letter_id])
                
        def add_date(sheet_title: str, key: str, period: str, date: str):
            if sheet_title in processed_dates:
                if key in processed_dates[sheet_title]:
                    if period in processed_dates[sheet_title][key]:
                        processed_dates[sheet_title][key][period].append(date)
                    else:
                        processed_dates[sheet_title][key][period] = [date]
                else:
                    processed_dates[sheet_title][key] = {
                        period: [date]
                    }
            else:
                processed_dates[sheet_title] = {
                    key: {
                        period: [date]
                    }
                }
            print(f"add {date=}", processed_dates[sheet_title][key])
                
        def check_processed(letter_id: bytes, sheet_title: str) -> bool:
            return sheet_title in processed_letters and letter_id in processed_letters[sheet_title]

        def check_repeated(cell: str, comment: str, key: str, period: str, date: str, sheet_title: str) -> bool:
            dates: list[str] = processed_dates[sheet_title][key][period]
            required: int = dates.count(date)
            appears: int = comment[comment.find(start_comment):].count(date)
            print(f"Требуется {required} В комментарии уже {appears} =>", appears >= required)
            if appears == 0 and required == 1 and cell.strip() == date.strip():
                return True
            else:
                return appears >= required
        
        def count_repeated(comment: str, date: str) -> int:
            if date.count(".") > 1:
                date = date[date.rfind("."):]
                
            if start_comment in comment:
                comment = comment[comment.find(start_comment)+len(start_comment):].strip()
                return comment.count(date)
                
            return 0
            
            
        class ThreadControl():
            
            def __init__(self, max_threads: int, target: Callable, args: list, working: list):
                self.max_threads: int = int(max_threads)
                self.args: list = args
                self.target: Callable = target
                self.working: bool = working

                self.cycling: bool = False
                self.cycle_id: int = 0
                self.threads: list[Thread] = []
                
            def start(self):
                if not self.cycling:
                    Thread(target = self.cycle).start()
                
            def cycle(self):
                self.cycle_id += 1
                cycle_id: int = int(self.cycle_id)
                self.cycling = True
                # print("start cycle", cycle_id)
                
                while self.working and self.args and cycle_id == self.cycle_id:
                    if len(self.threads) < self.max_threads:
                        arg = self.args.pop(0)
                        thread: Thread = None
                        
                        def start_target():
                            # print(cycle_id, "Запущена задача", len(self.threads) + 1, "из", self.max_threads)
                            self.threads.append(thread)
                            
                            self.target(arg)
                            
                            if thread in self.threads:
                                self.threads.remove(thread)
                                
                            # print(cycle_id, "Завершена задача", len(self.threads) + 1, "из", self.max_threads)
                            self.start()
                            
                        thread: Thread = Thread(target = start_target)
                        thread.start()
                    else:
                        # print(cycle_id, "остановлен")
                        self.cycle_id += 1
                
                self.cycling = False
                
            def append(self, arg):
                self.args.append(arg)
                self.start()
                
        def process_parsed(letter_id: bytes, sheet_title: str, update: dict[Literal["name", "date", "period", "counterparty"], str], key: str):
            """lock"""
            # FIXME: Thread.lock
            
            while key in processing and not check_sync():
                sleep(0.01)
                
            if check_sync():
                return
                
            processing.append(key)
            
            if check_processed(letter_id, sheet_title):
                return print("SKIP PROCESSED LETTER:", letter_id)
            
            date: str = update["date"]
            period: str = update["period"]
            print(f"{date=} {period=}")
            add_date(sheet_title, key, period, date)
            
            row: int  # Индекс строки
            column: int  # Индекс столбца
            sheet: Sheet = sheets[sheet_title]
            row, column = sheet.get_coords(update)
            
            if column < 1:
                print("TOO OLD MES:", date)
                return
            
            cell: str = sheet.cells[row][column] or ""
            comment: str = sheet.comments[row][column] or ""
            
            old_cell: str = str(cell)
            old_comment: str = str(comment)
            print(f"{old_cell=} {old_comment=}")
            
            # Ищем столбец для указанной даты
            address: str = f"{columns[column]}{row + 1}"
                
            exists_date: datetime | None = None
            
            try:
                exists_date = datetime.strptime(str(cell), DATE_FORMAT_3)
            except ValueError:
                pass
            # TODO: Убрать
            except Exception as err:
                pass
            
            new_date: datetime = datetime.strptime(date, DATE_FORMAT_3)
            
            # Если ячейка занята
            if cell:
                # Составляем текст комментария
                if check_repeated(cell, comment, key, period, date, sheet_title):
                    print("Сообщение уже было обработано")
                elif start_comment in comment:
                    comment += f"\n{date}"
                    print("Дополнен комментарий")
                elif comment:
                    comment += f"\n{start_comment}\n{cell}\n{date}"
                    print("Обновлён комментарий")
                else:
                    comment = f"{start_comment}\n{cell}\n{date}"
                    print("Добавлен комментарий")
                
                # Если в ячейке находится дата
                if exists_date:
                    if cell.strip() == date:
                        # Если даты совпадают, то сообщение уже было обработано
                        print("Сообщение уже было обработано")
                    elif exists_date < new_date:
                        # Если дата из ячейки меньше новой, то заменяем её
                        sheet.cells[row][column] = date
                        sheet.cells_changes[address] = [[date]]
                        print("Значение перезаписано:", address, date)
                
                # Цвет ячейки зависит от количества дат (писем) в этом месяце
                repeated: int = count_repeated(comment, date)    
                
                if repeated > 0 or start_comment in comment:
                    color: tuple[float, float, float] | None = self.get_color(repeated)
                    print(f"{repeated=} {color=}")
                    if color:
                        sheet.colors_changes[address] = color
                
                if comment and comment != old_comment:
                    sheet.comments[row][column] = comment
                    sheet.comments_changes[address] = comment
                    old_comment = comment
                    print("Комментарий изменился:", len(sheet.comments[row][column]))
                    print(f"{comment=}")
                else:
                    print("Комментарий не изменился:", len(sheet.comments[row][column]))
            else:
                sheet.cells[row][column] = date
                sheet.cells_changes[address] = [[date]]
                print("Значение добавлено", address, date)
                            
            # Предупреждение о дате из будущего
            comment = sheet.comments[row][column] or ""
            old_comment = str(comment)
            if new_date > date_now:
                warning: str = f"{DATE_WRONG} - {date}\n"
                
                if warning not in old_comment:
                    comment = warning + comment
                    # print(comment)
                else:
                    comment = ""
                
                sheet.colors_changes[address] = (1, 0, 0)
                
                if comment and comment != old_comment:
                    sheet.comments[row][column] = comment
                    sheet.comments_changes[address] = comment
                    print("Добавлено предупреждение", sheet.colors_changes[address])
                    print("Комментарий изменился:", len(sheet.comments[row][column]))
                else:
                    print("Комментарий не изменился:", len(sheet.comments[row][column]))
            
            add_letter(letter_id, sheet_title)
            
        def process_letter(letter: Letter):
            if check_sync():
                return
            
            triggers_checked: list[Trigger] = get_triggers(letter)
            if not triggers_checked:
                print("MESSAGE: trigger False", letter.subject)
                return processed(letter.id)
            
            print(f"\nLETTER {letter.id}:", letter.subject)
            for trigger in triggers_checked:
                update: dict[Literal["name", "date", "period", "counterparty"], str] = self.get_update(letter, trigger, add_error)
                key: str = get_accurate(update["name"], must_accurate)
                
                print("========================================")
                print(update, "=>", trigger.sheet_title)
                process_parsed(letter.id, trigger.sheet_title, update, key)
                processed(letter.id)
                
                if key in processing:
                    processing.remove(key)
        
        def process(arg):
            if check_sync():
                return
            
            msg_data, mail_id = arg
            
            try:
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)
            except Exception as err:
                self.logger.error("Непредвиденная ошибка при парсинге сообщения",type(err),err)
                processed(mail_id)
                return print("MESSAGE:",err)
            
            date: datetime
            message_date = msg["Date"]
            
            try:
                date = strptime(message_date)
            except Exception as err:
                # Сообщение без даты больше не во входящих
                processed(mail_id)
                self.logger.error("Непредвиденная ошибка при парсинге сообщения",type(err),err)
                return
            
            if date <= min_date:
                processed(mail_id)
                print("MESSAGE:", date, "<=", min_date)
                return 
            # print(date,">=",min_date)
            
            message_sender = msg["From"]
            message_subject = msg["Subject"]
            
            try:
                subject: str = decode_header(message_subject)[0][0].decode()
                sender: str = decode_header(message_sender)
                sender = sender[len(sender) - 1][0]
                if not isinstance(sender, str):
                    sender = sender.decode()
                sender = sender[sender.find("<") + 1:sender.find(">")]
            except Exception as err:
                processed(mail_id)
                self.logger.error("Непредвиденная ошибка при парсинге сообщения",type(err),err)
                return
            
            body = ""
            try:
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                            break
                else:
                    body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")
            except Exception as err:
                processed(mail_id)
                self.logger.error("Непредвиденная ошибка при парсинге сообщения",type(err),err)
                return
            else:
                body = body[:100]
            
            letter: Letter = Letter(mail_id, subject, sender, date, body)
            self.logger.info(f"Письмо {mail_id=} «{subject}»")
            return process_letter(letter)
            
        MAX_THREADS: int = 1
        process_order: ThreadControl = ThreadControl(MAX_THREADS, process, [], sync_status)
            
        if quantity > 0:
            update_progress(0)
            
            for mail_id in messages[-quantity:]:
                m_id: str = mail_id
                
                def mail_fetch():
                    _, msg_data = self.mail.mail.fetch(m_id, "(RFC822)")
                    return msg_data
                
                msg_data = try_mail(mail_fetch) or err
                
                if msg_data is err:
                    print("err4")
                    processed(mail_id)
                    sync_status.clear()
                    break
                
                process_order.append((msg_data, mail_id))
        
            while len(messages_processed) < quantity and sync_status and not check_sync():
                sleep(0.1)
            
        self.logger.info("Завершение синхронизации")
        sync_status.clear()
        
        if check_sync():
            return
        
        self.logger.info("Установка столбцов по месяцам")
        for sheet in sheets.values():
            sheet.set_months()
        
        def spreadsheet_save():
            self.data.set_sync_progress(SPREADSHEET_SAVING)
            spreadsheet.update_sheets(sheets.values())
            print("spreadsheet saved")

        self.logger.info("Сохранение таблицы")
        if try_spreadsheet(spreadsheet_save) is err:
            self.logger.error(f"Ошибка при сохранении таблицы")
            return print("errFINAL")
        
        def format_time(ms: int) -> str:
            """
            Форматирует время в человекочитаемый вид.
            :param ms: Время в миллисекундах.
            :rtype: str
            """
            seconds, milliseconds = divmod(ms, 1000)
            minutes, seconds = divmod(seconds, 60)
            
            parts = []
            if minutes > 0:
                parts.append(f"{minutes} минут")
            if seconds > 0 or minutes > 0:
                parts.append(f"{seconds} секунд")
            parts.append(f"{milliseconds} мс")
            
            return " ".join(parts)
        
        elapsed_time = int((time() - start_time) * 1000)
        # FIXME: comment
        print("Синхронизация выполнена за",format_time(elapsed_time))
        self.logger.info(f"Синхронизация выполнена за {format_time(elapsed_time)}")
        return self.sync_finished()

    
    def create_columns(self, sheet_title: str, column_index: int):
        """
        Создаёт недостающие столбцы, чтобы их общее количество на листе было не меньше `column_index`.
        Заполняет первую ячейку новых столбцов датами. Первую ячейку таблицы оставляет пустой.
        
        :param sheet_title: Название листа таблицы.
        :param column_index: Минимальное количество столбцов на листе (недостающие будут созданы).
        """
        cells: list[str] = self.spreadsheet.get_row(1)
        updates: list[str] = []
        
        # Первая ячейка таблицы остаётся нетронутой
        start_column: int = max(1, len(cells))
        
        # Создаём недостающие столбцы: C1 = "31.03.2025"
        for i in range(start_column, column_index):
            updates.append(get_date(i))
            
        columns: list[str] = get_columns(start_column + len(updates) + 1)
        # print(start_column, len(columns), start_column + len(updates) + 1)
        
        address: str = f"{columns[start_column]}1:{columns[start_column + len(updates) - 1]}1"
        self.spreadsheet.set(sheet_title, address, [updates])
                
    def start_cycle(self):
        thread: Thread = Thread(target=self.cycle, name = "Tasks cycle")
        thread.start()
        self.threads.append(thread)
        
    def cycle(self):        
        def update_last_sync(schedule: Schedule):
            if schedule:
                miss_date: datetime | None = schedule.get_miss_date()
                
                if miss_date:
                    miss_sync: str = strftime(miss_date)
                    self.data.set_last_sync(miss_sync, miss = True)
                else:
                    last_date: datetime = schedule.last_sync
                    if last_date:
                        last_sync: str = strftime(last_date)
                        self.data.set_last_sync(last_sync, miss = False)
        
        def update_next_sync(schedule: Schedule):
            if self.in_progress:
                self.data.set_next_sync(IN_PROGRESS)
                self.data.set_sync_status("loader")
                return
            elif schedule:
                next_date: datetime = schedule.get_first_date(schedule.tasks.values())
                if next_date:
                    next_sync: str = strftime(next_date)
                    return self.data.set_next_sync(next_sync)
            self.data.set_next_sync(MANUAL)
                
        def update_tasks(schedule: Schedule):
            if not schedule:
                return
            
            for task, date in schedule.tasks.items():
                if schedule.is_now(date):
                    if self.in_progress:
                        print("Пропущена синхронизация", datetime.now(),"во время выполнения другой синхронизации")
                        schedule.add_miss_task(task, datetime.now())
                    else:
                        self.start_task(task)
                elif date < datetime.now():
                    schedule.update_task(task, date)
                elif task in schedule.tasks and task.seconds > 0:
                    next_sync: datetime | None = schedule.tasks[task]
                    if next_sync:
                        if abs((next_sync - datetime.now()).total_seconds()) > task.seconds:
                            schedule.update_task(task, datetime.now())
                            # print("next_sync updated:", schedule.tasks[task])
                    
        self.schedule: Schedule = self.data.get_schedule()
        while self.check_working():
            update_last_sync(self.schedule)
            update_next_sync(self.schedule)
            update_tasks(self.schedule)
                    
            if self.sync_now and not self.in_progress:
                self.start_sync()
                
            if not self.in_progress:
                self.data.gui["sync_finished"] = False
                
            sleep(1)
            
    def sync_finished(self, error: str = ""):
        # FIXME: Добавить task в miss_sync
        self.sync_status = False
        self.data.set_sync_progress("")
        self.data.set_sync_status("checked" if not error else "failed")
        self.data.set_sync_error(error)
        self.update_tray()
        
        if error:
            pass
        elif self.schedule:
            self.schedule.miss_tasks.clear()
        
        if self.sync_errors:
            string = "\n".join(set(self.sync_errors))
            
            if len(string) > 1000:
                string = f"{string[:1000]}..."
            
            self.data.set_sync_error(f"{error}\n{string}".strip())
        
        self.in_progress.clear()
        self.data.save()
    
    def disable_sync(self):
        self.schedule = None
    
    def enable_sync(self):
        self.schedule = self.get_schedule()
            
    def start_sync(self):
        self.in_progress.append(True)
        
        def start_sync():
            self.sync_now = False
            sync_date: datetime = datetime.now()
            self.sync()
            
            if not self.check_working():
                return
            
            if self.schedule:
                self.schedule.last_sync = sync_date
            
        Thread(target=start_sync, name = "Manual sync").start()
            
    def start_task(self, task: str):
        self.in_progress.append(task)
        self.sync_now = False
        sync_date: datetime = datetime.now()
        
        def start_task():
            self.logger.info(f"Синхронизация по расписанию {task=}")
            self.sync()
            
            if not self.check_working():
                return
            
            if self.schedule and task:
                self.schedule.last_sync = sync_date

                if task in self.schedule.tasks:
                    if self.schedule.tasks[task] <= datetime.now():
                        if sync_date:
                            task.last_sync = sync_date
                        next_date: datetime = get_next_date(task, sync_date)
                        self.schedule.tasks[task] = next_date
                        # print(task.last_sync)
                        self.save_schedule()
                else:
                    self.logger.error("Непредвиденная ошибка при сохранении задачи: Неизвестная задача")
                
        Thread(target=start_task, name = f"Task {task.number} {task.unit} execution").start()
        
        
    # eel
            
    def manual_sync(self, index: int, data: dict):
        sync_number: int = data["sync_number"] if data and "sync_number" in data else SYNC_NUMBER_DEFAULT
        sync_date: str = data["sync_date"] if data and "sync_date" in data else SYNC_DATE_DEFAULT
        print("manual_sync", sync_number, sync_date)
        self.logger.info(f"Синхронизация вручную {sync_number} {sync_date}")
        
        self.data.set_last_messages(sync_number)
        self.data.set_last_date(sync_date)
        self.data.save()
        self.sync_now = True
        self.enable_sync()
        
    def update_settings(self, index: int, data: dict):
        self.updating = True
        print("update_settings", data["settings"])
        # FIXME: Дополнительные проверки
        self.data.update(data["settings"], {})
        self.data.set_success(index, {})
        self.data.set_sync_status("checked")
        self.update_schedule()
        self.load_mail()
        self.load_spreadsheet()
        self.updating = False
        self.data.set_menu("sync")
    
    def update_schedule(self):
        if self.schedule:
            schedule: Schedule = self.get_schedule()
            tasks: list[Task] = list(schedule.tasks.keys())
            
            for i in range(len(tasks)):
                if i < len(self.schedule.tasks):
                    number: float = tasks[i].number
                    unit: str = tasks[i].unit
                    time: str = tasks[i].time
                    
                    existing_task, date = tuple(self.schedule.tasks.items())[i]
                    existing_task.update(number, unit, time, None)
                    self.schedule.update_task(existing_task, None)
                    tasks[i].update_date()
                else:
                    self.schedule.update_task(tasks[i], None)
                    
            for i in range(len(self.schedule.tasks) - 1, len(tasks) - 1, -1):
                task: Task = tuple(self.schedule.tasks.keys())[i]
                del self.schedule.tasks[task]
                
            self.schedule.update_tasks()
        else:
            self.schedule = self.get_schedule()
                
        
    def grant_access(self, index: int, data: dict):
        gmail: str = data["gmail"]
        filename: str  = data["filename"]
        spreadsheet_id: str  = data["spreadsheet_id"]
        
        if not (filename and isfile(filename)):
            self.data.set_failed(index, CREDENTIALS_REQUIRED, {})
            return
            
        if not spreadsheet_id:
            self.data.set_failed(index, SPREADSHEET_WRONG, {})
            return
        
        spreadsheet: Spreadsheet = Spreadsheet("", filename, spreadsheet_id)
        
        try:
            spreadsheet.start()
        except googleapiclient.errors.HttpError:
            return self.data.set_failed(index, SPREADSHEETS_UNAVAILABLE, {})
        # TODO: Убрать
        except Exception as err:
            self.logger.error("Непредвиденная ошибка при обращении к таблице",type(err),err)
            return self.data.set_failed(index, CREDENTIALS_WRONG, {})
        
        try:
            spreadsheet.share_spreadsheet(gmail)
        # TODO: Убрать
        except Exception as err:
            self.logger.error("Непредвиденная ошибка при выдаче доступа к таблице",type(err),err)
            self.data.set_failed(index, GRANT_FAILED, {})
        else:
            self.data.set_success(index, {})
        
    def open_spreadsheet(self, index: int, data: dict):
        url: str = ""
        if "spreadsheet_id" in data:
            spreadsheet_id: str  = data["spreadsheet_id"]
            if spreadsheet_id:
                url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id
        else:
            url = self.get_url()
        
        if url:
            startfile(url)
            self.data.set_success(index, {})
        else:
            self.data.set_failed(index, SPREADSHEET_WRONG, {})

    def ask_openfilename(self, index: int, data: dict):
        window = tk.Tk()
        window.title("Выбор файла credentials")
        window.lift()
        window.attributes("-topmost", True)
        window.after(1, lambda: window.focus_force())
        filename = askopenfilename(initialdir=getcwd())
        window.destroy()
        
        if not filename:
            return self.data.set_failed(index, CREDENTIALS_NOFILE, {})
        
        if not isfile(filename):
            return self.data.set_failed(index, CREDENTIALS_NOTEXIST, {})
        
        spreadsheet: Spreadsheet = Spreadsheet("", filename, "")
        
        try:
            spreadsheet.start()
        except googleapiclient.errors.HttpError:
            return self.data.set_failed(index, SPREADSHEETS_UNAVAILABLE, {})
        # TODO: Убрать
        except Exception as err:
            # FIXME: SPREADSHEET: ('Unexpected credentials type', None, 'Expected', 'service_account')
            self.logger.error("Непредвиденная ошибка при обращении к таблице",type(err),err)
            return self.data.set_failed(index, CREDENTIALS_WRONGFILE, {})
        else:
            return self.data.set_success(index, {
                "filename": filename
            })
            
    def create_spreadsheet(self, index: int, data: dict):
        filename: str = data["filename"]
        sheet_title: str = data["sheet_title"]
        print("create_spreadsheet", sheet_title)
        
        if not (filename and isfile(filename)):
            return self.data.set_failed(index, CREDENTIALS_REQUIRED, {})
        
        if not sheet_title:
            return self.data.set_failed(index, SHEET_REQUIRED, {})
        
        spreadsheet: Spreadsheet = Spreadsheet("", filename, "")
        
        try:
            spreadsheet.start()
        except FileNotFoundError as err:
            # FIXME: Подробный текст ошибки
            return self.data.set_failed(index, CREDENTIALS_WRONGFILE, {})
        except googleapiclient.errors.HttpError as err:
            return self.data.set_failed(index, SPREADSHEETS_UNAVAILABLE, {})
        # TODO: Убрать
        except Exception as err:
            self.logger.error("Непредвиденная ошибка при обращении к таблице",type(err),err)
            return self.data.set_failed(index, SPREADSHEETS_UNAVAILABLE, {})
        
        print("creating spreadsheet")
        try:
            spreadsheet.create_spreadsheet(APP_TITLE, sheet_title)
        # TODO: Убрать
        except Exception as err:
            print("Непредвиденная ошибка при создании таблицы",type(err),err)
            self.logger.error("Непредвиденная ошибка при создании таблицы",type(err),err)
            
        print("created:", spreadsheet.spreadsheetId)
        if spreadsheet.spreadsheetId:
            return self.data.set_success(index, {
                "spreadsheet_id": spreadsheet.spreadsheetId
            })
        else:
            return self.data.set_failed(index, SPREADSHEET_FAILED, {})
        
    def check_mail(self, index: int, data: dict):
        address: str = data["address"]
        password: str = data["password"]
        print("check_mail", address, password)
        
        mail: IMAP = IMAP("imap.mail.ru", address, password)
        error = ""

        try:
            mail.start()
        except gaierror as err:
            error = CONNECTION_LOST
        except IMAP4.error as err:
            error = MAIL_ERROR
        # TODO: Убрать
        except Exception as err:
            print("Непредвиденная ошибка при обращении к почте",type(err),err)
            self.logger.error("Непредвиденная ошибка при обращении к почте",type(err),err)
            error = MAIL_ERROR
        
        if error:
            print("check: failed")
            return self.data.set_failed(index, error, {})
        else:
            print("check: success")
            return self.data.set_success(index, {})
            
    def _process_command(self, data: dict):
        command: str = data["command"]
        index: int = data["index"]
        port: int = data.get("port", -1)
        self.set_started(int(port))
        
        match command:
            case "cancel_sync":
                self.stop_sync()
            case "save_settings":
                return self.update_settings(index, data)
            case "manual_sync":
                return self.manual_sync(index, data)
            case "grant_access":
                return self.grant_access(index, data)
            case "open_spreadsheet":
                return self.open_spreadsheet(index, data)
            case "ask_openfilename":
                return self.ask_openfilename(index, data)
            case "create_spreadsheet":
                return self.create_spreadsheet(index, data)
            case "check_mail":
                return self.check_mail(index, data)
            
    def process_command(self, data: dict):
        def process_command():
            return self._process_command(data)
        
        if data and type(data) is dict and "command" in data and "index" in data:
            thread: Thread = Thread(target=process_command, name = f"Process {data['command'] if data and 'command' in data else '???'}")
            thread.start()
            self.threads.append(thread)
    
    def ondestroy(self, closed: str, still_working: list):
        pass
        print(f"destroy {closed=} {still_working=} {self.window_port=}")
        # if still_working:
        #     print(type(still_working[0]),still_working[0])
        
        self.window_showing = False
        
        for _, ws in eel._websockets:
            try:
                ws.close()
            except Exception as e:
                print(f"Ошибка при закрытии соединения: {e}")
        # self.stop()
        self.update_tray()
            
    def init_eel(self):
        print("init eel...")
        eel.init(".")

        @eel.expose
        def get_from_python():
            self.data.gui["working"] = self.check_working()
            self.data.gui["port"] = int(self.window_port)
            b = self.data.gui["sync_finished"]
            r = json.dumps({
                "data": self.data.data,
                "gui": self.data.gui
            })
            if b:
                self.data.gui["sync_finished"] = False
                
            if self.window_showing:
                return r
            else:
                return json.dumps({
                    "close_window": True
                })

        @eel.expose
        def send_to_python(data: dict):
            b = self.data.gui["sync_finished"]
            print("got", data)
            self.process_command(data)
            if b:
                self.data.gui["sync_finished"] = False
                
    def start_eel(self):
        # return self.eel()
        thread: Thread = Thread(target = self.eel, name = "GUI")
        thread.start()
        self.threads.append(thread)
        
    def set_started(self, port: int):
        # FIXME: if port == self.window_port
        if port != -1:
            if self.windows:
                self.windows[0] = int(port)
            else:
                self.windows.append(int(port))
    
    def check_repeated(self, port: int, interval: float = 0.1, timeout: float = 3) -> bool:
        while self.check_working() and self.window_port == port and timeout > 0:
            if port in self.windows:
                print(f"eel started at {port=}, {self.window_port=} {self.windows=}")
                return True
            else:
                sleep(interval)
                timeout -= interval
                
        return False
        
    
    def eel(self):
        print("start eel", self.html_filename)
        # if platform == "linux" or platform == "linux2":
        #     eel.browsers.set_path("chrome", r"/usr/bin/google-chrome")
        # elif platform == "win32":
        #     eel.browsers.set_path("chrome", r"GoogleChromePortable\GoogleChromePortable.exe")
        started: bool = False
        
        while self.check_working():
            while self.window_port >= 0 and self.window_port < 65536:
                print(f"eel starting at {self.window_port=}")
                try:
                    eel.start(self.html_filename, mode = "chrome", size = (675, 530), port = self.window_port, block = True, close_callback = self.ondestroy, cmdline_args = self.cmdline_args)
                    print(f"eel ends at {self.window_port=}")
                    started = True
                except OSError:
                    if self.window_port in self.windows:
                        if self.check_repeated(self.window_port):
                            self.update_tray()
                            return
                    else:
                        print(f"eel skip {self.window_port=}")
                        self.window_port += 1
                        pass
                # TODO: Убрать
                except Exception as err:
                    self.logger.error("Непредвиденная ошибка при создании окна",type(err),err)
                    return
            
            if not started:
                # FIXME: infinity loop
                self.window_port = 0
        
        if started:
            print("eel stoped, working=", self.check_working())
        else:
            window = tk.Tk()
            window.title("Parser")
            window.lift()
            window.attributes("-topmost", True)
            window.after(1, lambda: window.focus_force())
            
            showerror("Parser", "Ошибка при запуске программы: Не удалось открыть интерфейс")
        
        self.window_showing = False
    
    def get_tray_title(self):
        # FIXME: Объединить с getNextSync из script.js чтобы не делать одно и то же
        next_sync: str = self.data.gui.get("next_sync", "")
        last_sync: str = self.data.gui.get("last_sync", "")
        
        if len(self.data.gui.get("sync_progress", "")) > 0:
            return self.data.gui["sync_progress"]
        elif self.data.gui.get("sync_error", ""):
            error_text: str = self.data.gui["sync_error"]
            return cut(error_text, 128, dots = True)
        elif self.data.gui.get("sync_in_progress"):
            return self.get_lang_phrase("next_sync_in_progress")
        elif self.data.gui.get("sync_finished"):
            return self.get_lang_phrase("next_sync_finished")
        elif next_sync.lower() == MANUAL:
            if last_sync.strip():
                last_sync = parse_date(last_sync, True)
                string: str = self.get_lang_phrase("last_sync_was")
                last_sync = f". {string}".format(last_sync)
            s: str = self.get_lang_phrase("next_sync_manual")
            
            result: str = f"{s}{last_sync}"
            
            if len(result) <= 128:
                return result
            elif last_sync:
                return last_sync
            elif s:
                return s
            else:
                return self.get_lang_phrase("app")
        elif next_sync.lower() != IN_PROGRESS and next_sync.strip() and ":" in next_sync:
            result: list[str] = []
            
            if last_sync:
                last_sync = parse_date(last_sync, True)
                if last_sync.strip():
                    last_sync = self.get_lang_phrase("last_sync_was").format(last_sync)
                    result.append(last_sync)
                
            if next_sync:
                next_sync = parse_date(next_sync, True)
                if next_sync.strip():
                    next_sync = self.get_lang_phrase("next_sync_will") + next_sync
                    result.append(next_sync)
            
            result: str = ". ".join(result) or self.get_lang_phrase("app")
            
            if len(result) <= 128:
                return result
            elif next_sync:
                return next_sync
            elif last_sync:
                return last_sync
            else:
                return self.get_lang_phrase("app")
        else:
            return self.get_lang_phrase("app")
    
    def update_tray(self):
        title: str = cut(self.get_tray_title(), 128)
        self._update_tray(title)
    
    def show_tray(self):
        def create_tray_icon():
            """
            try:
                image = Image.open("mail.ico")
            except:
                try:
                    # Получаем путь к иконке в временной директории exe
                    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
                    icon_path = join(base_path, "icons", "mail.ico")
                    image = Image.open(icon_path)
                except Exception as e:
                    print(f"Tray icon error: {str(e)}")
                    # Можно создать пустое изображение как fallback
            """
            image = create_tray_image()
            
            # Функции управления окном
            def open_window(icon, item):
                if self.check_working() and not self.window_showing:
                    self.window_showing = True
                    self.start_eel()
                    update_tray()
            
            def close_window(icon, item):
                if self.check_working() and self.window_showing:
                    self.window_showing = False
                    update_tray()
            
            def toggle_window(icon, item):
                if self.window_showing:
                    close_window(icon, item)
                else:
                    open_window(icon, item)
            
            def exit_app(icon, item):
                if self.check_working():
                    icon.stop()
                    self.stop()
                    
            def cancel_sync(icon, item):
                if self.check_working():
                    self.stop_sync()
            
            def manual_sync(icon, item):
                if self.check_working():
                    if self.sync_now or self.in_progress or self.sync_status:
                        pass
                    else:
                        self.sync_now = True
                        self.enable_sync()
            
            # Создаем меню с учетом текущего состояния
            def create_menu():
                buttons: list[pystray.MenuItem] = []
                
                if self.window_showing:
                    buttons.append(
                        pystray.MenuItem("Свернуть", close_window)
                    )
                else:
                    buttons.append(
                        pystray.MenuItem("Открыть", open_window, default = True),
                    )
                
                if self.sync_now or self.in_progress or self.sync_status:
                    buttons.append(
                        pystray.MenuItem("Отмена синхронизации", cancel_sync)
                    )
                else:
                    buttons.append(
                        pystray.MenuItem("Выполнить синхронизацию", manual_sync),
                    )
                
                buttons.append(
                    pystray.MenuItem("Выход", exit_app)
                )
                
                return pystray.Menu(*buttons)
            
            # Обновляем иконку в трее
            def update_tray(title: str = "Parser"):
                nonlocal icon
                if icon is not None:
                    icon.menu = create_menu()
                    icon.title = title
                    try:
                        icon.update_menu()
                    except:
                        pass
            
            # Создаем иконку
            icon = pystray.Icon(
                "Parser",
                image,
                f"Parser",
                create_menu()
            )
            
            # Запускаем в отдельном потоке
            icon.run_detached()
            
            # Первоначальное обновление
            update_tray()
            
            self._update_tray = update_tray
        
        # Запускаем иконку в трее в отдельном потоке
        tray_thread = Thread(target=create_tray_icon, daemon=True, name="TrayIcon")
        tray_thread.start()
        self.threads.append(tray_thread)

cmdline_args: list[str] = [
    # "--start-fullscreen",
    "--incognito",
    "--kiosk",  # Включает киоск-режим, скрывая все элементы интерфейса браузера.
    "--no-first-run",  # Пропускает начальные настройки и приветственные диалоги при первом запуске.
    "--disable-infobars",  # Отключает информационные панели (например, запросы на сохранение паролей).
    "--disable-popup-blocking",  # Отключает блокировку всплывающих окон.
    "--disable-save-password-bubble",  # Отключает всплывающее окно с предложением сохранить пароль.
    "--disable-translate",  # Отключает функционал перевода страниц.
    "--disable-notifications",  # Отключает уведомления.
    "--disable-extensions",  # Отключает все расширения.
    "--disable-background-networking",  # Отключает фоновые сетевые запросы.
    "--disable-background-timer-throttling",  # Отключает ограничение таймеров в фоновых вкладках.
    "--disable-breakpad",  # Отключает отчеты о сбоях.
    "--disable-component-update",  # Отключает автоматическое обновление компонентов браузера.
    "--disable-domain-reliability",  # Отключает мониторинг надежности доменов.
    "--disable-default-apps",  # Отключает установку приложений по умолчанию.
    "--disable-sync",  # Отключает синхронизацию данных.
    "--disable-web-security",  # Отключает политику безопасности одного источника (используйте с осторожностью).
    "--disable-prompt-on-repost",  # Отключает запрос подтверждения при повторной отправке формы.
    "--disable-hang-monitor",  # Отключает мониторинг зависаний.
    "--disable-component-extensions-with-background-pages",  # Отключает компонентные расширения с фоновыми страницами.
    "--disable-backgrounding-occluded-windows",  # Отключает фоновую обработку окон, которые не видны пользователю.
    "--disable-renderer-backgrounding",  # Отключает фоновую обработку рендереров.
    "--disable-logging",  # Отключает логирование.
    "--disable-gpu",  # Отключает использование GPU.
    "--disable-software-rasterizer",  # Отключает использование программного растеризатора.
    "--disable-dev-shm-usage",  # Отключает использование общей памяти для устройств с малым объемом оперативной памяти.
    # "--disable-extensions-file-access-check",  # Отключает проверку доступа к файлам для расширений.
    # "--disable-extensions-http-throttling",  # Отключает ограничение HTTP-запросов для расширений.
    # "--disable-extensions-on-chrome-urls",  # Отключает расширения на страницах Chrome.
    # "--disable-extensions-except"  # Отключает все расширения, кроме указанных.
]

if __name__ == "__main__":
    filename = f"parser.exe"
    print(filename, count_processes(filename))
    if filename and isfile(filename) and count_processes(filename) > 2:
        window = tk.Tk()
        window.title("Повторный запуск программы")
        window.lift()
        window.attributes("-topmost", True)
        window.after(1, lambda: window.focus_force())
        
        if askyesno("Повторный запуск программы", "Вы хотите закрыть все окна Parser? После этого вы сможете запустить программу без помех"):
            window.destroy()
            kill_process(filename)
        else:
            window.destroy()
            destroy()
    
    # setup_logging()
    
    # FIXME: -> index.html before build
    app: App = App("index.html", cmdline_args)
    app.start()
    app.logger.info("Приложение запущено")
    
    while app.check_working():
        sleep(1)
        
    app.stop()
    destroy()
    

# Примеры локальной отладки с реальными учётными данными удалены из репозитория.
