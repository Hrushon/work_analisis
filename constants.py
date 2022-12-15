"""Модуль с константами."""

import os
from typing import Tuple

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

COMPRES_TUPLE = ('вв', 'акв', 'дэн')

EXCEL_NAME = 'analisis.xlsx'
FILENAME = os.path.join(BASE_DIR, EXCEL_NAME)

DB_NAME = 'db.sqlite'
DB_FILE = 'file:' + os.path.join(BASE_DIR, DB_NAME)

HARDWARE: Tuple[str] = ('УКТОЛ', 'ВВ', 'АКВ', 'ДЭН', 'Прочее')

RESPONDENTS: Tuple[str, ...] = (
    'РЖД', 'СТМ-Сервис', 'НЭРЗ', 'Входная', 'Иртышское',
    'Комбинатская', 'Автоматный', 'ТР-1', 'ТР-3'
    )

TABLES_NAME = ['uktol', 'compressors', 'others']
