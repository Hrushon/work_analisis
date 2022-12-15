"""Модуль с основной логикой."""

import sqlite3

from openpyxl import load_workbook, Workbook

from constants import COMPRES_TUPLE, DB_FILE, FILENAME, TABLES_NAME
from manage_db import (add_data_in_tables, annotate_data_in_tables,
                       create_table, summary_data_in_tables)

uktol = list()
compressors = list()
others = list()

CHOICE = {
    'уктол': uktol,
    'мк': compressors,
    'прочее': others
}

TABLE_NAME_AND_DATA = {
    'uktol': uktol,
    'compressors': compressors,
    'others': others
}

wb = load_workbook(filename=FILENAME, data_only=True)

wb.active = 0

sheet = wb.active

for i in range(1, sheet.max_row):
    col = sheet['A' + str(i + 1)]
    if col.value.lower() in COMPRES_TUPLE:
        target = CHOICE['мк']
    elif col.value.lower() == 'уктол':
        target = CHOICE[col.value.lower()]
    else:
        target = others

    result = tuple(
        col[i].value for col in sheet.iter_cols(0, sheet.max_column)
    )
    target.append(result)

con = sqlite3.connect(DB_FILE, uri=True)
cur = con.cursor()

create_table(*TABLES_NAME, cur=cur)

for table, data in TABLE_NAME_AND_DATA.items():
    add_data_in_tables(table, cur=cur, data=data)

ann_uktol, ann_compressors, ann_others = (
    annotate_data_in_tables(item, cur=cur)
    for item in TABLES_NAME
)

sum_result = [
    summary_data_in_tables(item, cur=cur)
    for item in TABLES_NAME
]

con.commit()
con.close()

RESULT_TABLES = {
    'уктол': ann_uktol,
    'мк': ann_compressors,
    'прочее': ann_others
}

result_wb = Workbook()

dest_filename = 'shit.xlsx'

del result_wb['Sheet']

wb_sheet = result_wb.create_sheet(title='общая')
wb_sheet.append(('Оборудование', 'Месяц', 'Год', 'Количество'))
for value in sum_result:
    for item in value:
        wb_sheet.append(item)

for name, value in RESULT_TABLES.items():
    wb_sheet = result_wb.create_sheet(title=name)
    wb_sheet.append(('Узел', 'Месяц', 'Год', 'Количество'))
    for item in value:
        wb_sheet.append(item)

result_wb.save(filename=dest_filename)
