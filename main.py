"""Модуль с основной логикой."""

import sqlite3

from openpyxl import load_workbook, Workbook

from constants import COMPRES_TUPLE, DB_FILE, FILENAME, TABLES_NAME
from manage_db import (add_data_in_tables, annotate_data_in_tables,
                       annotate_data_previous_year, create_table,
                       current_month, current_year)

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

sum_result = [ann_uktol, ann_compressors, ann_others]

RESULT_TABLES = {
    'уктол': ann_uktol,
    'мк': ann_compressors,
    'прочее': ann_others
}

COL_NAMES = (
    'Оборудование',
    'Узел',
    f'{current_year}г.',
    f'{current_year - 1}г.',
    f'{current_year - 2}г.',
    f'{current_year - 3}г.',
    f'{current_year - 4}г.'
)

result_wb = Workbook()

dest_filename = 'shit.xlsx'

del result_wb['Sheet']

wb_sheet = result_wb.create_sheet(title='общая')
wb_sheet.append(COL_NAMES)

for name, value in RESULT_TABLES.items():
    wb_sheet = result_wb.create_sheet(title=name)
    wb_sheet.append(COL_NAMES)
    for item in value:
        details = item[:2]
        if details[0].lower() in COMPRES_TUPLE:
            details = ('compressors', details[1])
        elif details[0].lower() == 'уктол':
            details = ('uktol', details[1])
        else:
            details = ('others', details[1])

        prev_data = annotate_data_previous_year(details, cur=cur)
        item = item + prev_data
        wb_sheet.append(item)
        result_wb['общая'].append(item)

result_wb.save(filename=dest_filename)

con.commit()
con.close()
