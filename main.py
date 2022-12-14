import os
import sqlite3

from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_NAME = 'analisis.xlsx'
DB_NAME = 'db.sqlite'
DB_FILE = 'file:' + os.path.join(BASE_DIR, DB_NAME)

uktol = list()
compressors = list()
others = list()

COMPRES_TUPLE = ('вв', 'акв', 'дэн')

CHOICE = {
    'уктол': uktol,
    'мк': compressors,
    'прочее': others
}

wb = load_workbook(filename=os.path.join(BASE_DIR, FILE_NAME), data_only=True)

wb.active = 0

sheet = wb.active

for i in range(1, sheet.max_row):
    col = sheet['A' + str(i + 1)]
    if col.value.lower() in COMPRES_TUPLE:
        target = CHOICE['мк']
    else:
        target = CHOICE.setdefault(col.value.lower(), list())

    result = tuple(
        col[i].value for col in sheet.iter_cols(0, sheet.max_column)
    )
    target.append(result)

con = sqlite3.connect(DB_FILE, uri=True)
cur = con.cursor()

cur.executescript('''
    CREATE TABLE IF NOT EXISTS compressors(
        id INTEGER PRIMARY KEY,
        type TEXT,
        detail TEXT,
        reason TEXT,
        respondent TEXT,
        month INTEGER
    );

    CREATE TABLE IF NOT EXISTS uktol(
        id INTEGER PRIMARY KEY,
        type TEXT,
        detail TEXT,
        reason TEXT,
        respondent TEXT,
        month INTEGER
    );

    CREATE TABLE IF NOT EXISTS others(
        id INTEGER PRIMARY KEY,
        type TEXT,
        detail TEXT,
        reason TEXT,
        respondent TEXT,
        month INTEGER
    );
''')

cur.executemany(
    'INSERT INTO uktol ('
    'type, detail, reason, respondent, month'
    ') VALUES(?, ?, ?, ?, ?);',
    uktol
)
cur.executemany(
    'INSERT INTO compressors ('
    'type, detail, reason, respondent, month'
    ') VALUES(?, ?, ?, ?, ?);',
    compressors
)
cur.executemany(
    'INSERT INTO others ('
    'type, detail, reason, respondent, month'
    ') VALUES(?, ?, ?, ?, ?);',
    others
)

con.commit()
con.close()
