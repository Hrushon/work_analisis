"""Модуль с функциями управления БД."""


def create_table(*args, cur=None):
    """
    Создает таблицы в БД, получая на вход неименованные аргументы с
    названием таблицы и объект Cursor.
    """
    for name in args:
        cur.executescript(f'''
            CREATE TABLE IF NOT EXISTS {name}(
                id INTEGER PRIMARY KEY,
                type TEXT,
                detail TEXT,
                reason TEXT,
                respondent TEXT,
                month INTEGER,
                year INTEGER,
                mileage REAL
            );
        ''')


def add_data_in_tables(name, cur=None, data=None):
    """
    Добавляет данные в таблицу получая в качестве агрументов имя таблицы,
    объект Cursor и, собственно, сами данные.
    """
    cur.executemany(f'''
        INSERT INTO {name} (
        type, detail, reason, respondent, month, year, mileage
        ) VALUES(?, ?, ?, ?, ?, ?, ?);''', data
    )


def annotate_data_in_tables(name, current_month, current_year, cur=None):
    """
    Предоставляет данные из таблицы получая в качестве агрументов имя таблицы,
    объект Cursor. Данные группируются по полю detail и сортируются по
    убыванию.
    """
    cur.execute(f'''
        SELECT type, detail, COUNT(*) AS total
        FROM {name}
        WHERE year = {current_year} AND month <= {current_month}
        GROUP BY detail
        ORDER BY total DESC
    ''')
    return cur.fetchall()


def annotate_data_previous_year(
    details, current_month, current_year, cur=None
):
    """
    Предоставляет количество отказов из таблицы за четыре предыдущих года,
    получая в качестве агрументов кортеж из типа оборудования и отказавшего
    узла, объект Cursor. Данные группируются по полю detail.
    """
    name, detail = details
    result = []

    for i in range(5):
        cur.execute(f'''
            SELECT COUNT(*) AS total
            FROM {name}
            WHERE year = {current_year - i} AND month <= {current_month}
            AND detail = '{detail}'
            GROUP BY detail
        ''')
        data = cur.fetchone()
        if data is None:
            result.append(0)
        else:
            result.append(data[0])
        cur.execute(f'''
            SELECT MAX(mileage) AS result
            FROM {name}
            WHERE year = {current_year - i} AND month <= {current_month}
        ''')
        data = cur.fetchone()
        if data[0] is None:
            result.append(0)
        else:
            result.append(data[0])
    return tuple(result)
