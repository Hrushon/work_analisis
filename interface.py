import tkinter as tk
from tkinter import filedialog

excel_file_name = ''
current_month = 0
current_year = 0


def excel_select():
    global excel_file_name
    comm = filedialog.askopenfilename(
        filetypes=(('Файлы Excel', '*.xls;*.xlsx'),)
    )
    label = tk.Label(text=f'Путь к файлу Excel: {str(comm)}')
    label.pack()
    excel_file_name = comm


def add_data_in_db():
    from main import add_data_in_db_func
    try:
        add_data_in_db_func(excel_file_name)
    except Exception as error:
        label = tk.Label(
            text=f'Что-то пошло не так! Попробуйте еще раз! {error}'
        )
        label.pack()
    else:
        label = tk.Label(text='<<<<<Данные успешно загружены!>>>>>')
        label.pack()    


def get_date():
    global current_month
    global current_year
    current_year = int(e1.get())
    current_month = int(e2.get())
    label = tk.Label(
        text=(f'Будет собрана статистика по {current_month} '
              f'месяц {current_year} года.')
    )
    label.pack()


def analyse():
    from main import analyse_func
    try:
        analyse_func(current_month, current_year)
    except Exception as error:
        label = tk.Label(
            text=f'Что-то пошло не так! Попробуйте еще раз! {error}'
        )
        label.pack()
    else:
        label = tk.Label(text='<<<<<Статистика успешно собрана в файл!>>>>>')
        label.pack()


def out_off():
    raise SystemExit


window = tk.Tk(className=' AvtoAnalizis ')
button1 = tk.Button(
    text='Выберите файл Excel',
    command=excel_select
)
button2 = tk.Button(
    text='Загрузить данные из файла в БД',
    command=add_data_in_db
)
label1 = tk.Label(text='Введите интересующий вас год')
e1 = tk.Entry(width=4)
label2 = tk.Label(text='Введите интересующий вас месяц')
e2 = tk.Entry(width=2)
button3 = tk.Button(
    text='Запомнить год и месяц',
    command=get_date
)
button4 = tk.Button(
    text='Собрать статистику в файл',
    command=analyse
)
button5 = tk.Button(
    text='Выйти',
    command=out_off
)


button1.pack()
button2.pack()
label1.pack()
e1.pack()
label2.pack()
e2.pack()
button3.pack()
button4.pack()
button5.pack()
window.mainloop()
