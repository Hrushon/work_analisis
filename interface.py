import tkinter as tk
from tkinter import filedialog

list_excel = ''
current_month = 0
current_year = 0


def excel_select():
    global list_excel
    comm = filedialog.askopenfilename(
        filetypes=(('Файлы Excel', '*.xls;*.xlsx'),)
    )
    label = tk.Label(text=f'Путь к файлу Excel: {str(comm)}')
    label.pack()
    list_excel = comm


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


def analizis():
    from main import main_func
    main_func(list_excel, current_month, current_year)
    label = tk.Label(text='<<<<<Статистика успешно собрана в файл!>>>>>')
    label.pack()


def out_off():
    exit()


window = tk.Tk(className=' AvtoAnalizis ')
button1 = tk.Button(
    text='Выберите файл Excel',
    command=excel_select
)
label1 = tk.Label(text='Введите интересующий вас год')

e1 = tk.Entry(width=4)
label2 = tk.Label(text='Введите интересующий вас месяц')
e2 = tk.Entry(width=2)
button2 = tk.Button(
    text='Запомнить год и месяц',
    command=get_date
)
button3 = tk.Button(
    text='Собрать статистику в файл',
    command=analizis
)
button4 = tk.Button(
    text='Выйти',
    command=out_off
)


button1.pack()
label1.pack()
e1.pack()
label2.pack()
e2.pack()
button2.pack()
button3.pack()
button4.pack()
window.mainloop()
