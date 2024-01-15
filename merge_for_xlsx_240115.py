"""
Merge двух таблиц по двум столбцам
    Version_240115
"""
from pandas import read_excel, merge

import tkinter as tk
from tkinter import filedialog

from datetime import date
from os import startfile, path
import sys


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = path.abspath(".")
    return path.join(base_path, relative_path)


def find_file1():
    filepath = filedialog.askopenfilename()
    file1_text.set(filepath)


def find_file2():
    filepath = filedialog.askopenfilename()
    file2_text.set(filepath)


def add_entry():
    global number_entry
    number_entry += 1

    if number_entry == 1:
        entry_1_text = tk.StringVar()
        entry_1_text.set("Y")
        globals()[f"entry_{number_entry}"] = tk.Entry(frame3, textvariable=entry_1_text)

    else:
        globals()[f"entry_{number_entry}"] = tk.Entry(frame3)

    globals()[f"entry_{number_entry}"].pack(fill='both')


def del_entry():
    global number_entry

    if number_entry > 0:
        globals()[f"entry_{number_entry}"].destroy()
        number_entry -= 1



def process_file():

    df1 = read_excel(file1_text.get())
    df2 = read_excel(file2_text.get())
    file_output = var3_entry.get()

    cols = []
    for i in range(number_entry + 1):
        cols.append(globals()[f"entry_{i}"].get())

    merge(df1, df2,
          on=cols,
          how=method.get(),
          suffixes=('_file1', '_file2')
          ).to_excel(file_output)

    startfile(file_output)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Сопоставление данных") #заголовок окна

    base_path = resource_path("picture")
    root.iconbitmap(path.join(base_path, "my_icon.ico"))

    tk.Label(root, text="Merge двух excel таблиц по столбцам", bd=10).pack(fill='both')

    # Создание фреймов
    frame1 = tk.Frame(root, bd=10)
    frame2 = tk.Frame(root, bd=10)
    frame3 = tk.Frame(root, bd=10)
    frame4 = tk.Frame(root, bd=10)
    frame5 = tk.Frame(root, bd=10)

    frame1.pack(fill='both')
    frame2.pack(fill='both')
    frame3.pack(fill='both')
    frame4.pack(fill='both')
    frame5.pack(fill='both')

    """
    Таблица 1
    """
    tk.Label(frame1, text="Введите путь к файлу 1:").pack(fill='both')
    tk.Button(frame1, text="файл 1", command=find_file1).pack(side='right')

    file1_text = tk.StringVar()
    file1_text.set("для ввода")
    file1_entry = tk.Entry(frame1, textvariable=file1_text)
    file1_entry.pack(expand=True)

    """
    Таблица 2
    """
    tk.Label(frame2, text="Введите путь к файлу 2:").pack(fill='both')
    tk.Button(frame2, text="файл 2", command=find_file2).pack(side='right')

    file2_text = tk.StringVar()
    file2_text.set("для ввода")
    file2_entry = tk.Entry(frame2, textvariable=file2_text)
    file2_entry.pack(expand=True)

    """
    Столбцы merge
    """
    tk.Label(frame3, text="Введите имя столбцов").pack(fill='both')
    tk.Button(frame3, text="Добавить", command=add_entry).pack(fill='both')
    tk.Button(frame3, text="Удалить", command=del_entry).pack(fill='both')

    entry_0_text = tk.StringVar()
    entry_0_text.set("X")
    entry_0 = tk.Entry(frame3, textvariable=entry_0_text)
    entry_0.pack(fill='both')

    number_entry = 0

    """
    Тип merge с помощью радиокнопок
    """
    tk.Label(frame4, text="Тип группировки").pack(fill='both')

    method = tk.StringVar(value="inner")

    all_methods = [("Все строки", "outer"),
                   ("Строки из файла 1", "left"),
                   ("Строки из файла 2", "right"),
                   ("Только пересечения", "inner")
                   ]

    photos = []

    for row in all_methods:
        txt, value = row
        png_path = path.join(base_path, value + ".png")
        photos.append(tk.PhotoImage(file=png_path))
        tk.Radiobutton(frame4,
                       text=txt,
                       value=value,
                       variable=method,
                       image=photos[-1]
                       ).pack(anchor=tk.N)

    """
    Виджеты для файла сохранения
    """
    tk.Label(frame5, text="Имя файла для сохранения в формате .xlsx").pack(fill='both')
    var3_text = tk.StringVar()
    var3_text.set(f"result_{date.today()}.xlsx")
    var3_entry = tk.Entry(frame5, textvariable=var3_text)
    var3_entry.pack(fill='both')

    # Кнопка для запуска обработки
    tk.Button(root, text="Запустить", height=3, bg="black", fg="white", command=process_file).pack(fill='both')

    root.mainloop()
