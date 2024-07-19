from tkinter import *
from tkinter import ttk

import pyodbc

from modules.sql_requests import sql_main
from modules.make_xls import xl_main


window = Tk()
window.geometry('600x600')
window.title('Вывод сводной таблицы')
window.resizable(False, False)

# Стили
header1 = ('Calibry', 12)
header2 = ('Calibry', 10)

# Функция для переключения всех чекбоксов
def toggle_forma22_expenses():
    state = check_box_formа22.get()
    for checkbox in losses_expense_checkboxes:
        losses_checkbox_var[checkbox].set(state)

def get_selected_checkboxes():
    selected = [category for category, var in losses_checkbox_var.items() if var.get() == 1]
    return selected

def create_xls():
    warning_label.config(text="")

    selected_checkboxes = get_selected_checkboxes()
    if len(selected_checkboxes) == 0:
        warning_label.config(text="Формы мне самому выбрать?", fg="#9E9900")
        return 0

    region_dir = {
        "Брестская": 1,
        "Витебская": 2,
        "Гомельская": 3,
        "Гродненская": 4,
        "Минская": 6,
        "Могилевская": 7
    }

    try:
        region = region_dir[combobox.get()]
    except KeyError:
        warning_label.config(text="Регион мне самому выбрать?", fg="#9E9900")
        return 0

    start_time = start_time_entry.get()
    end_time = end_time_entry.get()

    try:
        print(start_time_entry.get())
        rows = sql_main(region, selected_checkboxes, start_time, end_time, read_entry.get())
        xl_main(start_time, end_time, rows, write_entry.get())

    except SyntaxError:
        warning_label.config(
            text="Не правильно введено время!\n"
                 "  Формат: \n"
                 "01-10-2005 \n"
                 "  Или \n"
                 "01-10-2005 02:58:24 \n",
            fg="#E32636")
        return 0
    except PermissionError:
        warning_label.config(text="Закройте файл excel 'summary' и повторите", fg="#9E9900")
        return 0
    except TabError:
        warning_label.config(text="Подходящих записей не найдено", fg="#9E9900")
        return 0
    except FileNotFoundError:
        warning_label.config(text="Проверьте путь для сохранения", fg="#9E9900")
        return 0
    """ except pyodbc.Error:
        warning_label.config(text="Не верный путь к БД", fg="#9E9900")
        return 0"""


    warning_label.config(text="Успешно записано!", fg="#008000")



warning_label = Label(font=("Calibri", 14), justify="left", anchor="w")
warning_label.place(x=50, y=300)

regions = ["Брестская", "Витебская", "Гомельская", "Гродненская", "Минская", "Могилевская"]

Label(text="Выбор региона", font=header1).place(x=50, y=40)
combobox = ttk.Combobox(window, values=regions, height=3, width=20)
combobox.set("Выберите поле")
combobox.place(x=50, y=65)


Label(text="Forma22", font=header1).place(x=400, y=40)

forma22_frame = Frame(window)
forma22_frame.place(x=400, y=60)

check_box_formа22 = IntVar()

losses_checkbox_var = {}
losses_expense_checkboxes = []

Label(forma22_frame, text="Выбрать все", font=header1).place(x=0, y=0)
cb_expenses = Checkbutton(forma22_frame, variable=check_box_formа22, command=toggle_forma22_expenses)
cb_expenses.grid(row=1, sticky="e", column=3)

expense_categories = range(1, 25)

for i, category in enumerate(expense_categories):
    var = IntVar()
    losses_checkbox_var[category] = var
    Label(forma22_frame, text=category, font=header2).grid(row=(i % 11) + 2, column=(i//11)*2, sticky='w', padx=10)
    cb = Checkbutton(forma22_frame, variable=var)
    cb.grid(row=(i % 11) + 2, column=((i//11)+1)*2-1, sticky='e')
    losses_expense_checkboxes.append(category)


Label(window, text="Период времени:", font=header1).place(x=50, y=150)
start_time_label = Label(window, text="Начать с:")
start_time_entry = Entry(window)
start_time_label.place(x=50, y=180)
start_time_entry.place(x=150, y=180)

end_time_label = Label(window, text="Закончить на:")
end_time_entry = Entry(window)
end_time_label.place(x=50, y=210)
end_time_entry.place(x=150, y=210)
entry_comment_label = Label(window, text="(формат: 01-01-2024 00:00:00)", font=('Calibri', 7))
entry_comment_label.place(x=145, y=230)


entry_text = StringVar()
entry_text.set(r"data\DKR.mdb")
read_entry = Entry(window, textvariable=entry_text)
Label(text="Путь до БД:", font=('Calibri', 10)).place(x=280, y=400)
read_entry.place(x=430, y=400)


entry_text = StringVar()
entry_text.set(r"data\summary.xlsx")
write_entry = Entry(window, textvariable=entry_text)
Label(text="Путь сохранения excel: ", font=('Calibri', 10)).place(x=280, y=450)
write_entry.place(x=430, y=450)


output_button = Button(
    bg="#22CC22",
    text="Создать",
    command=create_xls,
    width=9,
    height=2
)
output_button.place(x=270, y=500)


window.mainloop()