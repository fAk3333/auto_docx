from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from docx_proccessing import *


def activate_start_button():
    name_of_project = name.get()
    cipher = shifr.get()
    customer = zakazchick.get()
    name_of_obj = name_of_object.get()
    leng = length.get()

    print(name_of_project, cipher, customer, name_of_obj, leng)
    if len(name_of_project) == 0 or len(cipher) == 0 or len(customer) == 0 or len(name_of_obj) == 0 or len(leng) == 0:
        messagebox.showerror('Error', 'заполните ВСЕ поля!!!')
    else:
        folder_creater(name_of_project)
        word_processor(name_of_project, cipher, customer, name_of_obj, leng)
        #docx_to_pdf_to_jpeg(name_of_project)

#General
root = Tk()
root.title('docx_auto')
root.geometry("1395x165+200+400")

#varaibles
name_of_object = StringVar(value='')
shifr = StringVar(value='')
zakazchick = StringVar(value='')
name = StringVar(value='')
length = StringVar(value='')

#Text_fields
entry_of_name_of_project = ttk.Entry(textvariable=name_of_object, width=200)
entry_of_shifr = ttk.Entry(textvariable=shifr, width=200)
entry_of_zakazchick = ttk.Entry(textvariable=zakazchick, width=200)
entry_of_project = ttk.Entry(textvariable=name, width=200)
entry_of_length = ttk.Entry(textvariable=length, width=200)

#init
entry_of_project.grid(row=0, column=1)
entry_of_shifr.grid(row=1, column=1)
entry_of_name_of_project.grid(row=2, column=1)
entry_of_zakazchick.grid(row=3, column=1)
entry_of_length.grid(row=4, column=1)

#labels
name_of_prof_label = ttk.Label(text='Название проекта:', font=('Arial', 15))
name_of_prof_label.grid(row=0, column=0)

vert_scale_label = ttk.Label(text='Шифр:', font=('Arial', 15))
vert_scale_label.grid(row=1, column=0)

min_otmetka_label = ttk.Label(text='Название объекта:', font=('Arial', 15))
min_otmetka_label.grid(row=2, column=0)

min_zakazchick = ttk.Label(text='Заказчик', font=('Arial', 15))
min_zakazchick.grid(row=3, column=0)

leng_label = ttk.Label(text='Длина', font=('Arial', 15))
leng_label.grid(row=4, column=0)
#Buttons
start_button = ttk.Button(text="Начать", command=activate_start_button)
start_button.grid(row=5, column=1)


root.mainloop()
