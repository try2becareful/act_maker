import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches
import time
import random
from datetime import datetime
import os
import PySimpleGUI as sg
import openpyxl

data = pd.read_excel('Таблица данные.xlsx')


'''
data['Дата выдачи'] = pd.to_datetime(data['Дата выдачи'])
data['Начало срока действия'] = str(data['Начало срока действия'])
data['окончание срока действия'] = str(data['окончание срока действия'])
'''


sg.theme('DarkAmber')


def mkd(n):
    tmp = DocxTemplate("act_tmp — копия.docx")
    new_act = {1: {"act_number": idn, "date": date, "position_of_broadcast": br_dolzh, "name_of_broadcast": name,
                   "position_of_receiving": dolzh, "name_of_receiving": fio, "token": sert,
                   "name_of_broadcast_short": sh_br, "name_of_receiving_short": sh_re},
               }
    wait = time.sleep(random.randint(1, 2))
    tmp.render(new_act[n])
    tmp.save('%s.docx' % str(n))


def short_name(s):
    s = s.split()
    new_s = s[0] + ' ' + s[1][0] + '.' + s[2][0] + '.'
    return new_s


def make_win0():
    layout = [
        [sg.Button('Формирование акта')],
        [sg.Button('Добавить пользователя в таблицу')]
    ]
    return sg.Window('Выбор действия', layout, finalize=True)


def make_win1():
    layout = [
        [sg.Text('Подразделение'), sg.InputText()],
        [sg.Text('Фамилия Имя Отчество'), sg.InputText()],
        [sg.Text('Должность'), sg.InputText()],
        [sg.Text('Номер заявки'), sg.InputText()],
        [sg.Text('Тип, номер носителя'), sg.InputText()],
        [sg.Text('Дата выдачи'), sg.InputText()],
        [sg.Text('Акт №/номер сейф-пакета'), sg.InputText()],
        [sg.Text('сертификат'), sg.InputText()],
        [sg.Text('Начало срока действия'), sg.InputText()],
        [sg.Text('Окончание срока действия'), sg.InputText()],
        [sg.Button('Ok'), sg.Button('Выход'), sg.Button('Формирование акта')]
    ]
    return sg.Window('Добавление пользователя', layout, finalize=True)


def make_win2():
    layout = [
        [
            sg.Text('Данные добавленны в таблицу')
        ]
    ]
    return sg.Window('Добавление пользователя', layout, finalize=True)


def make_win3():
    layout = [
        [sg.Text('Введите имя передавателя'), sg.InputText()],
        [sg.Text('Введите должность'), sg.InputText()],
        [sg.Text('Введите номер строки'), sg.InputText()],
        [sg.Button('Сформировать акт')]
    ]
    return sg.Window('Формирование акта', layout, finalize=True)


def make_win4():
    layout = [
        [
            sg.Text('Акт добавлен в папку, где находится программа')
        ]
    ]
    return sg.Window('Акт свормирован', layout, finalize=True)


window0 = make_win0()
window1 = None
window2 = None
window3 = None
window4 = None
while True:
    window, event, values = sg.read_all_windows()
    if event == sg.WIN_CLOSED or event == 'Выход':  # if user closes window or clicks cancel
        window.close()
        if window == window2:
            window2 = None
        elif window == window1:  # if closing win 1, exit program
            break
    elif event == 'Добавить пользователя в таблицу':
        window1 = make_win1()
    elif event == 'Ok' and not window2:
        new_user = {'Подразделение': values[0], 'Фамилия Имя Отчество': values[1], 'Должность': values[2],
                    'номер заявки': values[3], 'тип, номер носителя': values[4], 'Дата выдачи': str(values[5]),
                    'Акт №/номер сейф-пакета': values[6], 'сертификат': values[7],
                    'Начало срока действия': str(values[8]), 'окончание срока действия': str(values[9])}
        data = data.append(new_user, ignore_index=True)
        data.to_excel('./Таблица данные.xlsx')
        window2 = make_win2()
    elif event == 'Формирование акта':
        window3 = make_win3()
    elif event == 'Сформировать акт':
        idn = int(values[2])
        name = values[0]
        br_dolzh = values[1]
        pod = data.at[idn, 'Подразделение']
        fio = data.at[idn, 'Фамилия Имя Отчество']
        dolzh = data.at[idn, 'Должность']
        num_za = data.at[idn, 'номер заявки']
        typeofn = data.at[idn, 'тип, номер носителя']
        date_v = data.at[idn, 'Дата выдачи']
        number_akt = data.at[idn, 'Акт №/номер сейф-пакета']
        sert = data.at[idn, 'сертификат']
        start = data.at[idn, 'Начало срока действия']
        finish = data.at[idn, 'окончание срока действия']
        sh_br = short_name(name)
        sh_re = short_name(fio)
        cur_date = datetime.now()
        date = str(cur_date.day) + '-' + str(cur_date.month) + '-' + str(cur_date.year)
        mkd(1)
        window4 = make_win4()
window.close()
