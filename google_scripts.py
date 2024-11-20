from datetime import datetime

import gspread
from gspread_formatting import get_effective_format,format_cell_ranges,CellFormat
from config import CREDENTIALS_FILENAME, PORUCH_SPREADSHEET_URL, MESSAGE_SPREADSHEET_URL
import pandas as pd
import re

"""
Модуль для загрузки оформленных писем в гугл таблиц и выгрузки из них
"""

def load_tasks():
    """
    Метод выгрузки поручений из гугл-таблицы в эксель-таблицу
    """
    gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
    sh = gc.open_by_url(PORUCH_SPREADSHEET_URL)

    worksheet = sh.get_worksheet(0)
    data = worksheet.get_all_values()
    data = pd.DataFrame(data[3:],columns = data[2])

    def clean_string(s): #Очистка для корректности создания экселя
        return re.sub(r'[^a-zA-Zа-яА-Я0-9 (),.?!–;:]', '', s)

    df_cleaned = data.map(clean_string)
    filename = rf'C:\Users\Дарья\PycharmProjects\excel\Кампус-поручения_{datetime.today().strftime("%Y-%m-%d")}.xlsx'
    df_cleaned.to_excel(filename,index=False)

def post_enterletter(letter):
    """
    Метод загрузки входящих писем в гугл-таблицу
    """
    gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
    sh = gc.open_by_url(MESSAGE_SPREADSHEET_URL)
    worksheet = sh.worksheet('Вр входящее')

    vr = letter[1]
    vr_column = worksheet.col_values(2)
    if not any(vr == item or (item.lower().startswith("вр-") and item[3:] == vr) for item in vr_column): #Если найден не такой же вр как и у письма
        row_data = ('',letter[1], str(datetime.now().strftime("%d.%m.%Y")), letter[3], letter[0], '', 'срок не указан', letter[2])
        worksheet.append_row(row_data)
        worksheet.format("D", {"wrapStrategy": "WRAP"}) #перенос
        return True
    else:
        return False

def post_outerletter(letter):
    """
    Метод загрузки исходящих писем в гугл-таблицу
    """
    gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
    sh = gc.open_by_url(MESSAGE_SPREADSHEET_URL)
    worksheet = sh.worksheet('Вр исходящее')

    vr = letter[1][0]
    vr_column = worksheet.col_values(2)
    if not any(vr[0] == item or (item.lower().startswith("вр-") and item[3:] == vr) for item in vr_column):#Если найден не такой же вр как и у письма
        ansvr = letter[1][1]
        if ansvr: #Если есть вр ответа
            row_data = ('', str(*vr),str(letter[2]),str(letter[0]),'на согласовании', '', f'Ответ на Вр-{ansvr}')
        else:
            row_data = ('', str(*vr), str(letter[2]), str(letter[0]), 'на согласовании', '', '')
        worksheet.append_row(row_data)
        worksheet.format("D", {"wrapStrategy": "WRAP"})
        return True
    else:
        return False

def post_request(letter):
    """
    Метод загрузки запросов в гугл-таблицу
    """
    gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
    sh = gc.open_by_url(MESSAGE_SPREADSHEET_URL)
    worksheet = sh.worksheet('СВПО')

    rp = letter[1]
    rp_column = worksheet.col_values(2)
    if not any(rp == item for item in rp_column): #Если нет таких же запросов
        num_column = worksheet.col_values(1)
        #           Следующий номер       RP         Дата отправки                             Текст          Срок
        row_data = (int(num_column[-1])+1,letter[1], str(datetime.now().strftime("%d.%m.%Y")), letter[0], '', letter[2])
        worksheet.append_row(row_data)
        formatting = get_effective_format(worksheet, f'A{len(rp_column)}:G{len(rp_column)}') #Захват и применение форматирования предыдущей строки таблицы
        print(formatting)
        format_cell_ranges(worksheet, [(f'A{len(rp_column)+1}:G{len(rp_column)+1}', formatting),(f'D{len(rp_column)+1}', formatting.add(CellFormat(horizontalAlignment='LEFT')))])
        return True
    else:
        return False

def post_ansvr(letter):
    if letter[1][1] != '':
        gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
        sh = gc.open_by_url(MESSAGE_SPREADSHEET_URL)
        worksheet = sh.worksheet('Вр входящее')
        rp_column = worksheet.col_values(2)
        if str(letter[1][1]) in rp_column:
            row_index = rp_column.index(str(letter[1][1])) + 1
            worksheet.update_cell(row_index, 11, str(*letter[1][0]))

def change(data):
    gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
    if data[0] == 'enter':
        sh = gc.open_by_url(MESSAGE_SPREADSHEET_URL)
        worksheet = sh.worksheet('Вр входящее')
        rp_column = worksheet.col_values(2)
        if str(data[1]) in rp_column:
            row_index = rp_column.index(str(data[1])) + 1
            worksheet.update_cell(row_index, 12, str(data[2]))
        else:
            raise ValueError

    elif data[0] == 'outer':
        sh = gc.open_by_url(MESSAGE_SPREADSHEET_URL)
        worksheet = sh.worksheet('Вр исходящее')
        rp_column = worksheet.col_values(2)
        if str(data[1]) in rp_column:
            row_index = rp_column.index(str(data[1])) + 1
            worksheet.update_cell(row_index, 5, str(data[2]))
        else:
            raise ValueError

    elif data[0] == 'request':
        sh = gc.open_by_url(MESSAGE_SPREADSHEET_URL)
        worksheet = sh.worksheet('СВПО')
        rp_column = worksheet.col_values(2)
        if str(data[1]) in rp_column:
            row_index = rp_column.index(str(data[1])) + 1
            worksheet.update_cell(row_index, 5, str(data[2]))
            worksheet.update_cell(row_index, 7, str(datetime.now().strftime("%d.%m.%Y")))
        else:
            raise ValueError

