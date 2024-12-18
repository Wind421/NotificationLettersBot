from datetime import datetime

import gspread
from gspread_formatting import get_effective_format,format_cell_ranges,CellFormat
from config import CREDENTIALS_FILENAME, PORUCH_SPREADSHEET_URL, MESSAGE_SPREADSHEET_URL,TRUE_MESSAGE_SPREADSHEET_URL,TRUE_CVPO_SPREADSHEET_URL
import pandas as pd
import re
import os

textformat_orange = {
    "textFormat": {
      "foregroundColor": {
        "red": 1,
        "green": 0.647,
        "blue": 0
      }
    }}
textformat_black = {
    "textFormat": {
      "foregroundColor": {
        "red": 0,
        "green": 0,
        "blue": 0
      }
    }}
textformat_red = {
    "textFormat": {
      "foregroundColor": {
        "red": 1,
        "green": 0,
        "blue": 0
      }
    }}

"""
Модуль для загрузки оформленных писем в гугл таблиц и выгрузки из них
"""
gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
def load_tasks():
    """
    Метод выгрузки поручений из гугл-таблицы в эксель-таблицу
    """
    sh = gc.open_by_url(PORUCH_SPREADSHEET_URL)

    worksheet = sh.get_worksheet(0)
    data = worksheet.get_all_values()
    data = pd.DataFrame(data[3:],columns = data[2])

    def clean_string(s): #Очистка для корректности создания экселя
        return re.sub(r'[^a-zA-Zа-яА-Я0-9 \-(),.?!;:]', '', s)

    df_cleaned = data.map(clean_string)
    filename = os.path.join('.','excel',f'Кампус-поручения_{datetime.today().strftime("%Y-%m-%d")}.xlsx')
    df_cleaned.to_excel(filename,index=False)

def post_enterletter(letter):
    """
    Метод загрузки входящих писем в гугл-таблицу
    """
    sh = gc.open_by_url(TRUE_MESSAGE_SPREADSHEET_URL)
    worksheet = sh.worksheet('ВР входящие')

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
    sh = gc.open_by_url(TRUE_MESSAGE_SPREADSHEET_URL)
    worksheet = sh.worksheet('ВР исходящие')

    vr = letter[1][0]
    vr_column = worksheet.col_values(2)
    if not any(vr[0] == item or (item.lower().startswith("вр-") and item[3:] == vr) for item in vr_column):#Если найден не такой же вр как и у письма
        ansvr = ', '.join(letter[1][1])
        if ansvr: #Если есть вр ответа
            row_data = ('', str(*vr),str(letter[2]),str(letter[0]),'на согласовании', '', f'Ответ на Вр-{ansvr}')
        else:
            row_data = ('', str(*vr), str(letter[2]), str(letter[0]), 'на согласовании', '', '')
        worksheet.append_row(row_data)
        worksheet.format("D", {"wrapStrategy": "WRAP"})

        row_index = len(vr_column) + 1 #Красим строчку на согласовании в оранжевый
        worksheet.format(f'E{row_index}', textformat_orange)
        return True
    else:
        return False

def post_request(letter):
    """
    Метод загрузки запросов в гугл-таблицу
    """
    sh = gc.open_by_url(TRUE_CVPO_SPREADSHEET_URL)
    worksheet = sh.worksheet('СВПО')

    rp = letter[1]
    rp_column = worksheet.col_values(2)
    if not any(rp == item for item in rp_column): #Если нет таких же запросов
        num_column = worksheet.col_values(1)
        #           Следующий номер       RP         Дата отправки                             Текст          Срок
        row_data = (int(num_column[-1])+1,letter[1], str(datetime.now().strftime("%d.%m.%Y")), letter[0], '', letter[2])
        worksheet.append_row(row_data)
        formatting = get_effective_format(worksheet, f'A{len(rp_column)}:G{len(rp_column)}') #Захват и применение форматирования предыдущей строки таблицы
        format_cell_ranges(worksheet, [(f'A{len(rp_column)+1}:G{len(rp_column)+1}', formatting),(f'D{len(rp_column)+1}', formatting.add(CellFormat(horizontalAlignment='LEFT')))])
        return True
    else:
        return False

def post_ansvr(letter):
    if letter[1][1] != '':
        sh = gc.open_by_url(TRUE_MESSAGE_SPREADSHEET_URL)
        worksheet = sh.worksheet('ВР входящие')
        rp_column = worksheet.col_values(2)
        for vr in letter[1][1]:
            if str(vr) in rp_column:
                row_index = rp_column.index(str(vr)) + 1
                curr = worksheet.cell(row_index, 11).value
                ansvr = ''.join(letter[1][0])
                if curr:
                    worksheet.update_cell(row_index, 11, curr + ',' + str(ansvr))

                else:
                    worksheet.update_cell(row_index, 11, str(ansvr))
        worksheet.format("K", {"wrapStrategy": "WRAP"})

def change(data):
    if data['what'] == 'enter':
        sh = gc.open_by_url(TRUE_MESSAGE_SPREADSHEET_URL)
        worksheet = sh.worksheet('ВР входящие')
        rp_column = worksheet.col_values(2)
        if str(data['vr']) in rp_column:
            row_index = rp_column.index(str(data['vr'])) + 1
            worksheet.update_cell(row_index, 12, str(data['status']))
        else:
            raise KeyError

    elif data['what'] == 'outer':
        sh = gc.open_by_url(TRUE_MESSAGE_SPREADSHEET_URL)
        worksheet = sh.worksheet('ВР исходящие')
        rp_column = worksheet.col_values(2)
        if str(data['vr']) in rp_column:
            row_index = rp_column.index(str(data['vr'])) + 1
            if 'не согл' in data['status'].lower():
                worksheet.update_cell(row_index, 5, str(data['status']))
                worksheet.format(f'E{row_index}', textformat_red)
            elif 'согл' in data['status'].lower():
                worksheet.update_cell(row_index, 5, str(data['status']))
                worksheet.format(f'E{row_index}',textformat_orange)
            elif 'подпис' in data['status'].lower():
                worksheet.update_cell(row_index, 5, str('подписано'))
                worksheet.update_cell(row_index, 6, ' '.join(data['status'].split()[1:]))
                worksheet.format(f'E{row_index}', textformat_black)

            if worksheet.cell(row_index,7) != '':
                worksheet2 = sh.worksheet('ВР входящие')
                rp_column2 = worksheet2.col_values(11)
                for vr in rp_column2:
                    if data['vr'] in vr:
                        row_index2 = rp_column2.index(vr) + 1
                        if 'согл' in data['status']:
                            worksheet2.update_cell(row_index2, 12, 'на согласовании')
                        elif 'подпис' in data['status']:
                            worksheet2.update_cell(row_index2, 12, 'согласовано')
            else:
                raise ValueError
        else:
            raise KeyError

    elif data['what'] == 'request':
        sh = gc.open_by_url(TRUE_CVPO_SPREADSHEET_URL)
        worksheet = sh.worksheet('СВПО')
        rp_column = worksheet.col_values(2)
        if ('RP'+str(data['vr'])) in rp_column:
            if 'выполн' in data['status'].lower():
                row_index = rp_column.index('RP'+str(data['vr'])) + 1
                worksheet.update_cell(row_index, 5, str(data['status']))
                worksheet.update_cell(row_index, 7, str(datetime.now().strftime("%d.%m.%Y")))
        else:
            raise KeyError