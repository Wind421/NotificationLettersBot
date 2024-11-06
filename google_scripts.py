from datetime import datetime

import gspread
from config import CREDENTIALS_FILENAME, PORUCH_SPREADSHEET_URL, MESSAGE_SPREADSHEET_URL
import pandas as pd
import re

def load_tasks():
    gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
    sh = gc.open_by_url(PORUCH_SPREADSHEET_URL)

    worksheet = sh.get_worksheet(0)
    data = worksheet.get_all_values()
    data = pd.DataFrame(data[3:],columns = data[2])

    def clean_string(s):
        return re.sub(r'[^a-zA-Zа-яА-Я0-9 (),.?!–;:]', '', s)

    df_cleaned = data.map(clean_string)
    filename = rf'C:\Users\Дарья\PycharmProjects\excel\Кампус-поручения_{datetime.today().strftime("%Y-%m-%d")}.xlsx'
    df_cleaned.to_excel(filename,index=False)

def post_enterletter(letter):
    gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
    sh = gc.open_by_url(MESSAGE_SPREADSHEET_URL)
    worksheet = sh.worksheet('Вр входящее')

    vr = letter[1]
    vr_column = worksheet.col_values(2)
    if not any(vr == item or (item.lower().startswith("вр-") and item[3:] == vr) for item in vr_column):
        row_data = (letter[1], str(datetime.now().strftime("%d.%m.%Y")),'', letter[0], '', '', letter[2])
        worksheet.append_row(row_data)
        worksheet.format("D", {"wrapStrategy": "WRAP"})
        return True
    else:
        return False

def post_outerletter(letter):
    gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
    sh = gc.open_by_url(MESSAGE_SPREADSHEET_URL)
    worksheet = sh.worksheet('Вр исходящее')

    vr = letter[1][0]
    vr_column = worksheet.col_values(2)
    if not any(vr == item or (item.lower().startswith("вр-") and item[3:] == vr) for item in vr_column):
        ansvr = letter[1][1]
        if ansvr:
            row_data=('',letter[1][0],'',letter[0],'на согласовании', '', f'Ответ на Вр-{ansvr}')
        else:
            row_data = ('',letter[1][0], '', letter[0], 'на согласовании', '', '')
        worksheet.append_row(row_data)
        worksheet.format("D", {"wrapStrategy": "WRAP"})
        return True
    else:
        return False
