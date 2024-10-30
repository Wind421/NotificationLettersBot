from datetime import datetime

import gspread
from config import CREDENTIALS_FILENAME, QUESTIONS_SPREADSHEET_URL
import pandas as pd
import re

def load_tasks():
    gc = gspread.service_account(filename=CREDENTIALS_FILENAME)
    sh = gc.open_by_url(QUESTIONS_SPREADSHEET_URL)

    worksheet = sh.get_worksheet(0)
    data = worksheet.get_all_values()
    data = pd.DataFrame(data[3:],columns = data[2])

    def clean_string(s):
        # Оставляем только буквы русского и английского алфавитов, а также пробелы
        return re.sub(r'[^a-zA-Zа-яА-Я0-9 (),.?!–;:]', '', s)

    # Применяем очистку ко всем строкам DataFrame
    df_cleaned = data.map(clean_string)
    filename = rf'C:\Users\Дарья\PycharmProjects\excel\Кампус-поручения_{datetime.today().strftime("%Y-%m-%d")}.xlsx'
    df_cleaned.to_excel(filename,index=False)

load_tasks()

