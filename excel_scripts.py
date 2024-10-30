from datetime import datetime

import pandas as pd
import os
#pd.set_option('display.max_columns', None)

def read_dk_tasks(what):
    export_files = []
    directory = r'C:\Users\Дарья\PycharmProjects\excel'
    for root, dirs, files in os.walk(directory):
        for file in files:
            if what in file:
                file_path = os.path.join(root, file)
                last_modified_time = os.path.getmtime(file_path)
                export_files.append((file_path, last_modified_time))

    export_files.sort(key=lambda x: x[1],reverse = True)
    return export_files[0][0]

def filter_dk(df):
    mask_status = ('Завершена' not in df['Статус'])
    today = datetime.today()
    mask_deadline = df['Крайний срок'].apply(lambda x: (pd.to_datetime(x,dayfirst=True) - today).days in [1, 3, 7, 14, 30])
    filtered_df = df[mask_status & mask_deadline]
    return filtered_df

def filter_tasks(df):
    mask_status = ('Выполнено' not in df['Статус'])
    mask_status2 = ('Неактуально' not in df['Статус'])
    mask_importance = (pd.isna(df['Важность']))
    mask_deadline = df['Просроченность'].apply(lambda x: x in ['-1', '-3', '-7', '-30','0'])
    filtered_df = df[mask_status & mask_status2 & mask_importance & mask_deadline]
    return filtered_df

def post_user(user_id, first_name):
    filename = r'C:\Users\Дарья\PycharmProjects\TgBotNotifications\users.xlsx'
    try:
        df = pd.read_excel(filename)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['user_id', 'first_name'])

    if not df[df['user_id'] == user_id].empty:
        return

    new_user = pd.DataFrame({'user_id': [user_id], 'first_name': [first_name]})
    df = pd.concat([df, new_user], ignore_index=True)

    df.to_excel(filename, index=False)


