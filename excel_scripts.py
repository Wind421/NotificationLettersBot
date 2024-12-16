from datetime import datetime

import pandas as pd
import os

"""
Модуль для работы с экселем - 
    фильтрафия таблиц, 
    выгрузка экселя, 
    запись юзеров в эксель
"""
def read_dk_tasks(what):
    """
    Метод выгрузки данных из таблицы с поручениями или КТ -
    по заданному пути находит самый последний модифицированный файл и экспортирует путь до него
    """
    export_files = []
    directory = os.path.join('.', 'excel')
    for root, dirs, files in os.walk(directory):
        for file in files:
            if what in file:
                file_path = os.path.join(root, file)
                last_modified_time = datetime.fromtimestamp(os.path.getmtime(file_path)) #дата модификации
                export_files.append((file_path, last_modified_time))

    export_files.sort(key=lambda x: x[1],reverse = True) #сортировка и экспорт пути файла

    if (datetime.today()-export_files[0][1]).days > 7 and what == 'Экспорт':
        with open(os.path.join('.', 'message.txt'), 'w', encoding='ANSI') as file:
            file.write('Требуется обновление\n')
    print(export_files[0][0])
    return export_files[0][0]

def filter_dk(df):
    """
    Метод фильтрации КТ - сформированный df сортирует по завершенности и крайнему сроку
    """
    mask_status = ('Завершена' not in df['Статус'])
    today = datetime.today()
    mask_deadline = df['Крайний срок'].apply(lambda x: (pd.to_datetime(x,dayfirst=True) - today).days in [1, 3, 7, 10, 14, 30, 0])
    filtered_df = df[mask_status & mask_deadline]
    return filtered_df

def filter_tasks(df):
    """
    Метод фильтрации поручений - сформированный df сортирует по завершенности, приоритету и просроченности
    """
    mask_status = ('Выполнено' not in df['Статус'])
    mask_status2 = ('Неактуально' not in df['Статус'])
    mask_importance = (df['Приоритет (1 - высокий, 2 - средний, 3 - низкий)'] == 1)
    mask_deadline = df['Просроченность'].apply(lambda x: x in ['-1', '-3', '-7', '-10', '-14', '-30', '0'])
    filtered_df = df[mask_status & mask_status2 & mask_importance & mask_deadline]
    return filtered_df

def post_user(user_id, first_name):
    """
    Добавляет новых пользователей в таблицу user при запуске бота,
    учитывает если бот запускался \start несколько раз. В таблицу добавляет id и имя
    """
    filename = os.path.join('.', 'users.xlsx')
    try:
        df = pd.read_excel(filename)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['user_id', 'first_name'])

    if not df[df['user_id'] == user_id].empty:
        return

    new_user = pd.DataFrame({'user_id': [user_id], 'first_name': [first_name]})
    df = pd.concat([df, new_user], ignore_index=True)

    df.to_excel(filename, index=False)
