from datetime import datetime
import excel_scripts as es
import google_scripts as gs
import pandas as pd
import os
pd.set_option('display.max_columns', None)

"""
Модуль формирования уведомления для пользователя
"""

def notification():
    """
    Метод формирования уведомления для пользователей
    Возвращает строку-уведомление для последующей записи в файл.
    """
    gs.load_tasks() #выгружает поручения

    # ДОБАВИТЬ ПРОВЕРКУ НА ПУСТОТУ
    data_poruch = es.filter_tasks(pd.read_excel(es.read_dk_tasks('поручения'))).sort_values(by='Просроченность') #формирует отсортированный список с поручениями
    data_kt = es.filter_dk(pd.DataFrame(pd.read_html(es.read_dk_tasks("Экспорт"))[0])).sort_values(by='Крайний срок') #формирует отсортированный список с КТ
    message = f'Дата: {datetime.today().strftime("%d-%m-%Y")}\n'
    message += '**Скоро будут просрочены следующие поручения**\n'

    def format_poruch(row):
        sphere = row['Тема'] #Заголовок
        overdue_days = row['Просроченность'] #Подзаголовок
        tasks = row[['№', 'Инициатор поручения ', 'Текст поручения']] #Строки

        tasks_str = '\n'.join(
            f"            {id+1}) №: '{num}', Инициатор: '{initiator}',\n            Текст: '{text}'"
            for id, num, initiator, text in
            zip(range(tasks.shape[0]), tasks['№'], tasks['Инициатор поручения '],tasks['Текст поручения']))

        return f"\nТема '{sphere}':\n      Осталось {abs(int(overdue_days))} дней:\n{tasks_str}\n"

    if data_poruch.empty is not True:
        grouped_poruch = data_poruch.groupby(['Тема', 'Просроченность']).agg(list).reset_index()  # Группирует список по теме и просроченности
        message += ''.join(grouped_poruch.apply(format_poruch, axis=1))
    else:
        message += '\nПоручения с истекающим сроком не найдены. Сообщим о новых завтра :)\n'

    message += '\n**Скоро будут просрочены следующие КТ**\n'


    def format_kt(row):
        overdue_days = datetime.strptime(row['Крайний срок'].split()[0],'%d.%m.%Y').strftime("%d-%m-%Y") #Заголовок
        tasks = row[['ID', 'Название', 'Теги']] #Строки

        tasks_str = '\n'.join(
            f"            ID: '{id}', Объект: '{teg}',\n            Содержание: '{text}'"
            for id, text, teg in zip(tasks['ID'], tasks['Название'], tasks['Теги'])
        )

        return f"\nКрайний срок '{overdue_days}':\n{tasks_str}\n"

    if data_kt.empty is not True:
        grouped_kt = data_kt.groupby(['Крайний срок']).agg(list).reset_index()  # Группирует КТ по крайнему сроку
        message += ''.join(grouped_kt.apply(format_kt, axis=1))
    else:
        message += '\nКТ с истекающим сроком не найдены. Сообщим о новых завтра :)\n'
    return message

def write_current_date():
    """
    Метод открывает и записывать в файл сформированный notification
    """
    with open(os.path.join('message.txt'), 'w',encoding='utf-8') as file:
        file.write(notification())
