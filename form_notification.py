from datetime import datetime
import excel_scripts as es
import google_scripts as gs
import pandas as pd
pd.set_option('display.max_columns', None)

def notification():
    gs.load_tasks()

    data_poruch = es.filter_tasks(pd.read_excel(es.read_dk_tasks('поручения'))).sort_values(by='Просроченность')
    data_kt = es.filter_dk(pd.DataFrame(pd.read_html(es.read_dk_tasks("Экспорт"))[0])).sort_values(by='Крайний срок')
    message = f'Дата: {datetime.today()}\n'
    message += '**Скоро будут просрочены следующие поручения**\n'
    grouped_poruch = data_poruch.groupby(['Тема', 'Просроченность']).agg(list).reset_index()
   
    def format_poruch(row):
        sphere = row['Тема']
        overdue_days = row['Просроченность']
        tasks = row[['№', 'Инициатор поручения ', 'Текст поручения']]

        tasks_str = '\n'.join(
            f"            {id+1}) №: '{num}', Инициатор: '{initiator}',\n            Текст: '{text}'"
            for id, num, initiator, text in
            zip(range(tasks.shape[0]), tasks['№'], tasks['Инициатор поручения '],tasks['Текст поручения']))

        return f"\nТема '{sphere}':\n      Осталось {overdue_days} дней:\n{tasks_str}\n"
    message += ''.join(grouped_poruch.apply(format_poruch, axis=1))

    message += '\n**Скоро будут просрочены следующие КТ**\n'
    grouped_kt = data_kt.groupby(['Крайний срок']).agg(list).reset_index()

    def format_kt(row):
        overdue_days = row['Крайний срок']
        tasks = row[['ID', 'Название', 'Теги']]

        tasks_str = '\n'.join(
            f"            ID: '{id}', Объект: '{teg}',\n            Содержание: '{text}'"
            for id, text, teg in zip(tasks['ID'], tasks['Название'], tasks['Теги'])
        )

        return f"\nКрайний срок '{overdue_days}':\n{tasks_str}\n"

    message += ''.join(grouped_kt.apply(format_kt, axis=1))
    return message

def write_current_date():
    with open('message.txt', 'w') as file:
        file.write(notification())

