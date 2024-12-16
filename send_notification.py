import os
import asyncio

import pandas as pd
import schedule
from aiogram import Bot
from aiogram.enums import ParseMode
from aiogram.exceptions import TelegramForbiddenError

import form_notification as fn
from config import TOKEN

bot = Bot(token=TOKEN)

"""
Модуль для отправки уведомления пользователю
"""

async def send_message():
    """
    Метод отправки уведомления - выгружает нужный файл, читает и отправляет ботом
    """
    with open(os.path.join('.', 'message.txt'), 'r', encoding='ANSI') as file:
        file_content = file.readlines()

    df = pd.read_excel(os.path.join('.', 'users.xlsx'))

    for value in df.iloc[:, 0]: #Перебор списка с пользователями для отправки сообщения
        try:
            await bot.send_message(value,''.join(file_content[1:]),parse_mode = ParseMode.MARKDOWN)
        except TelegramForbiddenError:
            pass

    if file_content[0] == 'Требуется обновление\n':
        await bot.send_message(1617319542, f'Обновите контрольные точки!\nНажмите: /send_kt')

def schedule_message():
    """
    Метод для установки расписания - должно быть 10-00
    """
    schedule.every().day.at("10:00").do(lambda: asyncio.create_task(send_message()))

async def main():
    """
    Основной метод - формирует уведомление,
    а затем проверяет не пришло ли время его отправлять.
    """
    fn.write_current_date()
    schedule_message()
    while True:
        schedule.run_pending()
        await asyncio.sleep(60)

if __name__ == "__main__":
    asyncio.run(main())

