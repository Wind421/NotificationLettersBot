from datetime import time
import asyncio
import schedule
from aiogram import Bot
from config import TOKEN
import pandas as pd
import form_notification as fn

bot = Bot(token=TOKEN)

async def send_message():
    with open(r'C:\Users\Дарья\PycharmProjects\TgBotNotifications\message.txt', 'r', encoding='ANSI') as file:
        file_content = file.read()

    df = pd.read_excel(r'C:\Users\Дарья\PycharmProjects\TgBotNotifications\users.xlsx')

    for value in df.iloc[:, 0]:
        await bot.send_message(value,file_content)

def schedule_message():
    schedule.every().day.at("10:00").do(lambda: asyncio.create_task(send_message()))

async def main():
    fn.write_current_date()
    schedule_message()
    while True:
        schedule.run_pending()
        await asyncio.sleep(60)

if __name__ == "__main__":
    asyncio.run(main())
