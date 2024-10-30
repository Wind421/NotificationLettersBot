import excel_scripts as es

from aiogram import Bot, Dispatcher
from aiogram.types import Message,InlineKeyboardButton,InlineKeyboardMarkup,CallbackQuery
from aiogram.filters import Command
from aiogram.enums import ParseMode

import logging
logging.basicConfig(level=logging.INFO)

from .callback_defs import *

bot = Bot(token=TOKEN)
dp = Dispatcher()

@dp.message(Command("start"))
async def start_def(message: Message):

    await bot.send_message(
        message.chat.id,
        f'''Привет, {message.from_user.first_name}!\n
Здесь ты будешь получать уведомления в 10:00 по крайним срокам поручений и дорожной карты.\n
А также сможешь отправлять письма и запросы для заполнения их в таблицу.\n
Надеюсь мы сработаемся :)\n
Для вызова меню введите /menu
'''
    )
    user_id = message.from_user.id
    first_name = message.from_user.first_name
    es.post_user(user_id,first_name)

@dp.message(Command(commands=['menu']))
async def menu_def(message: Message):

    markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text='Отправить\nвходящее письмо', callback_data="enter_letter")],
        [InlineKeyboardButton(text='Отправить\nисходящее письмо', callback_data="outer_letter")],
        [InlineKeyboardButton(text='Отправить\nRPзапрос', callback_data="request_letter")],
    ])

    await bot.send_message(chat_id=message.chat.id,
                           text="*________Главное меню________*",
                           parse_mode=ParseMode.MARKDOWN,
                           reply_markup=markup)

@dp.message(Command('work'))
async def send_menu(message: Message):
    await bot.send_message(message.chat.id, 'Я работаю')

if __name__ == '__main__':
    dp.run_polling(bot)
