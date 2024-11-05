import excel_scripts as es

from aiogram import Bot, Dispatcher
from aiogram.types import Message,InlineKeyboardButton,InlineKeyboardMarkup,CallbackQuery,Document
from aiogram.filters import Command
from aiogram.enums import ParseMode
from aiogram.enums.content_type import ContentType
from aiogram.fsm.state import State, StatesGroup
import asyncio

import google_requests as gr

import logging
logging.basicConfig(level=logging.INFO)

from config import *
import os

bot = Bot(token=TOKEN)
dp = Dispatcher()
BOTMESS_ID = None
COUNTMESS_USER = None

class Form(StatesGroup):
    enter_letter = State()
    outer_letter = State()
    request_letter = State()
    files = State()
    letter = State()

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
    await bot.send_message(message.chat.id, 'Привет, Я работаю')

@dp.callback_query(lambda call: call.data in ['enter_letter','yes_other'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    await bot.answer_callback_query(callback_query.id)
    last_message = await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте исходящее письмо.\nНе забудьте указать ВР и краткое содержание письма,\n а также Вр-ответа если он присутствует')
    global BOTMESS_ID
    BOTMESS_ID = last_message.message_id
    await state.set_state(Form.enter_letter)

#region Menu functions
@dp.callback_query(lambda call: call.data in ['outer_letter','yes_enter'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте исходящее письмо.\nНе забудьте указать ВР и суть запроса')
    await state.set_state(Form.outer_letter)

@dp.callback_query(lambda call: call.data in ['request_letter','yes_request'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте направленный запрос.\nНе забудьте указать RP и суть запроса')
    await state.set_state(Form.request_letter)
#endregion Menu functions

@dp.message(Form.enter_letter)
async def process_letter(message: Message, state: Form):

    global COUNTMESS_USER
    if message.caption is not None:
        await state.update_data(letter=gr.wrap_enter_letter(message.caption))
        last_message = await message.answer("Письмо сохранено!")
        COUNTMESS_USER = last_message.message_id-BOTMESS_ID-1
    elif message.text is not None:
        await state.update_data(letter=gr.wrap_enter_letter(message.text))
        last_message = await message.answer("Письмо сохранено!")
        COUNTMESS_USER = last_message.message_id-BOTMESS_ID-2

    if message.content_type == ContentType.DOCUMENT:
        file_id = message.document.file_id
        file = await bot.get_file(file_id)
        file_path = os.path.join(r"C:\Users\Дарья\PycharmProjects\TgBotNotifications\downloaded_files\\",
                                 message.document.file_name)

        await bot.download_file(file.file_path, file_path)

        data = await state.get_data()
        if not data.get('files'):
            await state.update_data(files = [file_path])
        else:
            await state.update_data(files = list(data['files'])+[file_path])
        COUNTMESS_USER -=1

    if COUNTMESS_USER==0:
        markup = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text='Да', callback_data="yes_other"),
             InlineKeyboardButton(text='Нет', callback_data="no")],
        ])
        await bot.send_message(chat_id=message.chat.id,
                               text="Отправить ещё письмо?",
                               parse_mode=ParseMode.MARKDOWN,
                               reply_markup=markup)
        data = await state.get_data()
        await state.clear()
        print(data)


#region Forms
@dp.message(Form.outer_letter)
async def process_letter(message: Message, state: Form):
    await state.update_data(enter_letter=message.text)
    await bot.send_message(message.chat.id, 'Письмо успешно загружено в таблицу!')
    await state.clear()
    markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text='Да', callback_data="yes_enter"),
         InlineKeyboardButton(text='Нет', callback_data="no")],
    ])
    await bot.send_message(chat_id=message.chat.id,
                           text="Отправить ещё письмо?",
                           parse_mode=ParseMode.MARKDOWN,
                           reply_markup=markup)

@dp.message(Form.request_letter)
async def process_letter(message: Message, state: Form):
    await state.update_data(request_letter=message.text)
    await bot.send_message(message.chat.id, 'Запрос успешно добавлен в таблицу!')
    await state.clear()
    markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text='Да', callback_data="yes_request"),
         InlineKeyboardButton(text='Нет', callback_data="no")],
    ])
    await bot.send_message(chat_id=message.chat.id,
                           text="Отправить ещё запрос?",
                           parse_mode=ParseMode.MARKDOWN,
                           reply_markup=markup)
#endregion

@dp.callback_query(lambda call: call.data == 'no')
async def process_no(callback_query: CallbackQuery):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Нажмите /menu')


if __name__ == '__main__':
    dp.run_polling(bot)
