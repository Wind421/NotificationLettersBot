import excel_scripts as es

from aiogram import Bot, Dispatcher
from aiogram.types import Message,InlineKeyboardButton,InlineKeyboardMarkup,CallbackQuery,Document
from aiogram.filters import Command
from aiogram.enums import ParseMode
from aiogram.enums.content_type import ContentType
from aiogram.fsm.state import State, StatesGroup
import asyncio

import google_requests as gr
import google_scripts as gs

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

#region Menu functions
@dp.callback_query(lambda call: call.data in ['enter_letter','yes_enter'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте входящее письмо.\nОБЯЗАТЕЛЬНО: ВР\nНе забудьте указать суть письма')
    await state.set_state(Form.enter_letter)

@dp.callback_query(lambda call: call.data in ['outer_letter','yes_outer'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте исходящее письмо.\nОБЯЗАТЕЛЬНО: ВР\nНе забудьте указать суть письма и Вр-ответа, если есть')
    await state.set_state(Form.outer_letter)

@dp.callback_query(lambda call: call.data in ['request_letter','yes_request'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте направленный запрос.\nНе забудьте указать RP и суть запроса')
    await state.set_state(Form.request_letter)
#endregion Menu functions

#region Send_letters
@dp.message(Form.enter_letter)
async def process_letter(message: Message, state: Form):

    global COUNTMESS_USER
    if message.caption is not None:
        await state.update_data(letter=gr.wrap_enterletter(message.caption))
        await message.answer("Идет загрузка ⏳")
        COUNTMESS_USER = 0

    elif message.text is not None:
        await state.update_data(letter=gr.wrap_enterletter(message.text))
        await message.answer("Идет загрузка ⏳")
        COUNTMESS_USER = 0

    if COUNTMESS_USER == 0:
        data = await state.get_data()
        if not data['letter']:
            await message.answer("Неправильный формат письма.")
        else:
            result = gs.post_enterletter(data['letter'])
            if not result:
                await message.answer("Это письмо уже есть в таблице.")
            else:
                await message.answer("Письмо сохранено!")

        await state.clear()
        COUNTMESS_USER=1
        markup = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text='Да', callback_data="yes_enter"),
             InlineKeyboardButton(text='Нет', callback_data="no")],
        ])
        await bot.send_message(chat_id=message.chat.id,
                               text="Отправить одно входящее письмо?",
                               parse_mode=ParseMode.MARKDOWN,
                               reply_markup=markup)

@dp.message(Form.outer_letter)
async def process_letter(message: Message, state: Form):

    if message.caption is not None:
        data = await state.get_data()
        if data:
            await state.update_data(letter=gr.wrap_outerletter(message.caption,data['letter']))
        else:
            await state.update_data(letter=gr.wrap_outerletter(message.caption))
    elif message.text is not None:
        data = await state.get_data()
        if data:
            await state.update_data(letter=gr.wrap_outerletter(message.text,data['letter']))
        else:
            await state.update_data(letter=gr.wrap_outerletter(message.text))


    data = await state.get_data()
    if data['letter'][1][0]:
        await message.answer("Идет загрузка ⏳")
        if not data['letter']:
            await message.answer("Неправильный формат письма.")
        else:
            result = gs.post_outerletter(data['letter'])
            if not result:
                await message.answer("Это письмо уже есть в таблице.")
            else:
                await message.answer("Письмо сохранено!")

        markup = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text='Да', callback_data="yes_outer"),
                InlineKeyboardButton(text='Нет', callback_data="no")],
        ])
        await bot.send_message(chat_id=message.chat.id,
                               text="Отправить ещё одно исходящее письмо?",
                               parse_mode=ParseMode.MARKDOWN,
                               reply_markup=markup)
        await state.clear()

@dp.message(Form.request_letter)
async def process_letter(message: Message, state: Form):
    global COUNTMESS_USER
    if message.caption is not None:
        await state.update_data(letter=gr.wrap_request(message.caption))
        await message.answer("Идет загрузка ⏳")
        COUNTMESS_USER = 0
    elif message.text is not None:
        await state.update_data(letter=gr.wrap_request(message.text))
        await message.answer("Идет загрузка ⏳")
        COUNTMESS_USER = 0

    if COUNTMESS_USER == 0:
        data = await state.get_data()
        if not data['letter']:
            await message.answer("Неправильный формат запроса.")
        else:
            result = gs.post_request(data['letter'])
            if not result:
                await message.answer("Этот запрос уже есть в таблице.")
            else:
                await message.answer("Письмо сохранено!")

        await state.clear()
        COUNTMESS_USER = 1
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
