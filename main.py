import logging

from aiogram import Bot, Dispatcher
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import Message, InlineKeyboardButton, InlineKeyboardMarkup, CallbackQuery

import excel_scripts as es
import google_requests as gr
import google_scripts as gs

logging.basicConfig(level=logging.INFO)

from config import *

bot = Bot(token=TOKEN)
dp = Dispatcher()

BOTMESS_ID = None #ID - последнего отправленного ботом сообщения
COUNTMESS_USER = None #Количество между последним сообщением пользователя и последним письмом бота

class Form(StatesGroup):
    """
    Класс Form - отслеживает какое именно письмо хочет отправить пользователь,
    а также сохраняет информацию о файлах (пока нет необходимости) и содержании писем
    """
    enter_letter = State()
    outer_letter = State()
    request_letter = State()
    files = State()
    letter = State()

@dp.message(Command("start"))
async def start_def(message: Message):
    """
    Метод команды start - выводит приветственное сообщение и добавляет пользователя
    в таблицу для дальнейшей отправки ему уведомлений
    """
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
    """
    Метод команды menu - выводит клавиатуру для выбора отправки письма или запроса
    """
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
    """
    Метод обработки входящих сообщений - выводит сообщения с предложением для отправки письма
    и устанавливает состояние "отправить входящее письмо"
    """
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте входящее письмо.\nОБЯЗАТЕЛЬНО: ВР\nНе забудьте указать суть письма')
    await state.set_state(Form.enter_letter)

@dp.callback_query(lambda call: call.data in ['outer_letter','yes_outer'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    """
    Метод обработки исходящих сообщений - выводит сообщения с предложением для отправки письма
    и устанавливает состояние "отправить исходящее письмо"
    """
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте исходящее письмо.\nОБЯЗАТЕЛЬНО: ВР\nНе забудьте указать суть письма и Вр-ответа, если есть')
    await state.set_state(Form.outer_letter)

@dp.callback_query(lambda call: call.data in ['request_letter','yes_request'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    """
    Метод обработки запросов - выводит сообщения с предложением для отправки запроса
    и устанавливает состояние "отправить запрос"
    """
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте направленный запрос.\nНе забудьте указать RP и суть запроса')
    await state.set_state(Form.request_letter)
#endregion Menu functions

#region Send_letters
@dp.message(Form.enter_letter)
async def process_letter(message: Message, state: Form):
    """
    Метод обработки полученного входящего сообщения.
    Обрабатывает как ситуацию, где направляется письмо, а затем файлы
    так и ту, где отправляется только письмо с прикрепленным файлом
    """
    global COUNTMESS_USER #Нужен для модернизации с загрузкой файлов. Количество отправленных сообщений пользователем
    if message.caption is not None:
        await state.update_data(letter=gr.wrap_enterletter(message.caption)) #Вырываем из письма содержание, вр, сроки
        await message.answer("Идет загрузка ⏳")
        COUNTMESS_USER = 0

    elif message.text is not None:
        await state.update_data(letter=gr.wrap_enterletter(message.text))
        await message.answer("Идет загрузка ⏳")
        COUNTMESS_USER = 0

    if COUNTMESS_USER == 0:
        data = await state.get_data()
        if not data['letter']: #Если письмо не сохранилось
            await message.answer("Неправильный формат письма.")
        else:
            result = gs.post_enterletter(data['letter']) #Если письмо сохранилось, запускаем его в таблицу
            if not result:
                await message.answer("Это письмо уже есть в таблице.")
            else:
                await message.answer("Письмо сохранено!")

        await state.clear() #очистка состояния для отправки нового письма
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
    """
    Метод обработки полученного исходящего сообщения.
    Обрабатывает как ситуацию, где направляется письмо с вр и вр ответом,
    так и разделенное на два новых письма
    """
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
    if data['letter'][1][0]: #При наличии Вр
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
    """
    Метод обработки полученного запроса.
    Аналогичен методу отправки входящего письма.
    """
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
    """
    Метод обработки отказа от повторной отправки письма.
    Просто выводит сообщение о вызове меню
    """
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Нажмите /menu')


if __name__ == '__main__':
    dp.run_polling(bot)
