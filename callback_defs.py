from aiogram import Bot, Dispatcher
from aiogram.types import Message,InlineKeyboardButton,InlineKeyboardMarkup,CallbackQuery
from aiogram.enums import ParseMode
from aiogram.fsm.state import State, StatesGroup

import logging
logging.basicConfig(level=logging.INFO)

from config import TOKEN

bot = Bot(token=TOKEN)

class Form(StatesGroup):
    enter_letter = State()
    outer_letter = State()
    request_letter = State()

#region Menu functions
@dp.callback_query(lambda call: call.data in ['outer_letter','yes_other'])
async def callback_menu_r(callback_query: CallbackQuery, state: Form):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте исходящее письмо.\nНе забудьте указать ВР и краткое содержание письма,\n а также Вр-ответа если он присутствует')
    await state.set_state(Form.outer_letter)

@dp.callback_query(lambda call: call.data in ['enter_letter','yes_enter'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте исходящее письмо.\nНе забудьте указать ВР и суть запроса')
    await state.set_state(Form.enter_letter)

@dp.callback_query(lambda call: call.data in ['request_letter','yes_request'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте направленный запрос.\nНе забудьте указать RP и суть запроса')
    await state.set_state(Form.enter_letter)

@dp.message(Form.outer_letter)
async def process_letter(message: Message, state: Form):
    await state.update_data(outer_letter=message.text)

    #answer = gs.wrap_enter_letter(message.text)
    try:
        #gs.wrap_enter_letter(answer)
        #await bot.send_message(message.chat.id, f'Письмо успешно загружено в таблицу!{answer[0],answer[1],answer[2],answer[3]}')
        await bot.send_message(message.chat.id, f'Письмо успешно загружено в таблицу!')
    except Exception:
        await bot.send_message(message.chat.id, 'Что-то пошло не так...')
    await state.clear()

    markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text='Да', callback_data="yes_other"),
         InlineKeyboardButton(text='Нет', callback_data="no")],
    ])
    await bot.send_message(chat_id=message.chat.id,
                           text="Отправить ещё письмо?",
                           parse_mode=ParseMode.MARKDOWN,
                           reply_markup=markup)

@dp.message(Form.enter_letter)
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

@dp.message(Form.enter_letter)
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

@dp.callback_query(lambda call: call.data == 'no')
async def process_no(callback_query: CallbackQuery):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Нажмите /menu')
