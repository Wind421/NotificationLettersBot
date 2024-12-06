import logging

from aiogram import Bot, Dispatcher
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import Message, InlineKeyboardButton, InlineKeyboardMarkup, CallbackQuery,ContentType

import excel_scripts as es
import google_requests as gr
import google_scripts as gs

logging.basicConfig(level=logging.INFO)

from config import TOKEN

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

class FormStatus(StatesGroup):
    """
    Класс FormStatus - отслеживает статус какого письма хочет исправить пользователь
    """
    waiting_for_vr = State()
    waiting_for_status = State()
    vr = ''
    status = ''
    what = ''

class FormKT(StatesGroup):
    waiting_for_kt = State()

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
        [InlineKeyboardButton(text='Отправить входящее письмо', callback_data="enter_letter")],
        [InlineKeyboardButton(text='Отправить исходящее письмо', callback_data="outer_letter")],
        [InlineKeyboardButton(text='Отправить RP запрос', callback_data="request_letter")],
        [InlineKeyboardButton(text='Изменить статус', callback_data="status")],
        [InlineKeyboardButton(text='отмена', callback_data="no")]
    ])
    global BOTMESS_ID
    last_message = await bot.send_message(chat_id=message.chat.id,
                           text="-----------------Главное меню-----------------",
                           parse_mode=ParseMode.MARKDOWN,
                           reply_markup=markup)
    BOTMESS_ID = last_message.message_id

@dp.message(Command('work'))
async def send_menu(message: Message):
    await bot.send_message(message.chat.id, 'Привет, Я работаю')

@dp.message(Command('send_kt'))
async def send_menu(message: Message, state: FormKT):
    await bot.send_message(message.chat.id,
                           'Перешлите или отправьте экспорт контрольных точек из ГАСУ. \nНЕ ИЗМЕНЯЙТЕ ФАЙЛ ИЛИ ЕГО НАЗВАНИЕ.')
    await state.set_state(FormKT.waiting_for_kt)

@dp.message(FormKT.waiting_for_kt)
async def process_letter(message: Message, state: FormKT):
    if message.content_type == ContentType.DOCUMENT:
        try:
            file = await bot.get_file(message.document.file_id)
            file_path = '.\excel\\'+str(message.document.file_name)
            await bot.download_file(file.file_path, file_path)
            await bot.send_message(message.chat.id, 'Файл успешно сохранен!')
            await bot.send_message(message.chat.id, 'Нажмите /menu')
        except Exception:
            await bot.send_message(message.chat.id, 'Что-то пошло не так.')
            print(Exception.args)

#region Menu functions
@dp.callback_query(lambda call: call.data in ['enter_letter','yes_enter'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    """
    Метод обработки входящих сообщений - выводит сообщения с предложением для отправки письма
    и устанавливает состояние "отправить входящее письмо"
    """
    global BOTMESS_ID
    BOTMESS_ID=None

    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте входящее письмо.\nОБЯЗАТЕЛЬНО (формат вр-0000000): ВР\nНе забудьте указать суть письма и срок\n(формат срок 00.00.0000 или срок срочно или срок сегодня)')
    await state.set_state(Form.enter_letter)

@dp.callback_query(lambda call: call.data in ['outer_letter','yes_outer'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    """
    Метод обработки исходящих сообщений - выводит сообщения с предложением для отправки письма
    и устанавливает состояние "отправить исходящее письмо"
    """
    global BOTMESS_ID
    BOTMESS_ID = None
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.message.chat.id,
                           'Перешлите или отправьте исходящее письмо.\nОБЯЗАТЕЛЬНО (формат вр-0000000): ВР\nНе забудьте указать суть письма и Вр-ответа\n(формат Ответ на вр-00000000)')
    await state.set_state(Form.outer_letter)

@dp.callback_query(lambda call: call.data in ['request_letter','yes_request'])
async def callback_menu(callback_query: CallbackQuery, state: Form):
    """
    Метод обработки запросов - выводит сообщения с предложением для отправки запроса
    и устанавливает состояние "отправить запрос"
    """
    global BOTMESS_ID
    BOTMESS_ID = None

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
    if data['letter'][1][0]: #При наличии Вр, либо строка-либо False
        await message.answer("Идет загрузка ⏳")

        if not data['letter']:
            await message.answer("Неправильный формат письма.")
        else:
            result = gs.post_outerletter(data['letter'])
            if not result:
                await message.answer("Это письмо уже есть в таблице.")
            else:
                gs.post_ansvr(data['letter'])
                await message.answer("Письмо сохранено!")

        markup = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text='Да', callback_data="yes_outer"),
                InlineKeyboardButton(text='Нет', callback_data="no")],
        ])
        await state.clear()
        await bot.send_message(chat_id=message.chat.id,
                               text="Отправить ещё одно исходящее письмо?",
                               parse_mode=ParseMode.MARKDOWN,
                               reply_markup=markup)


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

#region Change
@dp.callback_query(lambda call: call.data in ['status'])
async def callback_menu(callback_query: CallbackQuery):
    """
    Метод изменения статуса - выводит сообщения с предложением для изменения статусов
    """
    await bot.answer_callback_query(callback_query.id)

    markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text='Входящего письма', callback_data="change_enter")],
        [InlineKeyboardButton(text='Исходящего письма', callback_data="change_outer")],
        [InlineKeyboardButton(text='Запроса', callback_data="change_request")],
        [InlineKeyboardButton(text='отмена', callback_data="no")]
    ])
    global BOTMESS_ID
    if BOTMESS_ID is not None:
        await bot.edit_message_text(text='Какой статус вы хотите обновить?',
                                    chat_id=callback_query.message.chat.id,
                                    message_id=BOTMESS_ID,
                                    reply_markup=markup)
        BOTMESS_ID = None
    else:
        await bot.send_message(text='Какой статус вы хотите обновить?',
                                    chat_id=callback_query.message.chat.id,
                                    reply_markup=markup)

@dp.callback_query(lambda call: call.data in ['change_enter',"change_outer","change_request"])
async def callback_change(callback_query: CallbackQuery, state: FormStatus):
    await state.update_data(what=callback_query.data.split('_')[1])

    if callback_query.data.split('_')[1] in ['enter',"outer"]:
        await bot.send_message(callback_query.message.chat.id,'Введите Вр:\n(формат вр-00000000)')
    else:
        await bot.send_message(callback_query.message.chat.id, 'Введите RP:\n(RP00000)')
    await state.set_state(FormStatus.waiting_for_vr)

@dp.message(FormStatus.waiting_for_vr)
async def waiting_vr(message: Message, state: FormStatus):
    if message.text is not None:
        try:
            vr_value = gr.wrap_change(message.text)
            await state.update_data(vr=vr_value)
            data = await state.get_data()
            print(data['what'])
            if data['what'] == 'enter':
                await bot.send_message(message.chat.id,'Введите статус:\n(отработано или согласовано)')
            elif data['what'] == 'outer':
                await bot.send_message(message.chat.id,'Введите статус:\n(на согласовании у ... \nили не согласовано у ...\nили подписано 00.00.0000 №Исх-...)')
            elif data['what'] == 'request':
                await bot.send_message(message.chat.id,'Введите статус:\n(выполнено или не выполнено)')
            await state.set_state(FormStatus.waiting_for_status)

        except KeyError:
            await state.clear()
            await bot.send_message(message.chat.id, 'Неправильный формат vr.\nПопробуйте снова или нажмите /menu',
                                   reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                                       [InlineKeyboardButton(text='Изменить статус', callback_data="status")]]))

@dp.message(FormStatus.waiting_for_status)
async def waiting_status(message: Message, state: FormStatus):
    if message.text is not None:
        await state.update_data(status=message.text)
        print(await state.get_data())
        try:
            await message.answer("Идет загрузка ⏳")
            gs.change(await state.get_data())
            await bot.send_message(message.chat.id, 'Изменение произошло успешно!')

            markup = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text='Да', callback_data="status"),
                 InlineKeyboardButton(text='Нет', callback_data="no")],
            ])
            await state.clear()
            await bot.send_message(chat_id=message.chat.id,
                                   text="Изменить что-то ещё?",
                                   parse_mode=ParseMode.MARKDOWN,
                                   reply_markup=markup)
        except ValueError:
            await state.clear()
            await bot.send_message(message.chat.id, 'Неправильный формат статуса.\nПопробуйте снова или нажмите /menu',
                                   reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                                       [InlineKeyboardButton(text='Изменить статус', callback_data="status")]]))
        except KeyError:
            await state.clear()
            await bot.send_message(message.chat.id, 'Письма или запроса с таким номером нет.\nПопробуйте снова или нажмите /menu',
                                   reply_markup = InlineKeyboardMarkup(inline_keyboard=[
                                       [InlineKeyboardButton(text='Изменить статус', callback_data="status")]
                                   ]))

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
