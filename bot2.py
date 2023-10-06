import io
import aiogram
import logging
import asyncio
from aiogram import Bot, Dispatcher, F,Router, types, html
from openpyxl import load_workbook
from aiogram.filters import Command
from aiogram.enums import ParseMode
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from typing import Any, Dict


from aiogram.types import (
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
)

from config import TOKEN
from utils import Form
from messages import MESSAGES


bot = Bot(token=TOKEN)
dp = Dispatcher()

#dp.middleware.setup(LoggingMiddleware())
my_router = Router(name= __name__)

#@form_router.message(CommandStart())
@my_router.message(Command("start"))
async def process_start_command(message: types.Message):
    await message.reply(MESSAGES['start'])

@my_router.message(Command("help"))
async def process_help_command(message: types.Message):
    await message.reply(MESSAGES['help'])


@my_router.message(Command("groupp"))
async def process_groupps(message: Message, state: FSMContext) -> None:
    await state.set_state(Form.choose_groupps)
    await message.answer(
        "Какую группу хотите проанализировать",
        reply_markup=ReplyKeyboardRemove(),
    )

@my_router.message(Form.choose_groupps)
async def process_choose_groupps(message: Message, state: FSMContext) -> None:
    await state.update_data(process_groupps=message.text)
    await state.set_state(Form.start_analyse)
    await message.answer(
        f"Вы выбрали {html.quote(message.text)}. Все правильно?",
        reply_markup=ReplyKeyboardMarkup(
            keyboard=[
                [
                    KeyboardButton(text="Yes"),
                    KeyboardButton(text="No"),
                ]
            ],
            resize_keyboard=True,
        ),
    )
 

@my_router.message(Form.start_analyse, F.text.casefold() == "no")
async def process_dont_like_write_bots(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    await state.clear()
    await message.answer(
        "Ничего страшного.",
        reply_markup=ReplyKeyboardRemove(),
    )

@my_router.message(Form.start_analyse, F.text.casefold() == "yes")
async def message_handler(message: Message) -> Any: 
    await message.answer(
        "Классно! Начинаю анализ...",
    reply_markup=ReplyKeyboardRemove(),
    )


async def handle_excel_file(message: types.Message):
    # Проверяем, является ли присланный файл Excel
    if message.document.mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        try:
            # Загружаем файл Excel
            file_id = message.document.file_id
            file = await bot.download_file_by_id(file_id)
            file_path = f"downloads/{file_id}.xlsx"  # Укажите путь для сохранения файла
            with open(file_path, 'wb') as f:
                f.write(file.read())
        
            # Загружаем файл с помощью openpyxl
            wb = load_workbook(filename=io.BytesIO(await file.download()))

            # Код для работы внутри файла
            groups_column = wb["Группа"]
            groups_name=input("Номер группы = ")
            all_estimate = wb["Оценка"]
            group_estimate = wb[groups_column == groups_name].shape[0]
            student_number = wb[groups_column == groups_name]["Личный номер студента"]
            unique_student_number = student_number.unique()
            unique_control_level = wb["Уровень контроля"].unique()
            unique_years = sorted(wb["Год"].unique())

            await message.reply('Файл обработан успешно')

            bot.send_message(message.chat.id, (f"В исходном датасете содержалось {all_estimate.size} оценок, из них {group_estimate} оценок относятся к группе ПИ101 \n"
                                                f"В датасете находится ", unique_student_number.size, "студентов со следующими личными номерами: {unique_student_number} \n"
                                                f"Исполmзуемые формы контроля: {unique_control_level}\n"
                                                f"Данные представлены по следующим учебным годам: ", {sorted(unique_years)}))
        except Exception as e:
            bot.send_message(message.chat.id, f'Произошла ошибка: {str(e)}')
    else:
        await message.answer("Пожалуйста, отправьте файл Excel в формате .xlsx.")

#dp.register_message_handler(process_excel_file, content_types=types.ContentType.DOCUMENT)

async def main():
    bot = Bot(token=TOKEN, parse_mode=ParseMode.HTML)
    dp = Dispatcher()
    dp.include_router(my_router)

    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())