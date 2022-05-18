import time
import openpyxl
from aiogram import Bot, Dispatcher, executor, types
from config import tg_token, file_link

bot = Bot(token=tg_token)
dp = Dispatcher(bot)


@dp.message_handler(commands=['start'])
async def start_command(message: types.Message):
    await message.answer("Приветствую!")
    await message.answer("Давай начнем поиск отделений Сбербанка, где можно приобрести валюту в кассе")
    await message.answer("Введи название города на русском языке")


@dp.message_handler()
async def get_data(message: types.Message):
    await message.answer("Пожалуйста, подожди. Идет сбор данных... \U0001F609")
    time.sleep(1)

    try:
        book = openpyxl.open("exchange.xlsx", read_only=True)
        sheet = book.active
        address = sheet['F3:F999']

        for row in address:
            for cell in row:
                address_list = cell.value
                if f'г.{message.text}'.lower().replace('ё', 'е') in address_list.lower().replace('ё', 'е'):
                    await message.answer(address_list)

    except:
        await message.answer('\U0001F635 Упс... что-то пошло не так. Попробуй снова')

    time.sleep(1)
    await message.answer("Можно приступать к поиску ближайшего к вам адреса. Удачи! \U0001F601")

if __name__ == '__main__':
    executor.start_polling(dp)
