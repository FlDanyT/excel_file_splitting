from get_tables import main_tables, delete_folder
import asyncio
import logging
import sys
from os import getenv
import os

from aiogram import Bot, Dispatcher, html,Router, types
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.types import Message, FSInputFile, InputMediaDocument
from aiogram.fsm.storage.memory import MemoryStorage

TOKEN = "TOKEN"


dp = Dispatcher()
bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))


# @dp.message(CommandStart())
# async def command_start_handler(message: Message) -> None:
#     await message.answer(f"Hello, {html.bold(message.from_user.full_name)}!")


@dp.message()
async def message(message: Message) -> None:    
    document = message.document
    
    chatid = message.from_user.id
    
    file_id = document.file_id
    file_name = document.file_name

    file = await bot.get_file(file_id)
    save_directory = 'Entrance'  # Папка для сохранения
    file_path = os.path.join(save_directory, file_name)


    await bot.download_file(file.file_path, file_path) # скачиваем файл
    
    main_tables()
    
    
    path = 'Exit.zip'

    media_group = [InputMediaDocument(type='document',
                                      media=FSInputFile(path=path),
                                      )]

    await bot.send_media_group(chat_id=chatid,
                               media=media_group)
    delete_folder()

    
async def main() -> None:
    await dp.start_polling(bot)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    asyncio.run(main())