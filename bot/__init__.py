import logging
from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage
from core.config import BOT_TOKEN

# Logging
logging.basicConfig(level=logging.INFO)

# FSM memory
storage = MemoryStorage()

# Initialize bot
telegram_bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=storage)
