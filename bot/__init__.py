import logging
from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.redis import RedisStorage
from redis.asyncio import Redis
from core.config import BOT_TOKEN, REDIS_HOST, REDIS_PORT, REDIS_DB#, REDIS_PASSWORD

# Logging
logging.basicConfig(level=logging.INFO)

# Initialize Redis storage
redis = Redis(
    host=REDIS_HOST,
    port=REDIS_PORT,
    db=REDIS_DB,
    #password=REDIS_PASSWORD,
    decode_responses=True
)

# Create Redis storage for FSM
storage = RedisStorage(redis=redis)

# Initialize bot
telegram_bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=storage)