import asyncio
import logging
from bot import telegram_bot, dp
from core.database import init_db
import bot.handlers  # register all handlers

logging.basicConfig(level=logging.INFO)


async def load_allowed_topics_cache():
    """Загружает кэш разрешённых тем из базы данных при запуске"""
    from bot.handlers import _allowed_topics_cache
    import sqlite3

    conn = sqlite3.connect("users.db")
    cursor = conn.cursor()
    cursor.execute("SELECT chat_id, thread_id FROM allowed_topics")
    rows = cursor.fetchall()
    conn.close()

    for chat_id, thread_id in rows:
        _allowed_topics_cache[(chat_id, thread_id)] = {"from_db": True}

    logging.info(f"Loaded {len(rows)} allowed topics from database")


async def main():
    init_db()
    await load_allowed_topics_cache()

    logging.info("🚀 Bot started and database initialized.")
    logging.info("📦 Using Redis for FSM storage")

    me = await telegram_bot.get_me()
    logging.info(f"Bot authorized: @{me.username}")
    logging.info(f"Redis connection: {dp.storage}")

    await dp.start_polling(telegram_bot)


if __name__ == "__main__":
    asyncio.run(main())