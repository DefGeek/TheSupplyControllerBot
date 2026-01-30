import os
from dotenv import load_dotenv

# Load .env file
load_dotenv()

# Redis configuration
REDIS_HOST = os.getenv("REDIS_HOST", "localhost")
REDIS_PORT = int(os.getenv("REDIS_PORT", 6379))
REDIS_DB = int(os.getenv("REDIS_DB", 0))

BOT_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_CREDENTIALS", "credentials.json")
GOOGLE_SCOPES = os.getenv("GOOGLE_SCOPES", "https://www.googleapis.com/auth/spreadsheets")

# Список администраторов бота (ID пользователей Telegram)
# Разделите ID запятыми, например: 123456789,987654321
ADMIN_IDS = [
    int(id.strip()) for id in os.getenv("ADMIN_IDS", "").split(",")
    if id.strip().isdigit()
]

# Добавьте себя как администратора по умолчанию
# (Замените 123456789 на ваш реальный Telegram ID)
if not ADMIN_IDS:
    ADMIN_IDS = [1340889852]  # Ваш Telegram ID

if not BOT_TOKEN:
    raise ValueError("❌ BOT_TOKEN is missing in environment variables (.env file)")

# Вспомогательная функция для проверки админа
def is_admin(user_id: int) -> bool:
    """Проверяет, является ли пользователь администратором бота"""
    return user_id in ADMIN_IDS
