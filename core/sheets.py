import os
import logging
import pickle
from typing import Optional
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ================== НАСТРОЙКИ ==================
SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Файл с OAuth 2.0 Desktop App credentials
OAUTH_CREDENTIALS_FILE = "/home/user/Desktop/TheSupplyControllerBot/bot/cred.json"

# Файл с токеном (создаётся автоматически)
TOKEN_FILE = "token.pickle"


# ================== ПРОВЕРКА ФАЙЛА ==================

def check_oauth_file():
    """Проверяет, что файл правильного формата"""
    import json

    if not os.path.exists(OAUTH_CREDENTIALS_FILE):
        return False, f"Файл {OAUTH_CREDENTIALS_FILE} не найден"

    try:
        with open(OAUTH_CREDENTIALS_FILE, 'r') as f:
            data = json.load(f)

        if "installed" in data:
            return True, "Desktop App OAuth 2.0 (правильно)"
        elif "web" in data:
            return False, "Web App OAuth 2.0 (нужен Desktop App)"
        elif "type" in data and data["type"] == "service_account":
            return False, "Service Account (нужен OAuth 2.0 Desktop App)"
        else:
            return False, "Неизвестный формат файла"

    except json.JSONDecodeError:
        return False, "Некорректный JSON файл"


# ================== OAUTH АУТЕНТИФИКАЦИЯ ==================

def get_credentials():
    """Получает OAuth 2.0 учетные данные для Desktop App"""

    # Проверяем файл
    is_valid, message = check_oauth_file()
    if not is_valid:
        raise ValueError(f"❌ {message}")

    creds = None

    # Проверяем существующий токен
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, "rb") as token:
                creds = pickle.load(token)
        except Exception:
            logging.warning("Не удалось загрузить токен, создаём новый")
            os.remove(TOKEN_FILE)
            creds = None

    # Если нет валидных учетных данных
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                logging.info("Токен обновлён")
            except Exception:
                logging.warning("Не удалось обновить токен, запрашиваем новый")
                creds = None

        if not creds:
            # ЗАПРАШИВАЕМ НОВЫЙ ТОКЕН
            logging.info("🔄 Запрашиваю OAuth авторизацию...")

            flow = InstalledAppFlow.from_client_secrets_file(
                OAUTH_CREDENTIALS_FILE,
                SCOPES
            )

            # Запускаем локальный сервер для авторизации
            creds = flow.run_local_server(
                port=0,
                authorization_prompt_message="📋 Авторизация для Telegram бота",
                success_message="✅ Авторизация успешна! Можно закрыть окно.",
                open_browser=True
            )

            logging.info("✅ OAuth авторизация успешна")

        # Сохраняем токен для будущих использований
        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(creds, token)
        logging.info(f"Токен сохранён в {TOKEN_FILE}")

    return creds


# ================== СОЗДАНИЕ ТАБЛИЦЫ ==================

def create_spreadsheet(title: str) -> str:
    """
    Создает новую Google таблицу через OAuth 2.0
    """
    try:
        # Получаем учетные данные
        creds = get_credentials()

        # Создаем сервис Google Drive
        drive_service = build("drive", "v3", credentials=creds)

        # Метаданные файла
        file_metadata = {
            "name": title,
            "mimeType": "application/vnd.google-apps.spreadsheet",
        }

        logging.info(f"🔄 Создаю таблицу: {title}")

        # Создаем файл таблицы
        file = drive_service.files().create(
            body=file_metadata,
            fields="id,webViewLink",
        ).execute()

        spreadsheet_id = file["id"]
        spreadsheet_url = file.get("webViewLink",
                                   f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}")

        logging.info(f"✅ Таблица создана!")
        logging.info(f"🆔 ID: {spreadsheet_id}")
        logging.info(f"🔗 Ссылка: {spreadsheet_url}")

        return spreadsheet_id

    except HttpError as e:
        error_msg = str(e)

        if "storageQuotaExceeded" in error_msg:
            raise Exception(
                "❌ Закончилось место в Google Drive!\n\n"
                "📊 Решение:\n"
                "1. Очистите корзину в Google Drive\n"
                "2. Удалите ненужные файлы\n"
                "3. Проверьте квоту: https://drive.google.com"
            )
        elif "invalid_grant" in error_msg:
            # Удаляем старый токен
            if os.path.exists(TOKEN_FILE):
                os.remove(TOKEN_FILE)
            raise Exception("Токен устарел. Удалён token.pickle, попробуйте снова.")
        else:
            raise Exception(f"❌ Ошибка Google API: {error_msg}")

    except Exception as e:
        logging.error(f"Ошибка создания таблицы: {e}")
        raise