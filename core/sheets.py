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
OAUTH_CREDENTIALS_FILE = "/app/bot/cred.json"

# Файл с токеном (создаётся автоматически)
TOKEN_FILE = "/app/token.pickle"


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
    import time
    import ssl
    import urllib3

    try:
        # Отключаем предупреждения о SSL (временно)
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        # Пробуем несколько раз с задержкой
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # Получаем учетные данные
                creds = get_credentials()

                # Создаем сервис Google Drive с увеличенным таймаутом
                drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)

                # Метаданные файла
                file_metadata = {
                    "name": title,
                    "mimeType": "application/vnd.google-apps.spreadsheet",
                }

                logging.info(f"🔄 Попытка {attempt + 1}/{max_retries}: Создаю таблицу: {title}")

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

                # Теперь работаем с Google Sheets API
                sheets_service = build("sheets", "v4", credentials=creds, cache_discovery=False)

                # 1. Переименовываем лист
                rename_request = [{
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": 0,
                            "title": "Заявки"
                        },
                        "fields": "title"
                    }
                }]

                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={"requests": rename_request}
                ).execute()

                logging.info("✅ Лист переименован в 'Заявки'")

                # 2. Добавляем заголовки
                headers = [
                    ["Дата создания", "ID заявки", "Раздел", "Подраздел", "Дата поставки",
                     "Наименование", "Единица", "Количество", "ID пользователя"]
                ]

                range_name = "Заявки!A1:I1"
                body = {"values": headers}

                sheets_service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption="RAW",
                    body=body
                ).execute()

                logging.info("✅ Заголовки добавлены")

                return spreadsheet_id

            except Exception as e:
                if attempt < max_retries - 1:
                    wait_time = 2 ** attempt  # Экспоненциальная backoff
                    logging.warning(f"Попытка {attempt + 1} не удалась: {e}. Жду {wait_time} секунд...")
                    time.sleep(wait_time)
                else:
                    raise

    except Exception as e:
        logging.error(f"Ошибка создания таблицы после {max_retries} попыток: {e}")

        # Даем более понятное сообщение об ошибке
        error_msg = str(e)
        if "SSL" in error_msg or "EOF" in error_msg:
            raise Exception(
                "❌ **Проблема с SSL соединением**\n\n"
                "**Возможные причины:**\n"
                "1. Проблемы с интернет-соединением\n"
                "2. Антивирус или фаервол блокируют соединение\n"
                "3. Проблемы с SSL сертификатами\n\n"
                "**Решения:**\n"
                "1. Проверьте интернет-соединение\n"
                "2. Отключите антивирус на время теста\n"
                "3. Попробуйте использовать VPN\n"
                "4. Обновите SSL сертификаты: `pip install --upgrade certifi`"
            )
        elif "quota" in error_msg.lower():
            raise Exception(
                "❌ **Превышена квота Google API**\n\n"
                "**Решение:**\n"
                "1. Подождите некоторое время\n"
                "2. Используйте другой аккаунт Google\n"
                "3. Обратитесь к администратору проекта"
            )
        else:
            raise Exception(f"❌ Ошибка создания таблицы: {error_msg}")

# ================== ДОБАВЛЕНИЕ ДАННЫХ В ТАБЛИЦУ ==================

def append_to_sheet(sheet_id: str, sheet_name: str, values: list):
    """
    Добавляет данные в существующий Google Spreadsheet

    Args:
        sheet_id: ID таблицы
        sheet_name: Название листа (например, "Заявки")
        values: Список списков с данными для добавления
    """
    try:
        # Получаем учетные данные
        creds = get_credentials()

        # Создаем сервис Google Sheets
        sheets_service = build("sheets", "v4", credentials=creds)

        # Определяем диапазон для добавления данных
        range_name = f"{sheet_name}!A:I"

        # Подготавливаем тело запроса
        body = {
            "values": values,
            "majorDimension": "ROWS"
        }

        logging.info(f"🔄 Добавляю {len(values)} строк в таблицу {sheet_id}")

        # Добавляем данные
        result = sheets_service.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range=range_name,
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body=body
        ).execute()

        logging.info(f"✅ Данные добавлены в таблицу!")
        logging.info(f"📊 Обновлено ячеек: {result.get('updatedCells', 0)}")

        return result

    except HttpError as e:
        error_msg = str(e)
        logging.error(f"Ошибка Google Sheets API: {error_msg}")

        if "invalid_grant" in error_msg:
            # Удаляем старый токен
            if os.path.exists(TOKEN_FILE):
                os.remove(TOKEN_FILE)
            raise Exception("Токен устарел. Удалён token.pickle, попробуйте снова.")
        else:
            raise Exception(f"❌ Ошибка добавления данных в таблицу: {error_msg}")

    except Exception as e:
        logging.error(f"Ошибка добавления данных: {e}")
        raise


# ================== СОЗДАНИЕ ТАБЛИЦЫ С ЗАГОЛОВКАМИ ==================

def create_spreadsheet_with_headers(title: str) -> str:
    """
    Создает новую Google таблицу с предустановленными заголовками
    """
    try:
        # Создаем таблицу
        spreadsheet_id = create_spreadsheet(title)

        # Добавляем заголовки
        headers = [
            ["Дата создания", "ID заявки", "Раздел", "Подраздел", "Дата поставки",
             "Наименование", "Единица", "Количество", "ID пользователя"]
        ]

        # Добавляем заголовки в таблицу
        append_to_sheet(spreadsheet_id, "Лист1", headers)

        # Переименовываем лист
        creds = get_credentials()
        sheets_service = build("sheets", "v4", credentials=creds)

        requests = [{
            "updateSheetProperties": {
                "properties": {
                    "sheetId": 0,
                    "title": "Заявки"
                },
                "fields": "title"
            }
        }]

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()

        logging.info(f"✅ Таблица '{title}' создана с заголовками")
        return spreadsheet_id

    except Exception as e:
        logging.error(f"Ошибка создания таблицы с заголовками: {e}")
        raise