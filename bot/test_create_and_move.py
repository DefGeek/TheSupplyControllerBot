import os
import logging
import pickle

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# ================== НАСТРОЙКИ ==================

logging.basicConfig(level=logging.INFO)

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

# OAuth-файл (скачанный из Google Cloud Console)
OAUTH_FILE = "credentials.json"

# Файл с токеном (создаётся автоматически)
TOKEN_FILE = "token.pickle"

# ID папки в ТВОЁМ Google Drive (или None)
FOLDER_ID = "1QWehy6-xg-lCAFl4TZqARwvUkzjFL7OA"
# если не нужно в папку:
# FOLDER_ID = None

TEST_TITLE = "ТЕСТОВАЯ ТАБЛИЦА - 2026"


# ================== OAUTH ==================

def get_credentials():
    creds = None

    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                OAUTH_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(creds, token)

    return creds


# ================== DRIVE ==================

def create_spreadsheet(title: str, folder_id: str | None = None) -> str:
    creds = get_credentials()
    drive_service = build("drive", "v3", credentials=creds)

    file_metadata = {
        "name": title,
        "mimeType": "application/vnd.google-apps.spreadsheet",
    }

    if folder_id:
        file_metadata["parents"] = [folder_id]
        logging.info(f"Создаём таблицу в папке: {folder_id}")
    else:
        logging.info("Создаём таблицу в корне Google Drive")

    file = drive_service.files().create(
        body=file_metadata,
        fields="id",
    ).execute()

    spreadsheet_id = file["id"]
    logging.info(f"Таблица создана! ID: {spreadsheet_id}")

    return spreadsheet_id


# ================== SHEETS ==================

def setup_spreadsheet(spreadsheet_id: str):
    creds = get_credentials()
    sheets_service = build("sheets", "v4", credentials=creds)

    requests = [
        # Переименовать лист + заморозить первую строку
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": 0,
                    "title": "Заявки",
                    "gridProperties": {"frozenRowCount": 1},
                },
                "fields": "title,gridProperties.frozenRowCount",
            }
        },
        # Цвет заголовков
        {
            "repeatCell": {
                "range": {
                    "sheetId": 0,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.2,
                            "green": 0.6,
                            "blue": 0.8,
                        },
                        "textFormat": {
                            "bold": True,
                            "foregroundColor": {
                                "red": 1.0,
                                "green": 1.0,
                                "blue": 1.0,
                            },
                        },
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        },
        # Заголовки
        {
            "updateCells": {
                "rows": [
                    {
                        "values": [
                            {"userEnteredValue": {"stringValue": "№"}},
                            {"userEnteredValue": {"stringValue": "Дата"}},
                            {"userEnteredValue": {"stringValue": "ФИО"}},
                            {"userEnteredValue": {"stringValue": "Телефон"}},
                            {"userEnteredValue": {"stringValue": "Объект"}},
                            {"userEnteredValue": {"stringValue": "Материал"}},
                            {"userEnteredValue": {"stringValue": "Количество"}},
                            {"userEnteredValue": {"stringValue": "Статус"}},
                        ]
                    }
                ],
                "start": {
                    "sheetId": 0,
                    "rowIndex": 0,
                    "columnIndex": 0,
                },
                "fields": "userEnteredValue",
            }
        },
    ]

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()

    logging.info("Таблица оформлена")


# ================== MAIN ==================

def main():
    try:
        sheet_id = create_spreadsheet(TEST_TITLE, FOLDER_ID)
        setup_spreadsheet(sheet_id)

        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"

        print("\n" + "=" * 60)
        print("УСПЕХ!")
        print(f"Таблица: {TEST_TITLE}")
        print(f"Ссылка: {url}")
        print("=" * 60 + "\n")

    except HttpError as e:
        print("\nОШИБКА Google API:")
        print(e)

    except Exception as e:
        print("\nОШИБКА:")
        print(e)


if __name__ == "__main__":
    main()
