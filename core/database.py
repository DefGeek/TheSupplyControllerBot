import sqlite3
import logging

DB_PATH = "users.db"


def init_db():
    conn = sqlite3.connect("users.db")
    cursor = conn.cursor()

    # Таблица пользователей
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            fio TEXT,
            position TEXT,
            phone TEXT
        )
    """)

    # Таблица разрешённых тем (с информацией об объекте и Google Sheet)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS allowed_topics (
            chat_id INTEGER,
            thread_id INTEGER,
            registered_by INTEGER,
            object_name TEXT,
            object_code TEXT,
            sheet_id TEXT,
            sheet_url TEXT,
            registered_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (chat_id, thread_id)
        )
    """)

    # Таблица разделов (категорий)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Таблица подразделов
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS subsections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            section_id INTEGER,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(name, section_id),
            FOREIGN KEY (section_id) REFERENCES sections (id)
        )
    """)

    # Таблица часто используемых единиц измерения
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS common_units (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Таблица заявок
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            chat_id INTEGER NOT NULL,
            thread_id INTEGER NOT NULL,
            section_id INTEGER,
            subsection_id INTEGER,
            delivery_date TEXT,
            status TEXT DEFAULT 'pending',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users (user_id),
            FOREIGN KEY (section_id) REFERENCES sections (id),
            FOREIGN KEY (subsection_id) REFERENCES subsections (id)
        )
    """)

    # Таблица позиций в заявке
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS request_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            request_id INTEGER NOT NULL,
            product_name TEXT NOT NULL,
            unit TEXT NOT NULL,
            quantity REAL NOT NULL,
            corrected_name TEXT,
            FOREIGN KEY (request_id) REFERENCES requests (id) ON DELETE CASCADE
        )
    """)

    # Вставляем базовые единицы измерения
    cursor.execute("""
        INSERT OR IGNORE INTO common_units (name, created_by) 
        VALUES ('шт', 0), ('кг', 0), ('л', 0), ('м', 0), ('м²', 0), ('м³', 0), ('уп', 0), ('пак', 0), ('кор', 0)
    """)

    conn.commit()
    conn.close()


def get_connection():
    """Return a new database connection"""
    return sqlite3.connect(DB_PATH)