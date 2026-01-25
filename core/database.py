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
    
    conn.commit()
    conn.close()

def get_connection():
    """Return a new database connection"""
    return sqlite3.connect(DB_PATH)
