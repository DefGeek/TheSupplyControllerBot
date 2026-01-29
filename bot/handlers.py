from aiogram import F
import logging
import pickle
from typing import Dict, Tuple, List
from datetime import datetime, timedelta
from aiogram import types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.enums import ChatMemberStatus
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.filters.callback_data import CallbackData
import asyncio
from bot import telegram_bot, dp
from core.database import get_connection
from core.sheets import create_spreadsheet, append_to_sheet
from bot.ai_spellcheck import check_spelling, check_list_spelling
from core.config import is_admin

# Кэш разрешённых тем: {(chat_id, thread_id): {"admin_id": int, "registered_at": datetime}}
_allowed_topics_cache: Dict[tuple[int, int], dict] = {}

# ================== CALLBACK DATA CLASSES ==================
class SectionCallback(CallbackData, prefix="section"):
    action: str  # select, create, cancel, select_for_subsection
    section_id: int = 0

class SubsectionCallback(CallbackData, prefix="subsection"):
    action: str  # select, create, cancel
    subsection_id: int = 0
    section_id: int = 0

class UnitCallback(CallbackData, prefix="unit"):
    action: str  # select, create
    name: str = ""

class DateCallback(CallbackData, prefix="date"):
    action: str  # select, quick
    date: str = ""

class RequestCallback(CallbackData, prefix="request"):
    action: str  # add_more, finish, cancel, confirm_correction
    item_index: int = 0

# ================== STATES ==================
class RegistrationStates(StatesGroup):
    waiting_for_fio = State()
    waiting_for_position = State()
    waiting_for_phone = State()

class BotRegistrationStates(StatesGroup):
    waiting_for_object_code = State()
    waiting_for_object_name = State()

class CreateSectionStates(StatesGroup):
    waiting_for_section_name = State()

class CreateSubsectionStates(StatesGroup):
    waiting_for_subsection_name = State()

class CreateRequestStates(StatesGroup):
    waiting_for_section = State()
    waiting_for_subsection = State()
    waiting_for_delivery_date = State()
    waiting_for_product_name = State()
    waiting_for_unit = State()
    waiting_for_quantity = State()
    waiting_for_correction = State()
    waiting_for_custom_date = State()
    waiting_for_custom_unit = State()
    waiting_for_custom_subsection = State()

# ================== HELPER FUNCTIONS ==================
async def is_user_admin(chat_id: int, user_id: int, bot) -> bool:
    """Проверяет, является ли пользователь администратором в чате"""
    try:
        member = await bot.get_chat_member(chat_id, user_id)
        return member.status in [
            ChatMemberStatus.ADMINISTRATOR,
            ChatMemberStatus.CREATOR
        ]
    except Exception as e:
        logging.error(f"Error checking admin status: {e}")
        return False

async def is_allowed_context(message: types.Message) -> bool:
    """Проверяет, разрешён ли боту работать в данном контексте"""
    chat = message.chat
    chat_type = chat.type

    # Личные сообщения всегда разрешены
    if chat_type == "private":
        return True

    # В группах/супергруппах проверяем тему
    if chat_type in ("group", "supergroup"):
        thread_id = getattr(message, "message_thread_id", None)
        # Сообщения вне тем игнорируем (кроме обсуждений с thread_id=0)
        if thread_id is None or thread_id == 0:
            return False

        # Проверяем, зарегистрирована ли тема
        cache_key = (chat.id, thread_id)
        if cache_key in _allowed_topics_cache:
            return True

        # Если нет в кэше, проверяем базу данных
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT 1 FROM allowed_topics WHERE chat_id = ? AND thread_id = ?",
            (chat.id, thread_id)
        )
        result = cursor.fetchone()
        conn.close()

        if result:
            _allowed_topics_cache[cache_key] = {"from_db": True}
            return True

        # Если тема не зарегистрирована, отправляем сообщение в зависимости от роли пользователя
        is_admin_in_group = await is_user_admin(chat.id, message.from_user.id, message.bot)
        if is_admin_in_group:
            await message.answer(
                "🤖 **Бот не активирован в этой теме**\n\n"
                "Для активации бота в этой теме выполните команду:\n"
                "`/register_bot`\n\n"
                "*Только администраторы могут активировать бота*"
            )
        else:
            await message.answer(
                "🤖 **Бот не активирован в этой теме**\n\n"
                "Обратитесь к администратору группы для активации бота."
            )
        return False

    return False

def get_menu_inline_keyboard() -> types.InlineKeyboardMarkup:
    """Создаёт меню с инлайн-кнопками"""
    keyboard = types.InlineKeyboardMarkup(
        inline_keyboard=[
            [types.InlineKeyboardButton(text="📝 Регистрация", callback_data="menu_registration")],
            [types.InlineKeyboardButton(text="📋 Создать заявку", callback_data="create_request_start")],
            [types.InlineKeyboardButton(text="📂 Управление разделами", callback_data="manage_sections")],
            [types.InlineKeyboardButton(text="❓ Справка", callback_data="menu_help")]
        ]
    )
    return keyboard

def get_cancel_keyboard() -> types.InlineKeyboardMarkup:
    """Клавиатура для отмены действия"""
    keyboard = types.InlineKeyboardMarkup(
        inline_keyboard=[
            [types.InlineKeyboardButton(text="❌ Отмена", callback_data="cancel_action")]
        ]
    )
    return keyboard

# ================== MENU HANDLERS ==================
@dp.message(Command("menu"))
async def menu_command(message: types.Message):
    """Показывает меню с кнопками"""
    if not await is_allowed_context(message):
        return
    await message.answer(
        "📱 **Главное меню**\n\n"
        "Выберите действие:",
        reply_markup=get_menu_inline_keyboard()
    )

@dp.callback_query(lambda c: c.data == "menu_registration")
async def process_menu_registration(callback_query: types.CallbackQuery, state: FSMContext):
    """Обработчик нажатия на кнопку 'Регистрация' в меню"""
    await callback_query.answer()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE user_id = ?", (callback_query.from_user.id,))
    user = cursor.fetchone()
    conn.close()
    if user:
        await callback_query.message.answer(
            "✅ Вы уже зарегистрированы в системе.\n\n"
            "Ваши данные:\n"
            f"👤 ФИО: {user[1]}\n"
            f"💼 Должность: {user[2] if user[2] != '-' else 'Не указана'}\n"
            f"📞 Телефон: {user[3]}"
        )
    else:
        await callback_query.message.answer("👋 Начнём регистрацию! Введите ваши ФИО:")
        await state.set_state(RegistrationStates.waiting_for_fio)

@dp.callback_query(lambda c: c.data == "create_request_start")
async def process_create_request_start(callback_query: types.CallbackQuery, state: FSMContext):
    """Начинает процесс создания заявки"""
    await callback_query.answer()
    # Проверяем, зарегистрирован ли пользователь
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE user_id = ?", (callback_query.from_user.id,))
    user = cursor.fetchone()
    conn.close()
    if not user:
        await callback_query.message.answer(
            "❌ Сначала зарегистрируйтесь с помощью команды /start или через меню",
            reply_markup=get_menu_inline_keyboard()
        )
        return

    # Получаем список разделов
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM sections ORDER BY name")
    sections = cursor.fetchall()
    conn.close()
    if not sections:
        await callback_query.message.answer(
            "❌ Разделы не созданы. Сначала создайте разделы через меню 'Управление разделами'.",
            reply_markup=get_menu_inline_keyboard()
        )
        return

    # Создаем клавиатуру с разделами
    builder = InlineKeyboardBuilder()
    for section_id, section_name in sections:
        builder.button(
            text=section_name,
            callback_data=SectionCallback(action="select", section_id=section_id)
        )
    builder.button(
        text="➕ Создать новый раздел",
        callback_data=SectionCallback(action="create")
    )
    builder.button(
        text="❌ Отмена",
        callback_data="cancel_action"
    )
    builder.adjust(1)
    await callback_query.message.answer(
        "📂 **Выберите раздел:**\n\n"
        "Или создайте новый, если нужного нет в списке:",
        reply_markup=builder.as_markup()
    )
    # Инициализируем список позиций в состоянии
    await state.update_data(items=[])
    await state.set_state(CreateRequestStates.waiting_for_section)

@dp.callback_query(lambda c: c.data == "manage_sections")
async def process_manage_sections(callback_query: types.CallbackQuery):
    """Меню управления разделами"""
    await callback_query.answer()
    builder = InlineKeyboardBuilder()
    builder.button(
        text="➕ Создать раздел",
        callback_data="create_section_menu"
    )
    builder.button(
        text="➕ Создать подраздел",
        callback_data="create_subsection_menu"
    )
    builder.button(
        text="📋 Список разделов",
        callback_data="list_sections"
    )
    builder.button(
        text="🔙 Назад",
        callback_data="back_to_menu"
    )
    builder.adjust(1)
    await callback_query.message.answer(
        "📂 **Управление разделами**\n\n"
        "Выберите действие:",
        reply_markup=builder.as_markup()
    )

@dp.callback_query(lambda c: c.data == "menu_help")
async def process_menu_help(callback_query: types.CallbackQuery):
    """Обработчик нажатия на кнопку 'Справка' в меню"""
    await callback_query.answer()
    help_text = (
        "📋 **Доступные команды:**\n\n"
        "**Основные:**\n"
        "• `/start` — регистрация в системе\n"
        "• `/menu` — главное меню с кнопками\n"
        "• `/create_request` — создать новую заявку\n"
        "• `/create_section` — создать новый раздел\n"
        "• `/topic_info` — информация о текущей теме\n"
        "• `/help` — эта справка\n"
        "• `/cancel` — отменить текущий процесс\n\n"
        "**Управление ботом в теме:**\n"
        "• `/register_bot` — активировать бота в этой теме (админы)\n"
        "• `/unregister_bot` — деактивировать бота в этой теме (админы)\n\n"
        "**Процесс создания заявки:**\n"
        "1. Выберите раздел (или создайте новый)\n"
        "2. Выберите подраздел (или создайте новый)\n"
        "3. Выберите дату поставки\n"
        "4. Вводите позиции: название, единица измерения, количество\n"
        "5. Бот проверяет орфографию через AI\n"
        "6. Завершите создание заявки\n\n"
        "**Использование меню:**\n"
        "Нажмите /menu для вызова меню с кнопками"
    )
    await callback_query.message.answer(help_text)

# ================== SECTION MANAGEMENT ==================
@dp.callback_query(lambda c: c.data == "create_section_menu")
async def create_section_menu(callback_query: types.CallbackQuery, state: FSMContext):
    """Начинает создание раздела"""
    await callback_query.answer()
    await callback_query.message.answer("📝 Введите название нового раздела:")
    await state.set_state(CreateSectionStates.waiting_for_section_name)

@dp.message(CreateSectionStates.waiting_for_section_name)
async def process_new_section_name(message: types.Message, state: FSMContext):
    """Обрабатывает ввод названия раздела"""
    section_name = message.text.strip()
    if not section_name:
        await message.answer("❌ Название раздела не может быть пустым. Введите снова:")
        return

    # Проверяем, существует ли уже такой раздел
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM sections WHERE name = ?", (section_name,))
    existing = cursor.fetchone()
    if existing:
        await message.answer(f"❌ Раздел '{section_name}' уже существует.")
        conn.close()
        await state.clear()
        return

    # Создаем раздел
    cursor.execute(
        "INSERT INTO sections (name, created_by) VALUES (?, ?)",
        (section_name, message.from_user.id)
    )
    conn.commit()
    section_id = cursor.lastrowid
    conn.close()
    await message.answer(f"✅ Раздел '{section_name}' успешно создан (ID: {section_id})!")
    await state.clear()

@dp.callback_query(lambda c: c.data == "create_subsection_menu")
async def create_subsection_menu(callback_query: types.CallbackQuery):
    """Начинает создание подраздела - сначала выбираем раздел"""
    await callback_query.answer()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM sections ORDER BY name")
    sections = cursor.fetchall()
    conn.close()
    if not sections:
        await callback_query.message.answer(
            "❌ Сначала создайте разделы. Нет доступных разделов."
        )
        return

    builder = InlineKeyboardBuilder()
    for section_id, section_name in sections:
        builder.button(
            text=section_name,
            callback_data=SectionCallback(action="select_for_subsection", section_id=section_id)
        )
    builder.button(
        text="❌ Отмена",
        callback_data="cancel_action"
    )
    builder.adjust(1)
    await callback_query.message.answer(
        "📂 **Выберите раздел для подраздела:**",
        reply_markup=builder.as_markup()
    )

@dp.callback_query(SectionCallback.filter(F.action == "select_for_subsection"))
async def select_section_for_subsection(
    callback_query: types.CallbackQuery,
    callback_data: SectionCallback,
    state: FSMContext
):
    """Обрабатывает выбор раздела для создания подраздела"""
    await callback_query.answer()
    # Сохраняем выбранный раздел
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sections WHERE id = ?", (callback_data.section_id,))
    section = cursor.fetchone()
    conn.close()
    if not section:
        await callback_query.message.answer("❌ Раздел не найден.")
        return

    await state.update_data(
        subsection_section_id=callback_data.section_id,
        subsection_section_name=section[0]
    )
    await callback_query.message.answer(f"📝 Введите название нового подраздела для раздела '{section[0]}':")
    await state.set_state(CreateSubsectionStates.waiting_for_subsection_name)

@dp.message(CreateSubsectionStates.waiting_for_subsection_name)
async def process_new_subsection_name(message: types.Message, state: FSMContext):
    """Обрабатывает ввод названия подраздела"""
    subsection_name = message.text.strip()
    if not subsection_name:
        await message.answer("❌ Название подраздела не может быть пустым. Введите снова:")
        return

    data = await state.get_data()
    section_id = data.get('subsection_section_id')
    section_name = data.get('subsection_section_name')

    # Проверяем, существует ли уже такой подраздел в этом разделе
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id FROM subsections WHERE name = ? AND section_id = ?",
        (subsection_name, section_id)
    )
    existing = cursor.fetchone()
    if existing:
        await message.answer(f"❌ Подраздел '{subsection_name}' уже существует в разделе '{section_name}'.")
        conn.close()
        await state.clear()
        return

    # Создаем подраздел
    cursor.execute(
        "INSERT INTO subsections (name, section_id, created_by) VALUES (?, ?, ?)",
        (subsection_name, section_id, message.from_user.id)
    )
    conn.commit()
    subsection_id = cursor.lastrowid
    conn.close()
    await message.answer(
        f"✅ Подраздел '{subsection_name}' успешно создан в разделе '{section_name}' (ID: {subsection_id})!"
    )
    await state.clear()

@dp.callback_query(lambda c: c.data == "list_sections")
async def list_sections(callback_query: types.CallbackQuery):
    """Показывает список всех разделов и подразделов"""
    await callback_query.answer()
    conn = get_connection()
    cursor = conn.cursor()
    # Получаем все разделы с подразделами
    cursor.execute("""
        SELECT s.id, s.name, s.created_at, GROUP_CONCAT(sb.name, ', ') as subsections
        FROM sections s
        LEFT JOIN subsections sb ON s.id = sb.section_id
        GROUP BY s.id
        ORDER BY s.name
    """)
    sections = cursor.fetchall()
    conn.close()
    if not sections:
        await callback_query.message.answer("📭 Разделы не созданы.")
        return

    sections_text = "📂 **Список разделов:**\n\n"
    for section_id, section_name, created_at, subsections in sections:
        sections_text += f"**{section_name}** (ID: {section_id})\n"
        sections_text += f"Создан: {created_at[:10]}\n"
        if subsections:
            sections_text += f"Подразделы: {subsections}\n"
        else:
            sections_text += "Подразделы: нет\n"
        sections_text += "\n"

    # Разделяем на части если текст слишком длинный
    if len(sections_text) > 4000:
        parts = [sections_text[i:i + 4000] for i in range(0, len(sections_text), 4000)]
        for part in parts:
            await callback_query.message.answer(part)
    else:
        await callback_query.message.answer(sections_text)

# ================== REQUEST CREATION ==================
@dp.callback_query(SectionCallback.filter(F.action == "select"))
async def select_section_for_request(
    callback_query: types.CallbackQuery,
    callback_data: SectionCallback,
    state: FSMContext
):
    """Обрабатывает выбор раздела для заявки"""
    await callback_query.answer()
    # Сохраняем выбранный раздел
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sections WHERE id = ?", (callback_data.section_id,))
    section = cursor.fetchone()
    if not section:
        await callback_query.message.answer("❌ Раздел не найден.")
        conn.close()
        return
    section_name = section[0]
    await state.update_data(section_id=callback_data.section_id, section_name=section_name)

    # Получаем подразделы для этого раздела
    cursor.execute(
        "SELECT id, name FROM subsections WHERE section_id = ? ORDER BY name",
        (callback_data.section_id,)
    )
    subsections = cursor.fetchall()
    conn.close()

    # Создаем клавиатуру с подразделами
    builder = InlineKeyboardBuilder()
    if subsections:
        for subsection_id, subsection_name in subsections:
            builder.button(
                text=subsection_name,
                callback_data=SubsectionCallback(
                    action="select",
                    subsection_id=subsection_id,
                    section_id=callback_data.section_id
                )
            )
    builder.button(
        text="➕ Создать новый подраздел",
        callback_data=SubsectionCallback(action="create", section_id=callback_data.section_id)
    )
    builder.button(
        text="❌ Отмена",
        callback_data="cancel_action"
    )
    builder.adjust(1)
    await callback_query.message.answer(
        f"📂 **Раздел:** {section_name}\n\n"
        "📁 **Выберите подраздел:**\n"
        "Или создайте новый, если нужного нет в списке:",
        reply_markup=builder.as_markup()
    )
    await state.set_state(CreateRequestStates.waiting_for_subsection)

@dp.callback_query(SectionCallback.filter(F.action == "create"))
async def create_section_during_request(
    callback_query: types.CallbackQuery,
    state: FSMContext
):
    """Создание раздела в процессе создания заявки"""
    await callback_query.answer()
    await callback_query.message.answer(
        "📝 Введите название нового раздела:"
    )
    # Устанавливаем состояние для создания раздела
    await state.set_state(CreateSectionStates.waiting_for_section_name)
    await state.update_data(creating_section_in_request=True)

@dp.callback_query(SubsectionCallback.filter(F.action == "select"))
async def select_subsection_for_request(
    callback_query: types.CallbackQuery,
    callback_data: SubsectionCallback,
    state: FSMContext
):
    """Обрабатывает выбор подраздела для заявки"""
    await callback_query.answer()
    # Сохраняем выбранный подраздел
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT name FROM subsections WHERE id = ? AND section_id = ?",
        (callback_data.subsection_id, callback_data.section_id)
    )
    subsection = cursor.fetchone()
    if not subsection:
        await callback_query.message.answer("❌ Подраздел не найден.")
        conn.close()
        return
    subsection_name = subsection[0]

    # Получаем название раздела
    cursor.execute("SELECT name FROM sections WHERE id = ?", (callback_data.section_id,))
    section = cursor.fetchone()
    section_name = section[0] if section else "Неизвестный раздел"
    conn.close()

    await state.update_data(
        subsection_id=callback_data.subsection_id,
        subsection_name=subsection_name
    )

    # Переходим к выбору даты поставки
    await show_date_selection(callback_query.message, state)

async def show_date_selection(message: types.Message, state: FSMContext):
    """Показывает выбор даты поставки"""
    data = await state.get_data()
    # Создаем кнопки с датами
    today = datetime.now().date()
    tomorrow = today + timedelta(days=1)
    after_tomorrow = today + timedelta(days=2)

    builder = InlineKeyboardBuilder()
    builder.button(
        text=f"Сегодня ({today.strftime('%d.%m.%Y')})",
        callback_data=DateCallback(action="quick", date=today.isoformat())
    )
    builder.button(
        text=f"Завтра ({tomorrow.strftime('%d.%m.%Y')})",
        callback_data=DateCallback(action="quick", date=tomorrow.isoformat())
    )
    builder.button(
        text=f"Послезавтра ({after_tomorrow.strftime('%d.%m.%Y')})",
        callback_data=DateCallback(action="quick", date=after_tomorrow.isoformat())
    )
    builder.button(
        text="📅 Выбрать другую дату",
        callback_data=DateCallback(action="select")
    )
    builder.button(
        text="❌ Отмена",
        callback_data="cancel_action"
    )
    builder.adjust(1)

    section_name = data.get('section_name', 'Не указан')
    subsection_name = data.get('subsection_name', 'Не указан')
    await message.answer(
        f"📂 **Раздел:** {section_name}\n"
        f"📁 **Подраздел:** {subsection_name}\n\n"
        "📅 **Выберите дату поставки:**",
        reply_markup=builder.as_markup()
    )
    await state.set_state(CreateRequestStates.waiting_for_delivery_date)

@dp.callback_query(SubsectionCallback.filter(F.action == "create"))
async def create_subsection_during_request(
    callback_query: types.CallbackQuery,
    callback_data: SubsectionCallback,
    state: FSMContext
):
    """Создание подраздела в процессе заявки"""
    await callback_query.answer()
    await state.update_data(
        creating_subsection_in_request=True,
        temp_section_id=callback_data.section_id
    )
    # Получаем название раздела
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sections WHERE id = ?", (callback_data.section_id,))
    section = cursor.fetchone()
    conn.close()
    section_name = section[0] if section else "Неизвестный раздел"
    await callback_query.message.answer(
        f"📝 Введите название нового подраздела для раздела '{section_name}':"
    )
    await state.set_state(CreateRequestStates.waiting_for_custom_subsection)

@dp.message(CreateRequestStates.waiting_for_custom_subsection)
async def handle_custom_subsection_creation(message: types.Message, state: FSMContext):
    """Обрабатывает создание подраздела в процессе заявки"""
    subsection_name = message.text.strip()
    if not subsection_name:
        await message.answer("❌ Название подраздела не может быть пустым.")
        return

    data = await state.get_data()
    section_id = data.get('temp_section_id')

    # Проверяем, существует ли уже такой подраздел
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id FROM subsections WHERE name = ? AND section_id = ?",
        (subsection_name, section_id)
    )
    existing = cursor.fetchone()
    if existing:
        await message.answer(f"❌ Подраздел '{subsection_name}' уже существует в этом разделе.")
        conn.close()
        return

    # Создаем подраздел
    cursor.execute(
        "INSERT INTO subsections (name, section_id, created_by) VALUES (?, ?, ?)",
        (subsection_name, section_id, message.from_user.id)
    )
    conn.commit()
    subsection_id = cursor.lastrowid

    # Получаем название раздела
    cursor.execute("SELECT name FROM sections WHERE id = ?", (section_id,))
    section = cursor.fetchone()
    section_name = section[0] if section else "Неизвестный раздел"
    conn.close()

    await state.update_data(
        subsection_id=subsection_id,
        subsection_name=subsection_name,
        creating_subsection_in_request=False,
        temp_section_id=None
    )
    await message.answer(
        f"✅ Подраздел '{subsection_name}' создан в разделе '{section_name}'!\n\n"
        "Переходим к выбору даты поставки..."
    )
    # Переходим к выбору даты
    await show_date_selection(message, state)

@dp.callback_query(DateCallback.filter(F.action == "quick"))
async def select_quick_date(
    callback_query: types.CallbackQuery,
    callback_data: DateCallback,
    state: FSMContext
):
    """Обрабатывает выбор быстрой даты"""
    await callback_query.answer()
    try:
        # Парсим дату
        selected_date = datetime.fromisoformat(callback_data.date).date()
        formatted_date = selected_date.strftime("%d.%m.%Y")
        await state.update_data(delivery_date=callback_data.date)
        data = await state.get_data()
        await callback_query.message.answer(
            f"📅 **Дата поставки:** {formatted_date}\n\n"
            "📝 **Введите наименование позиции:**\n"
            "(можно ввести несколько через запятую, или по одной)"
        )
        await state.set_state(CreateRequestStates.waiting_for_product_name)
    except Exception as e:
        await callback_query.message.answer(f"❌ Ошибка обработки даты: {e}")

@dp.callback_query(DateCallback.filter(F.action == "select"))
async def select_custom_date(
    callback_query: types.CallbackQuery,
    state: FSMContext
):
    """Запрос на ввод произвольной даты"""
    await callback_query.answer()
    await callback_query.message.answer(
        "📅 **Введите дату в формате ДД.ММ.ГГГГ:**\n"
        "Например: 15.12.2024"
    )
    await state.set_state(CreateRequestStates.waiting_for_custom_date)

@dp.message(CreateRequestStates.waiting_for_custom_date)
async def process_custom_date(message: types.Message, state: FSMContext):
    """Обрабатывает ввод произвольной даты"""
    date_text = message.text.strip()
    try:
        # Пробуем разные форматы даты
        date_formats = ["%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"]
        parsed_date = None
        for fmt in date_formats:
            try:
                parsed_date = datetime.strptime(date_text, fmt).date()
                break
            except ValueError:
                continue
        if not parsed_date:
            await message.answer("❌ Неверный формат даты. Введите дату в формате ДД.ММ.ГГГГ:")
            return

        if parsed_date < datetime.now().date():
            await message.answer("❌ Дата не может быть в прошлом. Введите будущую дату:")
            return

        formatted_date = parsed_date.strftime("%d.%m.%Y")
        await state.update_data(delivery_date=parsed_date.isoformat())
        data = await state.get_data()
        await message.answer(
            f"📅 **Дата поставки:** {formatted_date}\n\n"
            "📝 **Введите наименование позиции:**\n"
            "(можно ввести несколько через запятую, или по одной)"
        )
        await state.set_state(CreateRequestStates.waiting_for_product_name)
    except Exception as e:
        await message.answer(f"❌ Ошибка: {str(e)}\nВведите дату в формате ДД.ММ.ГГГГ:")

@dp.message(CreateRequestStates.waiting_for_product_name)
async def process_product_name(message: types.Message, state: FSMContext):
    """Обрабатывает ввод наименования позиции"""
    product_names = [name.strip() for name in message.text.split(',') if name.strip()]
    if not product_names:
        await message.answer("❌ Наименование не может быть пустым. Введите снова:")
        return

    # Сохраняем позиции
    data = await state.get_data()
    items = data.get('items', [])
    for product_name in product_names:
        items.append({
            'product_name': product_name,
            'unit': None,
            'quantity': None,
            'index': len(items)
        })
    await state.update_data(items=items, current_item_index=0)

    # Показываем выбор единицы измерения для первой позиции
    if len(product_names) > 1:
        await message.answer(
            f"✅ Добавлено {len(product_names)} позиций.\n\n"
            f"Теперь для каждой позиции нужно выбрать единицу измерения.\n"
            f"Начнем с первой позиции: **{product_names[0]}**"
        )
    else:
        await message.answer(
            f"✅ Добавлена позиция: **{product_names[0]}**"
        )
    await show_unit_selection(message, state, product_names[0])

async def show_unit_selection(message: types.Message, state: FSMContext, product_name: str):
    """Показывает выбор единицы измерения"""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM common_units ORDER BY name")
    common_units = [row[0] for row in cursor.fetchall()]
    conn.close()

    builder = InlineKeyboardBuilder()
    # Часто используемые единицы
    for unit in common_units:
        builder.button(
            text=unit,
            callback_data=UnitCallback(action="select", name=unit)
        )
    builder.button(
        text="➕ Другая единица",
        callback_data=UnitCallback(action="create")
    )
    builder.button(
        text="❌ Отмена",
        callback_data="cancel_action"
    )
    builder.adjust(3)  # 3 кнопки в ряду

    await message.answer(
        f"📦 **Позиция:** {product_name}\n\n"
        "📏 **Выберите единицу измерения:**",
        reply_markup=builder.as_markup()
    )
    await state.set_state(CreateRequestStates.waiting_for_unit)

@dp.callback_query(UnitCallback.filter(F.action == "select"))
async def select_unit(
    callback_query: types.CallbackQuery,
    callback_data: UnitCallback,
    state: FSMContext
):
    """Обрабатывает выбор единицы измерения"""
    await callback_query.answer()
    data = await state.get_data()
    items = data.get('items', [])
    current_index = data.get('current_item_index', 0)
    if current_index < len(items):
        items[current_index]['unit'] = callback_data.name
        await state.update_data(items=items)
        product_name = items[current_index]['product_name']
        await callback_query.message.answer(
            f"📏 **Единица измерения:** {callback_data.name}\n\n"
            f"🔢 **Введите количество для позиции '{product_name}':**"
        )
        await state.set_state(CreateRequestStates.waiting_for_quantity)

@dp.callback_query(UnitCallback.filter(F.action == "create"))
async def create_custom_unit(
    callback_query: types.CallbackQuery,
    state: FSMContext
):
    """Создание пользовательской единицы измерения"""
    await callback_query.answer()
    await callback_query.message.answer(
        "📝 **Введите название новой единицы измерения:**\n"
        "(например: 'упаковка', 'пара', 'комплект')"
    )
    await state.set_state(CreateRequestStates.waiting_for_custom_unit)

@dp.message(CreateRequestStates.waiting_for_custom_unit)
async def process_custom_unit(message: types.Message, state: FSMContext):
    """Обрабатывает ввод пользовательской единицы измерения"""
    unit_name = message.text.strip()
    if not unit_name:
        await message.answer("❌ Название единицы не может быть пустым. Введите снова:")
        return

    # Сохраняем новую единицу в базу
    conn = get_connection()
    cursor = conn.cursor()
    # Проверяем, существует ли уже
    cursor.execute("SELECT id FROM common_units WHERE name = ?", (unit_name,))
    existing = cursor.fetchone()
    if not existing:
        cursor.execute(
            "INSERT INTO common_units (name, created_by) VALUES (?, ?)",
            (unit_name, message.from_user.id)
        )
        conn.commit()
    conn.close()

    # Используем эту единицу для текущей позиции
    data = await state.get_data()
    items = data.get('items', [])
    current_index = data.get('current_item_index', 0)
    if current_index < len(items):
        items[current_index]['unit'] = unit_name
        await state.update_data(items=items)
        product_name = items[current_index]['product_name']
        await message.answer(
            f"✅ Новая единица измерения '{unit_name}' добавлена!\n\n"
            f"📏 **Единица измерения:** {unit_name}\n\n"
            f"🔢 **Введите количество для позиции '{product_name}':**"
        )
        await state.set_state(CreateRequestStates.waiting_for_quantity)

@dp.message(CreateRequestStates.waiting_for_quantity)
async def process_quantity(message: types.Message, state: FSMContext):
    """Обрабатывает ввод количества"""
    try:
        quantity_text = message.text.replace(',', '.').strip()
        quantity = float(quantity_text)
        if quantity <= 0:
            await message.answer("❌ Количество должно быть положительным. Введите снова:")
            return

        data = await state.get_data()
        items = data.get('items', [])
        current_index = data.get('current_item_index', 0)
        if current_index < len(items):
            items[current_index]['quantity'] = quantity
            await state.update_data(items=items)

            # Проверяем, есть ли еще позиции без единицы и количества
            next_index = current_index + 1
            while next_index < len(items) and items[next_index]['unit'] is not None and items[next_index][
                'quantity'] is not None:
                next_index += 1

            if next_index < len(items):
                # Есть еще позиции для обработки
                await state.update_data(current_item_index=next_index)
                product_name = items[next_index]['product_name']
                await message.answer(
                    f"✅ Позиция {current_index + 1} из {len(items)} добавлена!\n\n"
                    f"Переходим к следующей позиции: **{product_name}**"
                )
                await show_unit_selection(message, state, product_name)
            else:
                # Все позиции обработаны
                await message.answer(
                    f"✅ Все позиции добавлены! Всего: {len(items)} позиций.\n\n"
                    "🔍 **Проверяем орфографию через AI...**"
                )
                # Проверяем орфографию
                items_to_check = [{'product_name': item['product_name']} for item in items]
                corrected_items = await check_list_spelling(items_to_check)

                # Обновляем items с исправленными названиями
                has_corrections = False
                for i, corrected in enumerate(corrected_items):
                    if corrected.get('has_correction') and corrected.get('corrected_name'):
                        items[i]['corrected_name'] = corrected['corrected_name']
                        items[i]['has_correction'] = True
                        has_corrections = True
                    else:
                        items[i]['corrected_name'] = items[i]['product_name']
                        items[i]['has_correction'] = False

                await state.update_data(items=items)

                # Показываем результат проверки
                await show_correction_results(message, state, has_corrections)
    except ValueError:
        await message.answer("❌ Пожалуйста, введите число (можно с десятичной точкой):")

async def show_correction_results(message: types.Message, state: FSMContext, has_corrections: bool):
    """Показывает результаты проверки орфографии"""
    data = await state.get_data()
    items = data.get('items', [])
    if has_corrections:
        correction_text = "🔍 **Результаты проверки орфографии:**\n\n"
        for i, item in enumerate(items):
            if item.get('has_correction') and item.get('corrected_name') != item['product_name']:
                correction_text += f"{i + 1}. **Было:** {item['product_name']}\n"
                correction_text += f" **Исправление:** {item['corrected_name']}\n\n"
    else:
        correction_text = "✅ **Ошибок не найдено!**\n\n"

    # Создаем клавиатуру
    builder = InlineKeyboardBuilder()
    if has_corrections:
        builder.button(
            text="✅ Принять исправления",
            callback_data=RequestCallback(action="confirm_correction", item_index=-1)
        )
        builder.button(
            text="✏️ Оставить как есть",
            callback_data=RequestCallback(action="finish")
        )
    else:
        builder.button(
            text="✅ Продолжить",
            callback_data=RequestCallback(action="finish")
        )
    builder.button(
        text="➕ Добавить ещё позиции",
        callback_data=RequestCallback(action="add_more")
    )
    builder.button(
        text="❌ Отменить заявку",
        callback_data="cancel_action"
    )
    builder.adjust(1)

    await message.answer(
        correction_text + "**Выберите действие:**",
        reply_markup=builder.as_markup()
    )
    await state.set_state(CreateRequestStates.waiting_for_correction)

@dp.callback_query(RequestCallback.filter(F.action == "confirm_correction"))
async def confirm_corrections(
    callback_query: types.CallbackQuery,
    state: FSMContext
):
    """Подтверждает исправления орфографии"""
    await callback_query.answer()
    data = await state.get_data()
    items = data.get('items', [])
    # Применяем исправления
    for item in items:
        if item.get('corrected_name'):
            item['product_name'] = item['corrected_name']
    await state.update_data(items=items)
    # Показываем итоговый список
    await show_final_summary(callback_query.message, state)

@dp.callback_query(RequestCallback.filter(F.action == "add_more"))
async def add_more_items(
    callback_query: types.CallbackQuery,
    state: FSMContext
):
    """Добавляет еще позиции"""
    await callback_query.answer()
    await callback_query.message.answer(
        "📝 **Введите наименование позиции:**\n"
        "(можно ввести несколько через запятую, или по одной)"
    )
    await state.set_state(CreateRequestStates.waiting_for_product_name)

@dp.callback_query(RequestCallback.filter(F.action == "finish"))
async def finish_request(
    callback_query: types.CallbackQuery,
    state: FSMContext
):
    """Завершает создание заявки"""
    await callback_query.answer()
    await show_final_summary(callback_query.message, state)

async def show_final_summary(message: types.Message, state: FSMContext):
    """Показывает итоговую сводку и сохраняет заявку"""
    data = await state.get_data()
    items = data.get('items', [])
    if not items:
        await message.answer("❌ В заявке нет позиций. Заявка не создана.")
        await state.clear()
        return

    # Получаем информацию о теме (для ссылки на таблицу)
    sheet_url = None
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT sheet_url FROM allowed_topics WHERE chat_id = ? AND thread_id = ?",
            (message.chat.id, message.message_thread_id or 0)
        )
        sheet_info = cursor.fetchone()
        if sheet_info:
            sheet_url = sheet_info[0]
        conn.close()
    except Exception as e:
        logging.error(f"Error getting sheet URL: {e}")

    # Создаем заявку в базе данных
    conn = get_connection()
    cursor = conn.cursor()
    # Вставляем заявку
    cursor.execute("""
        INSERT INTO requests (user_id, chat_id, thread_id, section_id, subsection_id, delivery_date, status)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        message.from_user.id,
        message.chat.id,
        message.message_thread_id or 0,
        data.get('section_id'),
        data.get('subsection_id'),
        data.get('delivery_date'),
        'pending'
    ))
    request_id = cursor.lastrowid

    # Вставляем позиции
    for item in items:
        cursor.execute("""
            INSERT INTO request_items (request_id, product_name, unit, quantity, corrected_name)
            VALUES (?, ?, ?, ?, ?)
        """, (
            request_id,
            item['product_name'],
            item['unit'],
            item['quantity'],
            item.get('corrected_name')
        ))
    conn.commit()

    # Получаем информацию о разделе и подразделе
    cursor.execute("""
        SELECT s.name, sb.name, r.delivery_date
        FROM requests r
        LEFT JOIN sections s ON r.section_id = s.id
        LEFT JOIN subsections sb ON r.subsection_id = sb.id
        WHERE r.id = ?
    """, (request_id,))
    request_info = cursor.fetchone()
    conn.close()

    # Форматируем дату
    delivery_date = ""
    if request_info and request_info[2]:
        try:
            date_obj = datetime.fromisoformat(request_info[2]).date()
            delivery_date = date_obj.strftime("%d.%m.%Y")
        except:
            delivery_date = request_info[2]

    # Формируем сводку
    summary_text = (
        f"✅ **Заявка №{request_id} создана!**\n\n"
        f"📂 **Раздел:** {request_info[0] if request_info and request_info[0] else 'Не указан'}\n"
        f"📁 **Подраздел:** {request_info[1] if request_info and request_info[1] else 'Не указан'}\n"
        f"📅 **Дата поставки:** {delivery_date if delivery_date else 'Не указана'}\n\n"
        f"📦 **Позиции ({len(items)}):**\n"
    )
    total_quantity = 0
    for i, item in enumerate(items, 1):
        summary_text += f"{i}. {item['product_name']} - {item['quantity']} {item['unit']}\n"
        try:
            total_quantity += float(item['quantity'])
        except:
            pass
    summary_text += f"\n📊 **Всего единиц:** {total_quantity}"

    # Если есть ссылка на таблицу, добавляем её
    if sheet_url:
        summary_text += f"\n\n🔗 **Google Sheet:** {sheet_url}"

    summary_text += "\n📊 Данные заявки сохранены в Google Sheet!"

    # Отправляем сообщение
    await message.answer(summary_text)

    # Сохраняем в Google Sheet
    try:
        if sheet_url:
            # Извлекаем ID таблицы из URL
            import re
            match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheet_url)
            if match:
                sheet_id = match.group(1)
                # Подготавливаем данные для таблицы
                rows = []
                for item in items:
                    rows.append([
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        request_id,
                        data.get('section_name', ''),
                        data.get('subsection_name', ''),
                        delivery_date,
                        item['product_name'],
                        item['unit'],
                        str(item['quantity']),
                        message.from_user.id
                    ])
                # Добавляем в таблицу - ИСПРАВЛЕНО: добавлен третий аргумент
                result = append_to_sheet(sheet_id, "Заявки", rows)  # <- Добавлен "Заявки"
                if not result:
                    await message.answer("⚠️ Не удалось сохранить данные в Google Sheet")
    except Exception as e:
        logging.error(f"Error saving to Google Sheet: {e}")
        await message.answer("⚠️ Не удалось сохранить данные в Google Sheet")

    await state.clear()

# ================== CANCEL HANDLERS ==================
@dp.callback_query(lambda c: c.data == "cancel_action")
async def cancel_action(callback_query: types.CallbackQuery, state: FSMContext):
    """Отменяет текущее действие"""
    await callback_query.answer()
    await state.clear()
    await callback_query.message.answer(
        "❌ Действие отменено.",
        reply_markup=get_menu_inline_keyboard()
    )

@dp.callback_query(lambda c: c.data == "back_to_menu")
async def back_to_menu(callback_query: types.CallbackQuery, state: FSMContext):
    """Возвращает в главное меню"""
    await callback_query.answer()
    await state.clear()
    await menu_command(callback_query.message)

@dp.message(Command("cancel"))
async def cancel_command(message: types.Message, state: FSMContext):
    """Отменяет текущий процесс по команде"""
    await state.clear()
    await message.answer(
        "❌ Текущий процесс отменён.",
        reply_markup=get_menu_inline_keyboard()
    )

# ================== COMMAND HANDLERS ==================
@dp.message(Command("create_request"))
async def create_request_command(message: types.Message, state: FSMContext):
    """Команда для создания заявки"""
    # Создаем фейковый callback_query для использования существующего обработчика
    class FakeCallbackQuery:
        def __init__(self, message, user):
            self.message = message
            self.from_user = user
            self.data = "create_request_start"

        async def answer(self):
            pass

    fake_callback = FakeCallbackQuery(message, message.from_user)
    await process_create_request_start(fake_callback, state)

@dp.message(Command("create_section"))
async def create_section_command(message: types.Message, state: FSMContext):
    """Команда для создания раздела"""
    await create_section_menu(
        types.CallbackQuery(
            message=message,
            from_user=message.from_user,
            data="create_section_menu"
        ),
        state
    )

# ================== EXISTING REGISTRATION HANDLERS ==================
@dp.message(Command("start"))
async def start_command(message: types.Message, state: FSMContext):
    logging.info(
        "CMD /start from chat_id=%s title=%s thread=%s",
        message.chat.id,
        message.chat.title,
        message.message_thread_id,
    )
    if not await is_allowed_context(message):
        return
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE user_id = ?", (message.from_user.id,))
    user = cursor.fetchone()
    conn.close()
    if user:
        await message.answer(
            "✅ Вы уже зарегистрированы в системе.\n\n"
            "Что хотите сделать?",
            reply_markup=get_menu_inline_keyboard()
        )
    else:
        await message.answer("👋 Добро пожаловать! Для регистрации введите ваши ФИО:")
        await state.set_state(RegistrationStates.waiting_for_fio)

@dp.message(RegistrationStates.waiting_for_fio)
async def process_fio(message: types.Message, state: FSMContext):
    if not await is_allowed_context(message):
        await state.clear()
        return
    await state.update_data(fio=message.text)
    await message.reply("📝 Введите вашу должность (если есть, иначе напишите '-'):")
    await state.set_state(RegistrationStates.waiting_for_position)

@dp.message(RegistrationStates.waiting_for_position)
async def process_position(message: types.Message, state: FSMContext):
    if not await is_allowed_context(message):
        await state.clear()
        return
    await state.update_data(position=message.text)
    await message.reply("📞 Введите ваш контактный номер телефона:")
    await state.set_state(RegistrationStates.waiting_for_phone)

@dp.message(RegistrationStates.waiting_for_phone)
async def process_phone(message: types.Message, state: FSMContext):
    if not await is_allowed_context(message):
        await state.clear()
        return
    data = await state.get_data()
    fio = data["fio"]
    position = data["position"]
    phone = message.text
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR REPLACE INTO users VALUES (?, ?, ?, ?)",
        (message.from_user.id, fio, position, phone),
    )
    conn.commit()
    conn.close()
    await message.reply(
        "✅ Регистрация завершена!\n\n"
        "Ваши данные:\n"
        f"👤 ФИО: {fio}\n"
        f"💼 Должность: {position if position != '-' else 'Не указана'}\n"
        f"📞 Телефон: {phone}\n\n"
        "Используйте команду /menu для вызова главного меню."
    )
    await state.clear()

# ================== BOT REGISTRATION HANDLERS ==================
@dp.message(Command("register_bot"))
async def register_bot_command(message: types.Message, state: FSMContext):
    """Активирует бота в текущей теме (только для администраторов бота)"""
    # Проверка прав администратора бота
    if not is_admin(message.from_user.id):
        await message.answer(
            "❌ Эта команда доступна только администраторам бота.\n\n"
            "Для получения доступа обратитесь к администратору системы."
        )
        return

    if message.chat.type == "private":
        await message.answer("❌ Эта команда работает только в темах групп.")
        return

    thread_id = message.message_thread_id
    if not thread_id or thread_id == 0:
        await message.answer("❌ Эта команда работает только внутри тем.")
        return

    # Проверяем права пользователя в группе (должен быть админом группы)
    if not await is_user_admin(message.chat.id, message.from_user.id, message.bot):
        await message.answer(
            "❌ Для активации бота в теме вы должны быть администратором этой группы.\n\n"
            "Права администратора бота и администратора группы - это разные вещи.\n"
            "Вы администратор бота, но не администратор этой группы."
        )
        return

    chat_id = message.chat.id

    # Проверяем, не зарегистрирована ли уже тема
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT 1 FROM allowed_topics WHERE chat_id = ? AND thread_id = ?",
        (chat_id, thread_id)
    )
    result = cursor.fetchone()
    conn.close()
    if result:
        await message.answer("✅ Бот уже активирован в этой теме.")
        return

    # Начинаем процесс регистрации...
    await message.answer(
        "🤖 **Начинаем настройку бота для этой темы**\n\n"
        "Для завершения регистрации нужно создать Google Sheet для этой темы.\n"
        "**Каждая тема имеет только один Google Sheet!**\n\n"
        "🔢 **Введите код объекта (без слеша /):**\n"
        "Например: 'SKL-001', 'OF-2024', 'TC-MOS'"
    )
    await state.update_data(
        chat_id=chat_id,
        thread_id=thread_id,
        admin_id=message.from_user.id,
        admin_name=message.from_user.full_name
    )
    await state.set_state(BotRegistrationStates.waiting_for_object_code)

@dp.message(BotRegistrationStates.waiting_for_object_code)
async def process_bot_object_code(message: types.Message, state: FSMContext):
    """Обрабатывает ввод кода объекта (ПЕРВЫЙ ШАГ)"""
    object_code = message.text.strip()
    # Проверяем, не команда ли это (начинается с /)
    if object_code.startswith('/'):
        await message.answer("❌ Пожалуйста, введите код объекта БЕЗ слеша /")
        return
    if len(object_code) < 2:
        await message.answer("❌ Код объекта слишком короткий. Введите ещё раз:")
        return

    # Сохраняем код объекта в FSM (теперь в Redis)
    await state.update_data(object_code=object_code)
    await message.answer(
        f"✅ **Код объекта:** {object_code}\n\n"
        "📝 **Теперь введите наименование объекта (без слеша /):**\n"
        "Например: 'Склад №1', 'Офисное здание', 'Торговый центр'"
    )
    # Переходим ко второму шагу
    await state.set_state(BotRegistrationStates.waiting_for_object_name)

@dp.message(BotRegistrationStates.waiting_for_object_name)
async def process_bot_object_name(message: types.Message, state: FSMContext):
    """Обрабатывает ввод наименования объекта (ВТОРОЙ ШАГ)"""
    object_name = message.text.strip()
    # Проверяем, не команда ли это
    if object_name.startswith('/'):
        await message.answer("❌ Пожалуйста, введите наименование объекта БЕЗ слеша /")
        return
    if len(object_name) < 2:
        await message.answer("❌ Наименование объекта слишком короткое. Введите ещё раз:")
        return

    # Получаем все данные из FSM (теперь из Redis)
    data = await state.get_data()
    chat_id = data.get("chat_id")
    thread_id = data.get("thread_id")
    admin_id = data.get("admin_id")
    admin_name = data.get("admin_name")
    object_code = data.get("object_code")

    if not all([chat_id, thread_id, admin_id, object_code]):
        await message.answer("❌ Ошибка: данные регистрации не найдены. Начните заново с /register_bot")
        await state.clear()
        return

    try:
        # Создаём Google Sheet для темы
        sheet_title = f"{object_code} - {object_name}"
        sheet_id = create_spreadsheet(sheet_title)
        sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"

        # Регистрируем тему в базе данных
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            """INSERT INTO allowed_topics (chat_id, thread_id, registered_by, object_name, object_code, sheet_id, sheet_url)
            VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (chat_id, thread_id, admin_id, object_name, object_code, sheet_id, sheet_url)
        )
        conn.commit()
        conn.close()

        # Добавляем в кэш
        cache_key = (chat_id, thread_id)
        _allowed_topics_cache[cache_key] = {
            "admin_id": admin_id,
            "registered_at": datetime.now(),
            "object_name": object_name,
            "object_code": object_code,
            "sheet_id": sheet_id
        }

        # Отправляем сообщение об успешной регистрации
        await message.answer(
            f"✅ **Бот успешно активирован и настроен!**\n\n"
            f"**📊 Создан Google Sheet:**\n"
            f"• Код: {object_code}\n"
            f"• Наименование: {object_name}\n\n"
            f"**🔗 Ссылка на таблицу:**\n"
            f"{sheet_url}\n\n"
            f"**👤 Активировал:** {admin_name}\n"
            f"**📌 ID темы:** {thread_id}\n\n"
            f"Теперь бот готов к работе в этой теме!\n"
            f"Используйте /menu для доступа к основным функциям."
        )

        # Также отправляем кнопку для быстрого доступа к таблице
        keyboard = types.InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    types.InlineKeyboardButton(
                        text="📊 Открыть Google Sheet",
                        url=sheet_url
                    )
                ]
            ]
        )
        await message.answer("Нажмите для открытия таблицы:", reply_markup=keyboard)
    except Exception as e:
        logging.error(f"Error during bot registration: {e}")
        await message.answer(
            f"❌ **Ошибка при настройке бота:**\n\n"
            f"Ошибка: {str(e)}\n\n"
            f"Попробуйте снова командой `/register_bot`"
        )
    finally:
        await state.clear()

@dp.message(Command("unregister_bot"))
async def unregister_bot_command(message: types.Message):
    """Деактивирует бота в текущей теме (только для администраторов бота)"""
    # Проверка прав администратора бота
    if not is_admin(message.from_user.id):
        await message.answer(
            "❌ Эта команда доступна только администраторам бота.\n\n"
            "Для получения доступа обратитесь к администратору системы."
        )
        return

    if message.chat.type == "private":
        await message.answer("❌ Эта команда работает только в темах групп.")
        return

    thread_id = message.message_thread_id
    if not thread_id or thread_id == 0:
        await message.answer("❌ Эта команда работает только внутри тем.")
        return

    # Проверяем права пользователя в группе
    if not await is_user_admin(message.chat.id, message.from_user.id, message.bot):
        await message.answer("❌ Только администраторы группы могут деактивировать бота в теме.")
        return

    chat_id = message.chat.id
    cache_key = (chat_id, thread_id)

    # Удаляем из базы данных
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "DELETE FROM allowed_topics WHERE chat_id = ? AND thread_id = ?",
        (chat_id, thread_id)
    )
    deleted_rows = cursor.rowcount
    conn.commit()
    conn.close()

    # Удаляем из кэша
    if cache_key in _allowed_topics_cache:
        del _allowed_topics_cache[cache_key]

    if deleted_rows > 0:
        await message.answer("✅ Бот деактивирован в этой теме.")
    else:
        await message.answer("ℹ️ Бот не был активирован в этой теме.")

# ================== TOPIC INFO HANDLER ==================
@dp.message(Command("topic_info"))
async def topic_info_command(message: types.Message):
    """Показывает информацию о текущей теме и ссылку на Google Sheet"""
    if not await is_allowed_context(message):
        return
    thread_id = message.message_thread_id
    chat_id = message.chat.id
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT object_name, object_code, sheet_url, registered_by FROM allowed_topics WHERE chat_id = ? AND thread_id = ?",
        (chat_id, thread_id)
    )
    result = cursor.fetchone()
    conn.close()
    if result:
        object_name, object_code, sheet_url, admin_id = result
        info_text = (
            f"📋 **Информация о теме:**\n\n"
            f"**🏢 Объект:**\n"
            f"• Наименование: {object_name}\n"
            f"• Код: {object_code}\n\n"
            f"**📊 Google Sheet (один на тему):**\n"
            f"• Ссылка: {sheet_url}\n\n"
            f"**👤 Активатор:** ID {admin_id}"
        )
        # Добавляем кнопку для открытия таблицы
        keyboard = types.InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    types.InlineKeyboardButton(
                        text="📊 Открыть Google Sheet",
                        url=sheet_url
                    )
                ]
            ]
        )
        await message.answer(info_text, reply_markup=keyboard)
    else:
        await message.answer(
            "ℹ️ Тема не активирована. Для активации используйте команду /register_bot (только для администраторов).",
            reply_markup=get_menu_inline_keyboard()
        )

# ================== BOT ADDED TO GROUP HANDLER ==================
@dp.my_chat_member()
async def on_bot_added_to_group(update: types.ChatMemberUpdated):
    """Обработчик добавления бота в группу"""
    if update.new_chat_member.status in ("member", "administrator"):
        adder_id = update.from_user.id
        chat = update.chat
        try:
            chat_info = await update.bot.get_chat(chat.id)
            is_forum = getattr(chat_info, 'is_forum', False)
        except:
            is_forum = False

        if is_forum:
            await update.bot.send_message(
                adder_id,
                "✅ **Бот добавлен в форум-группу**\n\n"
                "**Для настройки:**\n"
                "1. Перейдите в тему, где должен работать бот\n"
                "2. Введите команду: `/register_bot`\n"
                "3. Следуйте инструкции для создания Google Sheet\n\n"
                "**Важно:**\n"
                "• Каждая тема имеет только один Google Sheet\n"
                "• Бот будет работать только в темах, где выполнена команда /register_bot!\n"
                "• Команда доступна только администраторам\n\n"
                "**Используйте команду `/menu` для вызова главного меню с кнопками**"
            )
        else:
            await update.bot.send_message(
                adder_id,
                "⚠️ **Внимание!**\n\n"
                "Этот бот предназначен для **форум-групп**.\n\n"
                "**Как настроить:**\n"
                "1. В настройках группы включите 'Форум'\n"
                "2. Создайте темы\n"
                "3. В нужных темах введите `/register_bot`\n\n"
                "**Используйте команду `/menu` для вызова главного меню с кнопками**"
            )
        return

# ================== HELP COMMAND ==================
@dp.message(Command("help"))
async def help_command(message: types.Message):
    if not await is_allowed_context(message):
        return
    help_text = (
        "📋 **Доступные команды:**\n\n"
        "**Основные:**\n"
        "• `/start` — регистрация в системе\n"
        "• `/menu` — главное меню с кнопками\n"
        "• `/create_request` — создать новую заявку\n"
        "• `/create_section` — создать новый раздел\n"
        "• `/topic_info` — информация о текущей теме\n"
        "• `/help` — эта справка\n"
        "• `/cancel` — отменить текущий процесс\n\n"
        "**Управление ботом в теме:**\n"
        "• `/register_bot` — активировать бота в этой теме (админы)\n"
        "• `/unregister_bot` — деактивировать бота в этой теме (админы)\n\n"
        "**Процесс создания заявки:**\n"
        "1. Выберите раздел (или создайте новый)\n"
        "2. Выберите подраздел (или создайте новый)\n"
        "3. Выберите дату поставки\n"
        "4. Вводите позиции: название, единица измерения, количество\n"
        "5. Бот проверяет орфографию через AI\n"
        "6. Завершите создание заявки\n\n"
        "**Важно:**\n"
        "• Каждая тема имеет только один Google Sheet\n"
        "• Все данные по теме хранятся в одной таблице\n\n"
        "**Использование меню:**\n"
        "Нажмите /menu для вызова меню с кнопками"
    )
    await message.answer(help_text)