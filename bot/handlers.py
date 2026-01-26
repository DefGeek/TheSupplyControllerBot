import logging
from typing import Dict, Tuple
from datetime import datetime
from aiogram import types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.enums import ChatMemberStatus

from bot import telegram_bot, dp
from core.database import get_connection
from core.sheets import create_spreadsheet

# Кэш разрешённых тем: {(chat_id, thread_id): {"admin_id": int, "registered_at": datetime}}
_allowed_topics_cache: Dict[tuple[int, int], dict] = {}


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

        # Если тема не зарегистрирована, предлагаем администраторам зарегистрировать её
        if await is_user_admin(chat.id, message.from_user.id, message.bot):
            await message.answer(
                "🤖 **Бот не активирован в этой теме**\n\n"
                "Для активации бота в этой теме выполните команду:\n"
                "`/register_bot`\n\n"
                "*Только администраторы могут активировать бота*"
            )
        return False

    return False


class RegistrationStates(StatesGroup):
    waiting_for_fio = State()
    waiting_for_position = State()
    waiting_for_phone = State()


class GroupStates(StatesGroup):
    waiting_for_object_name = State()
    waiting_for_object_code = State()


class BotRegistrationStates(StatesGroup):
    waiting_for_object_code = State()
    waiting_for_object_name = State()


# ================== ОБРАБОТЧИКИ РЕГИСТРАЦИИ БОТА ==================

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
        # Создаём Google Sheet
        sheet_title = f"{object_code} - {object_name}"
        sheet_id = create_spreadsheet(sheet_title)
        sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"

        # Регистрируем тему в базе данных
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            """INSERT INTO allowed_topics 
               (chat_id, thread_id, registered_by, object_name, object_code, sheet_id, sheet_url) 
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
            f"Теперь бот готов к работе в этой теме!"
        )

        # Также отправляем кнопку для быстрого доступа
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


# --- Регистрация бота в теме ---
@dp.message(Command("register_bot"))
async def register_bot_command(message: types.Message, state: FSMContext):
    """Активирует бота в текущей теме (только для администраторов)"""
    if message.chat.type == "private":
        await message.answer("❌ Эта команда работает только в темах групп.")
        return

    thread_id = message.message_thread_id
    if not thread_id or thread_id == 0:
        await message.answer("❌ Эта команда работает только внутри тем.")
        return

    # Проверяем права пользователя
    if not await is_user_admin(message.chat.id, message.from_user.id, message.bot):
        await message.answer("❌ Только администраторы могут активировать бота в теме.")
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

    # НАЧИНАЕМ ПРОЦЕСС: сначала код!
    await message.answer(
        "🤖 **Начинаем настройку бота для этой темы**\n\n"
        "Для завершения регистрации нужно создать Google Sheet для этой темы.\n\n"
        "🔢 **Введите код объекта (без слеша /):**\n"
        "Например: 'SKL-001', 'OF-2024', 'TC-MOS'"
    )

    # Сохраняем данные в FSM (теперь в Redis)
    await state.update_data(
        chat_id=chat_id,
        thread_id=thread_id,
        admin_id=message.from_user.id,
        admin_name=message.from_user.full_name
    )

    # Начинаем с ожидания кода объекта
    await state.set_state(BotRegistrationStates.waiting_for_object_code)


# --- Команда отмены регистрации ---
@dp.message(Command("cancel"))
async def cancel_registration(message: types.Message, state: FSMContext):
    """Отменяет текущий процесс регистрации"""
    await state.clear()
    await message.answer("❌ Процесс регистрации отменён.")


# --- Удаление регистрации темы (для администраторов) ---
@dp.message(Command("unregister_bot"))
async def unregister_bot_command(message: types.Message):
    """Деактивирует бота в текущей теме"""
    if message.chat.type == "private":
        await message.answer("❌ Эта команда работает только в темах групп.")
        return

    thread_id = message.message_thread_id
    if not thread_id or thread_id == 0:
        await message.answer("❌ Эта команда работает только внутри тем.")
        return

    # Проверяем права пользователя
    if not await is_user_admin(message.chat.id, message.from_user.id, message.bot):
        await message.answer("❌ Только администраторы могут деактивировать бота.")
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


# --- Команда /start ---
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
        await message.answer("✅ Вы уже зарегистрированы в системе.")
    else:
        await message.answer("👋 Добро пожаловать! Для регистрации введите ваши ФИО:")
        await state.set_state(RegistrationStates.waiting_for_fio)


# --- Обработчики регистрации пользователя ---
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

    await message.reply("✅ Регистрация завершена!")
    await state.clear()


# --- Команда для создания дополнительного Google Sheet ---
@dp.message(Command("create_sheet"))
async def create_sheet_command(message: types.Message, state: FSMContext):
    """Создаёт дополнительный Google Sheet в уже активированной теме"""
    if not await is_allowed_context(message):
        return

    # Проверяем, зарегистрирован ли пользователь
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE user_id = ?", (message.from_user.id,))
    user = cursor.fetchone()
    conn.close()

    if not user:
        await message.answer("❌ Сначала зарегистрируйтесь с помощью команды /start")
        return

    await message.answer(
        "📋 **Создание дополнительного Google Sheet**\n\n"
        "Введите наименование объекта:"
    )
    await state.set_state(GroupStates.waiting_for_object_name)


# --- Обработчики информации об объекте для создания доп. таблицы ---
@dp.message(GroupStates.waiting_for_object_name)
async def handle_object_name(message: types.Message, state: FSMContext):
    if not await is_allowed_context(message):
        await state.clear()
        return

    await state.update_data(object_name=message.text)
    await message.reply("🔢 Введите код объекта:")
    await state.set_state(GroupStates.waiting_for_object_code)


@dp.message(GroupStates.waiting_for_object_code)
async def handle_object_code(message: types.Message, state: FSMContext):
    if not await is_allowed_context(message):
        await state.clear()
        return

    data = await state.get_data()
    object_name = data["object_name"]
    object_code = message.text

    try:
        # Получаем информацию о пользователе
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT fio FROM users WHERE user_id = ?", (message.from_user.id,))
        user = cursor.fetchone()
        conn.close()

        user_fio = user[0] if user else "Неизвестный"

        # Создаём название таблицы
        sheet_title = f"{object_code} - {object_name} ({user_fio})"
        sheet_id = create_spreadsheet(sheet_title)
        sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"

        await message.reply(
            f"✅ **Дополнительный Google Sheet создан!**\n\n"
            f"**📊 Информация:**\n"
            f"• Название: {sheet_title}\n"
            f"• Код объекта: {object_code}\n"
            f"• Наименование: {object_name}\n"
            f"• Создатель: {user_fio}\n\n"
            f"**🔗 Ссылка:**\n{sheet_url}"
        )

        # Кнопка для открытия таблицы
        keyboard = types.InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    types.InlineKeyboardButton(
                        text="📊 Открыть таблицу",
                        url=sheet_url
                    )
                ]
            ]
        )

        await message.answer("Нажмите для открытия:", reply_markup=keyboard)

    except Exception as e:
        logging.error(f"Error creating additional spreadsheet: {e}")
        await message.reply(
            f"❌ **Ошибка при создании Google Sheet:**\n\n"
            f"{str(e)}"
        )
    finally:
        await state.clear()


# --- Обработчик добавления бота в группу ---
@dp.my_chat_member()
async def on_bot_added_to_group(update: types.ChatMemberUpdated):
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
                "**Бот будет работать только в темах, где выполнена команда /register_bot!**\n\n"
                "*Команда доступна только администраторам*"
            )
        else:
            await update.bot.send_message(
                adder_id,
                "⚠️ **Внимание!**\n\n"
                "Этот бот предназначен для **форум-групп**.\n\n"
                "**Как настроить:**\n"
                "1. В настройках группы включите 'Форум'\n"
                "2. Создайте темы\n"
                "3. В нужных темах введите `/register_bot`"
            )
        return


# --- Команда для просмотра информации о текущей теме ---
@dp.message(Command("topic_info"))
async def topic_info_command(message: types.Message):
    """Показывает информацию о текущей теме и связанном Google Sheet"""
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
            f"**📊 Google Sheet:**\n"
            f"• Ссылка: {sheet_url}\n\n"
            f"**👤 Активатор:** ID {admin_id}"
        )

        # Добавляем кнопку для открытия таблицу
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
        await message.answer("ℹ️ Информация о теме не найдена.")


# --- Справка ---
@dp.message(Command("help"))
async def help_command(message: types.Message):
    if not await is_allowed_context(message):
        return

    await message.answer(
        "📋 **Доступные команды:**\n\n"
        "**Основные:**\n"
        "• `/start` — регистрация в системе\n"
        "• `/create_sheet` — создать дополнительный Google Sheet\n"
        "• `/topic_info` — информация о текущей теме\n"
        "• `/help` — эта справка\n"
        "• `/cancel` — отменить текущий процесс регистрации\n\n"
        "**Управление ботом в теме:**\n"
        "• `/register_bot` — активировать бота в этой теме (админы)\n"
        "• `/unregister_bot` — деактивировать бота в этой теме (админы)\n\n"
        "**Процесс активации бота:**\n"
        "1. Введите `/register_bot` (только админы)\n"
        "2. Введите код объекта (без /)\n"
        "3. Введите название объекта (без /)\n"
        "4. Бот создаст таблицу и активируется"
    )