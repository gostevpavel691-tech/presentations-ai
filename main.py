"""
1. Заменена safe_execute_code на safe_execute_code_async с asyncio.create_subprocess_exec
2. Добавлен subprocess_semaphore = asyncio.Semaphore(10)
3. В _process_presentation_request_internal вызов теперь async with subprocess_semaphore:
4. В лог добавлено Subprocess: asyncio.create_subprocess_exec
5. После успешной генерации пользователь остаётся в режиме ввода (waiting_for_prompt),
   не возвращается в главное меню — может сразу ввести новый промпт или нажать Отмена
6. При ошибке выполнения кода (Mistral написал нерабочий код) показывается кнопка
   "🔧 Исправить" — по нажатию текст ошибки вставляется в промпт к Mistral для повторной генерации
"""

import asyncio
import os
import subprocess
import sys
import re
import uuid
import logging
import shutil
import traceback
import random
import signal
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
import sqlite3
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton

os.system('pip install mistralai==1.8.0')
os.system('pip install python-pptx')

from mistralai import Mistral

import config
from cache_storage import PresentationCache
from user_manager import DatabaseManager

# ==================== НАСТРОЙКИ ====================
BOT_TOKEN = config.BOT_TOKEN
MISTRAL_API_KEY = config.MISTRAL_API_KEY

# Пути (оставлены как есть, вы измените сами)
BASE_PATH = os.path.dirname(os.path.abspath(__file__)) + os.sep
TEMP_DIR = os.path.join(BASE_PATH, "temp")
COMPLETED_DIR = os.path.join(BASE_PATH, "completed")
CACHE_DIR = os.path.join(BASE_PATH, "cache")
DB_PATH = os.path.join(BASE_PATH, "users.db")

# Создаём папки
for dir_path in [TEMP_DIR, COMPLETED_DIR, CACHE_DIR]:
    os.makedirs(dir_path, exist_ok=True)

# Логирование
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ==================== ВЫБОР ТИПА КЭША ====================
CACHE_TYPE = "sqlite"

# ==================== ИНИЦИАЛИЗАЦИЯ ====================
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
mistral_client = Mistral(api_key=MISTRAL_API_KEY)

# Пул потоков для выполнения кода (оставляем для Mistral API)
executor = ThreadPoolExecutor(max_workers=10)

# Очередь и счетчики (с блокировкой для thread-safety)
task_queue = asyncio.Queue()
active_tasks = 0
active_tasks_lock = asyncio.Lock()
max_concurrent_tasks = 10

# Семафор для ограничения одновременных subprocess
subprocess_semaphore = asyncio.Semaphore(10)

# Словарь для хранения состояний пользователей
user_states = {}

# Словарь для хранения данных повторной генерации (retry по кнопке "Исправить")
# Формат: { user_id: {"original_request": str, "error_text": str} }
retry_data = {}

# ==================== АНТИ-СПАМ СИСТЕМА ====================
user_last_action = {}
ANTI_SPAM_SECONDS = 5

def check_spam(user_id: int) -> tuple[bool, int]:
    """Проверяет, не спамит ли пользователь."""
    now = datetime.now()
    if user_id in user_last_action:
        last_action = user_last_action[user_id]
        elapsed = (now - last_action).total_seconds()
        if elapsed < ANTI_SPAM_SECONDS:
            wait_seconds = int(ANTI_SPAM_SECONDS - elapsed) + 1
            return False, wait_seconds
    user_last_action[user_id] = now
    return True, 0

# Канал
CHANNEL_CHAT_ID = -1002281909552
CHANNEL_USERNAME = 'downoficeberg'
ADMIN_IDS = config.ADMIN_IDS

# Инициализация кэша и базы данных
presentation_cache = PresentationCache(CACHE_DIR, CACHE_TYPE)
db_manager = DatabaseManager(DB_PATH)

# Флаг для graceful shutdown
shutdown_flag = False

# ==================== ПРОВЕРКА ЗАВИСИМОСТЕЙ ====================
def check_dependencies():
    """Проверяет наличие необходимых библиотек"""
    try:
        import pptx
        logger.info("✅ python-pptx установлен")
    except ImportError:
        logger.error("❌ python-pptx НЕ УСТАНОВЛЕН!")
        logger.error("Установите: pip install python-pptx")
        return False
    return True

# ==================== КЛАВИАТУРЫ ====================
def get_main_keyboard():
    keyboard = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🆕 Создать")],
            [KeyboardButton(text="👑 Привилегии"), KeyboardButton(text="👤 Мой профиль")]
        ],
        resize_keyboard=True,
        one_time_keyboard=False
    )
    return keyboard

def get_cancel_keyboard():
    keyboard = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="❌ Отмена")]],
        resize_keyboard=True,
        one_time_keyboard=True
    )
    return keyboard

def remove_keyboard():
    return ReplyKeyboardRemove()

# ==================== ИНСТРУКЦИЯ ====================
INSTRUCTION = """
Ты — эксперт по созданию презентаций с помощью библиотеки python-pptx.

Твоя задача — генерировать Python код, который создаст красивую, информативную презентацию на заданную тему.

Создай файл "presentation.pptx" в текущей директории.

ВАЖНЫЕ ПРАВИЛА:
- НЕ используй внешние изображения, add_picture и любые картинки — их нет на сервере
- НЕ используй несуществующие атрибуты: вместо fill.fore_color используй fill.solid() затем fill.fore_color
- НЕ импортируй сторонние библиотеки кроме pptx и стандартных (os, random, math и т.д.)
- Для красоты используй: цветные фоны слайдов, градиентные фигуры, цветные прямоугольники, иконки из символов (✓ ★ → и т.д.), красивые шрифты и размеры

Отвечай ТОЛЬКО кодом на Python, без объяснений. Код должен быть готов к выполнению.
Постарайся сделать все без ошибок!
"""

# ==================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================
def extract_python_code(text):
    """Извлекает Python код из ответа Mistral."""
    pattern = r'```python\n(.*?)```'
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    pattern = r'```\n(.*?)```'
    match = re.search(pattern, text, re.DOTALL)
    if match:
        return match.group(1).strip()
    return text.strip()

def cleanup_old_temp_dirs():
    """Удаляет временные папки старше 1 часа."""
    now = datetime.now()
    if not os.path.exists(TEMP_DIR):
        return
    for item in os.listdir(TEMP_DIR):
        item_path = os.path.join(TEMP_DIR, item)
        if os.path.isdir(item_path):
            try:
                mtime = datetime.fromtimestamp(os.path.getmtime(item_path))
                if now - mtime > timedelta(hours=1):
                    shutil.rmtree(item_path)
                    logger.info(f"Удалена старая временная папка: {item_path}")
            except Exception as e:
                logger.error(f"Ошибка при очистке {item_path}: {e}")

def move_file_with_retry(src, dst, max_retries=3):
    """Перемещает файл с повторными попытками при блокировке"""
    for i in range(max_retries):
        try:
            if os.path.exists(dst):
                os.remove(dst)
            shutil.move(src, dst)
            return True
        except (PermissionError, OSError) as e:
            if i == max_retries - 1:
                raise
            time.sleep(0.5)
    return False

async def safe_execute_code_async(code, work_dir):
    """Асинхронное выполнение Python-кода (asyncio.create_subprocess_exec)"""
    try:
        code_file = os.path.join(work_dir, "script.py")
        
        # Сохраняем код с UTF-8
        with open(code_file, 'w', encoding='utf-8') as f:
            f.write(code)
        
        # Асинхронный запуск subprocess
        proc = await asyncio.create_subprocess_exec(
            sys.executable, code_file,
            cwd=work_dir,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE
        )
        
        try:
            stdout, stderr = await asyncio.wait_for(
                proc.communicate(),
                timeout=25
            )
        except asyncio.TimeoutError:
            proc.kill()
            await proc.wait()
            return None, "Превышено время выполнения (25 сек)"
        
        if proc.returncode != 0:
            error_text = stderr.decode('utf-8', errors='replace') if stderr else "Неизвестная ошибка"
            return None, f"Ошибка выполнения:\n{error_text}"
        
        pptx_file = os.path.join(work_dir, "presentation.pptx")
        if os.path.exists(pptx_file):
            return pptx_file, None
        else:
            return None, "Файл presentation.pptx не был создан"
            
    except Exception as e:
        return None, str(e)

# ==================== ПРОВЕРКА ПОДПИСКИ (БЕЗОПАСНАЯ) ====================
async def check_sub(message: types.Message) -> bool:
    """Безопасная проверка подписки на канал"""
    try:
        # Если канал не задан, пропускаем проверку
        if not CHANNEL_CHAT_ID or not CHANNEL_USERNAME:
            return True
        
        member = await bot.get_chat_member(chat_id=CHANNEL_CHAT_ID, user_id=message.from_user.id)
        if not member:
            return False
            
        is_subscribed = member.status in ['creator', 'administrator', 'member', 'restricted']
    except Exception as e:
        logger.error(f"Ошибка проверки подписки: {e}")
        # При ошибке пропускаем проверку, чтобы не блокировать пользователей
        return True
    
    if not is_subscribed:
        keyboard = types.InlineKeyboardMarkup(
            inline_keyboard=[
                [types.InlineKeyboardButton(text="✅ Подписаться на канал", url=f"https://t.me/{CHANNEL_USERNAME}")],
                [types.InlineKeyboardButton(text="✅ Я подписался", callback_data="check_subscription")]
            ]
        )
        
        await message.answer(
            f"📢 Для использования бота необходимо подписаться на канал {CHANNEL_USERNAME}\n\n"
            f"1. Нажмите кнопку '✅ Подписаться на канал'\n"
            f"2. После подписки нажмите '✅ Я подписался'",
            reply_markup=keyboard
        )
    
    return is_subscribed

@dp.callback_query(lambda c: c.data == "check_subscription")
async def subscription_callback(callback_query: types.CallbackQuery):
    """Callback проверки подписки"""
    try:
        member = await bot.get_chat_member(chat_id=CHANNEL_CHAT_ID, user_id=callback_query.from_user.id)
        is_subscribed = member.status in ['creator', 'administrator', 'member', 'restricted'] if member else False
    except Exception as e:
        logger.error(f"Ошибка проверки подписки в callback: {e}")
        is_subscribed = False
    
    if is_subscribed:
        await callback_query.message.delete()
        await callback_query.message.answer("✅ Отлично! Теперь вы можете пользоваться ботом!\nОтправьте /start чтобы начать.")
    else:
        await callback_query.answer("❌ Вы ещё не подписались на канал!", show_alert=True)

# ==================== ПРЕМИУМ ФУНКЦИИ ====================
@dp.callback_query(lambda c: c.data == "buy_premium")
async def premium_purchase_callback(callback_query: types.CallbackQuery):
    """Обработка нажатия на кнопку Premium"""
    moderator_id = db_manager.get_contact_moderator()
    
    if not moderator_id:
        moderator_id = ADMIN_IDS[0] if ADMIN_IDS else None
    
    if moderator_id:
        contact_keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="✉️ Связаться", url=f"tg://user?id={moderator_id}")],
                [InlineKeyboardButton(text="◀️ Назад", callback_data="back_to_privileges")]
            ]
        )
    else:
        contact_keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="◀️ Назад", callback_data="back_to_privileges")]
            ]
        )
    
    await callback_query.message.edit_text(
        "💎 Premium навсегда - 100 ₽\n\n"
        "Для покупки Premium доступа напишите модератору\n\n"
        "После оплаты вы получите:\n"
        "✅ Неограниченное количество презентаций\n"
        "✅ Приоритетную обработку запросов\n"
        "✅ Доступ к эксклюзивным шаблонам",
        reply_markup=contact_keyboard
    )
    
    await callback_query.answer()

@dp.callback_query(lambda c: c.data == "back_to_privileges")
async def back_to_privileges_callback(callback_query: types.CallbackQuery):
    """Возврат к списку привилегий"""
    premium_keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="💎 Премиум", callback_data="buy_premium")]
        ]
    )
    
    await callback_query.message.edit_text(
        "✨ Привилегии пользования ботом ✨\n\n"
        "Premium навсегда - 100 ₽\n\n"
        "Преимущества Premium:\n"
        "• 💲 Очень выгодное решение!\n"
        "• 📈 Неограниченное количество презентаций\n"
        "• ⚡ Приоритетная обработка запросов\n"
        "• 🎨 Доступ к эксклюзивным шаблонам\n"
        "• 🚀 Без лимитов и ограничений\n\n"
        "Нажмите на кнопку ниже, чтобы приобрести Premium:",
        reply_markup=premium_keyboard
    )
    
    await callback_query.answer()

# ==================== АДМИН КОМАНДЫ ====================
def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS

def is_admin_or_moderator(user_id: int) -> bool:
    return is_admin(user_id) or db_manager.is_moderator(user_id)

@dp.message(Command("mhelp"))
async def moderator_help(message: Message):
    user_id = message.from_user.id
    if not is_admin_or_moderator(user_id):
        return
    
    help_text = """
🛡️ Команды модератора и администратора

═════════════════════════════

👑 ТОЛЬКО ДЛЯ ГЛАВНЫХ АДМИНОВ:

`/delete <tg_id>` - Удалить пользователя
`/setpremoper <tg_id>` - Установить модератора для связи
`/unpremoper` - Удалить модератора для связи
`/moder <tg_id>` - Назначить модератора
`/unmoder <tg_id>` - Снять модератора
`/clear_cache` - Очистить кэш презентаций
`/users` - Список всех пользователей

═════════════════════════════

🛡️ КОМАНДЫ МОДЕРАТОРОВ:

`/prem <tg_id> [forever|month|year]` - Выдать премиум
`/unprem <tg_id>` - Снять премиум
`/stats` - Полная статистика бота
`/status` - Статус системы (очередь, задачи)
`/moders` - Список модераторов
`/users` - Список всех пользователей

═════════════════════════════

ℹ️ СПРАВКА:

`/mhelp` - Показать эту справку
    """
    await message.answer(help_text, parse_mode="Markdown")

@dp.message(Command("delete"))
async def delete_user(message: Message):
    user_id = message.from_user.id
    if not is_admin(user_id):
        return
    
    args = message.text.split()
    if len(args) < 2:
        await message.answer("❌ Использование: /delete <tg_id>")
        return
    
    try:
        tg_id_to_delete = int(args[1])
        user_stats = db_manager.get_user_stats(tg_id_to_delete)
        if not user_stats:
            await message.answer(f"❌ Пользователь с ID {tg_id_to_delete} не найден")
            return
        
        conn = None
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM presentations_history WHERE tg_id = ?', (tg_id_to_delete,))
            cursor.execute('DELETE FROM premium_users WHERE tg_id = ?', (tg_id_to_delete,))
            cursor.execute('DELETE FROM moderators WHERE tg_id = ?', (tg_id_to_delete,))
            cursor.execute('DELETE FROM users WHERE tg_id = ?', (tg_id_to_delete,))
            conn.commit()
        except Exception as e:
            logger.error(f"Ошибка при удалении из БД: {e}")
            await message.answer(f"❌ Ошибка при удалении: {str(e)}")
            return
        finally:
            if conn:
                conn.close()
        
        await message.answer(f"✅ Пользователь {tg_id_to_delete} удалён")
        try:
            await bot.send_message(tg_id_to_delete, "⚠️ Ваш аккаунт был удалён")
        except:
            pass
    except ValueError:
        await message.answer("❌ Неверный ID пользователя")

@dp.message(Command("prem"))
async def give_premium(message: Message):
    user_id = message.from_user.id
    if not is_admin_or_moderator(user_id):
        return
    
    args = message.text.split()
    if len(args) < 2:
        await message.answer("❌ Использование: /prem <tg_id> [forever|month|year]")
        return
    
    try:
        tg_id = int(args[1])
        duration = args[2] if len(args) > 2 else 'forever'
        
        if duration not in ['forever', 'month', 'year']:
            await message.answer("❌ Тип должен быть: forever, month или year")
            return
        
        if db_manager.add_premium(tg_id, message.from_user.id, duration):
            try:
                await bot.send_message(tg_id, "🎉 Вам выдан Premium доступ!")
            except:
                pass
            await message.answer(f"✅ Premium выдан пользователю {tg_id}")
        else:
            await message.answer("❌ Ошибка выдачи премиум")
    except ValueError:
        await message.answer("❌ Неверный ID пользователя")

@dp.message(Command("unprem"))
async def remove_premium(message: Message):
    user_id = message.from_user.id
    if not is_admin_or_moderator(user_id):
        return
    
    args = message.text.split()
    if len(args) < 2:
        await message.answer("❌ Использование: /unprem <tg_id>")
        return
    
    try:
        tg_id = int(args[1])
        if db_manager.remove_premium(tg_id):
            try:
                await bot.send_message(tg_id, "⚠️ Ваш Premium доступ был отключен")
            except:
                pass
            await message.answer(f"✅ Premium снят с пользователя {tg_id}")
        else:
            await message.answer("❌ Ошибка снятия премиум")
    except ValueError:
        await message.answer("❌ Неверный ID пользователя")

@dp.message(Command("setpremoper"))
async def set_premium_operator(message: Message):
    user_id = message.from_user.id
    if not is_admin(user_id):
        await message.answer("❌ Только для главных администраторов")
        return
    
    args = message.text.split()
    if len(args) < 2:
        await message.answer("❌ Использование: /setpremoper <tg_id>")
        return
    
    try:
        moderator_id = int(args[1])
        if not db_manager.is_moderator(moderator_id) and moderator_id not in ADMIN_IDS:
            await message.answer(f"❌ Пользователь {moderator_id} не является модератором")
            return
        
        if db_manager.set_default_moderator(moderator_id):
            await message.answer(f"✅ Модератор {moderator_id} установлен по умолчанию")
        else:
            await message.answer("❌ Ошибка")
    except ValueError:
        await message.answer("❌ Неверный ID")

@dp.message(Command("unpremoper"))
async def unset_premium_operator(message: Message):
    user_id = message.from_user.id
    if not is_admin(user_id):
        await message.answer("❌ Только для главных администраторов")
        return
    
    if db_manager.remove_default_moderator():
        await message.answer("✅ Модератор по умолчанию удален")
    else:
        await message.answer("❌ Ошибка")

@dp.message(Command("moder"))
async def add_moderator(message: Message):
    user_id = message.from_user.id
    if not is_admin(user_id):
        return
    
    args = message.text.split()
    if len(args) < 2:
        await message.answer("❌ Использование: /moder <tg_id>")
        return
    
    try:
        tg_id = int(args[1])
        db_manager.register_user(tg_id, None, None)
        
        if db_manager.add_moderator(tg_id, message.from_user.id):
            try:
                await bot.send_message(tg_id, "👑 Вам назначена роль модератора!")
            except:
                pass
            await message.answer(f"✅ Пользователь {tg_id} назначен модератором")
        else:
            await message.answer("❌ Ошибка")
    except ValueError:
        await message.answer("❌ Неверный ID")

@dp.message(Command("unmoder"))
async def remove_moderator(message: Message):
    user_id = message.from_user.id
    if not is_admin(user_id):
        return
    
    args = message.text.split()
    if len(args) < 2:
        await message.answer("❌ Использование: /unmoder <tg_id>")
        return
    
    try:
        tg_id = int(args[1])
        if db_manager.remove_moderator(tg_id):
            try:
                await bot.send_message(tg_id, "⚠️ Ваша роль модератора отозвана")
            except:
                pass
            await message.answer(f"✅ Пользователь {tg_id} снят с роли модератора")
        else:
            await message.answer("❌ Ошибка")
    except ValueError:
        await message.answer("❌ Неверный ID")

@dp.message(Command("stats"))
async def admin_stats(message: Message):
    user_id = message.from_user.id
    if not is_admin_or_moderator(user_id):
        return
    
    total_users = db_manager.get_total_users()
    total_presentations = db_manager.get_total_presentations()
    premium_count = db_manager.get_premium_count()
    moderators_count = len(db_manager.get_moderators())
    default_moderator = db_manager.get_default_moderator()
    
    stats_text = (
        f"📊 Полная статистика бота\n\n"
        f"👥 Всего пользователей: {total_users}\n"
        f"📊 Всего презентаций: {total_presentations}\n"
        f"👑 Премиум пользователей: {premium_count}\n"
        f"🛡️ Модераторов: {moderators_count}\n"
        f"⏳ В очереди: {task_queue.qsize()}\n"
        f"🔄 Активных задач: {active_tasks}/{max_concurrent_tasks}\n"
        f"📁 В кэше: {presentation_cache.size()} презентаций\n"
    )
    
    if default_moderator:
        stats_text += f"\n📌 Модератор для связи: `{default_moderator}`"
    
    await message.answer(stats_text, reply_markup=get_main_keyboard())

@dp.message(Command("users"))
async def list_users(message: Message):
    user_id = message.from_user.id
    if not is_admin_or_moderator(user_id):
        return
    
    users = db_manager.get_all_users(limit=50)
    if not users:
        await message.answer("📋 Нет зарегистрированных пользователей")
        return
    
    text = "👥 Последние пользователи:\n\n"
    for user in users[:20]:
        username = user.get('username', 'без юзернейма') or 'без юзернейма'
        text += f"• {user['tg_id']} - {username}\n"
        text += f"  Презентаций: {user['total_presentations']}\n"
    
    await message.answer(text)

@dp.message(Command("moders"))
async def list_moderators(message: Message):
    user_id = message.from_user.id
    if not is_admin_or_moderator(user_id):
        return
    
    moderators = db_manager.get_moderators()
    default_moderator = db_manager.get_default_moderator()
    
    if not moderators:
        await message.answer("📋 Нет назначенных модераторов")
        return
    
    text = "🛡️ Список модераторов:\n\n"
    for mod in moderators:
        username = mod.get('username', 'без юзернейма') or 'без юзернейма'
        marker = " ⭐ (по умолчанию)" if mod['tg_id'] == default_moderator else ""
        text += f"• {mod['tg_id']} - {username}{marker}\n"
    
    await message.answer(text)

@dp.message(Command("status"))
async def cmd_status(message: Message):
    user_id = message.from_user.id
    
    if not is_admin_or_moderator(user_id):
        await message.answer(
            f"📊 Статус системы\n"
            f"⏳ В очереди: {task_queue.qsize()}\n"
            f"🔄 Активных задач: {active_tasks}/{max_concurrent_tasks}",
            reply_markup=get_main_keyboard()
        )
    else:
        total_users = db_manager.get_total_users()
        total_presentations = db_manager.get_total_presentations()
        await message.answer(
            f"📊 Статус системы\n"
            f"⏳ В очереди: {task_queue.qsize()}\n"
            f"🔄 Активных задач: {active_tasks}/{max_concurrent_tasks}\n"
            f"👥 Пользователей: {total_users}\n"
            f"📊 Презентаций: {total_presentations}\n"
            f"📁 В кэше: {presentation_cache.size()} презентаций",
            reply_markup=get_main_keyboard()
        )

@dp.message(Command("clear_cache"))
async def cmd_clear_cache(message: Message):
    user_id = message.from_user.id
    if not is_admin(user_id):
        return
    
    presentation_cache.clear()
    await message.answer("🗑️ Кэш очищен", reply_markup=get_main_keyboard())

# ==================== ОСНОВНАЯ ЛОГИКА ОБРАБОТКИ (С ТАЙМАУТОМ И ПРЕДОХРАНИТЕЛЯМИ) ====================
async def process_presentation_request(user_id, request, message):
    """Обработка запроса на создание презентации с общим таймаутом и защитой от ошибок"""
    task_completed = False
    try:
        # Запускаем задачу с таймаутом
        try:
            await asyncio.wait_for(
                _process_presentation_request_internal(user_id, request, message),
                timeout=90  # Общий таймаут 90 секунд
            )
            task_completed = True
        except asyncio.TimeoutError:
            logger.error(f"Timeout для пользователя {user_id} (90 секунд)")
            try:
                await message.answer(
                    "❌ Время выполнения превышено (90 секунд).\n\n"
                    "Попробуйте упростить запрос или повторить позже.\n"
                    "💎 Премиум пользователи имеют приоритетную обработку."
                )
            except Exception as e:
                logger.error(f"Не удалось отправить сообщение о таймауте: {e}")
        except asyncio.CancelledError:
            logger.warning(f"Задача для пользователя {user_id} была отменена")
            try:
                await message.answer("❌ Задача была отменена. Попробуйте ещё раз.")
            except:
                pass
        except Exception as e:
            logger.error(f"Критическая ошибка в process_presentation_request: {traceback.format_exc()}")
            try:
                await message.answer("❌ Произошла непредвиденная ошибка. Попробуйте позже.")
            except:
                pass
        
    finally:
        # ВСЕГДА уменьшаем счетчик активных задач, даже при ошибке
        async with active_tasks_lock:
            global active_tasks
            if active_tasks > 0:
                active_tasks -= 1
            logger.info(f"Задача завершена. Активных задач: {active_tasks}")
            
            # Берем следующую задачу из очереди, если есть
            if not task_queue.empty():
                try:
                    next_user_id, next_request, next_msg = await task_queue.get()
                    active_tasks += 1
                    logger.info(f"Запускаем следующую задачу из очереди для пользователя {next_user_id}")
                    asyncio.create_task(process_presentation_request(next_user_id, next_request, next_msg))
                except Exception as e:
                    logger.error(f"Ошибка при запуске следующей задачи из очереди: {e}")

async def _process_presentation_request_internal(user_id, request, message):
    """Внутренняя реализация обработки запроса"""
    work_dir = None
    try:
        # Проверяем лимиты
        try:
            can_generate, limit_message = db_manager.can_generate_presentation(user_id)
        except Exception as e:
            logger.error(f"Ошибка проверки лимитов: {e}")
            await message.answer("❌ Ошибка проверки лимитов. Попробуйте позже.")
            return
        
        if not can_generate:
            await message.answer(
                f"⚠️ Лимит генераций исчерпан!\n\n"
                f"Вы можете создать до 3 презентаций каждые 5 часов.\n"
                f"⏰ {limit_message}\n\n"
                f"💎 Приобретите Premium для снятия лимитов: /start → 👑 Привилегии"
            )
            return
        
        # Проверяем кэш
        try:
            cached_path = presentation_cache.get(request)
            if cached_path and os.path.exists(cached_path):
                await message.answer_document(
                    FSInputFile(cached_path),
                    caption=(
                        "✅ Презентация из кэша (уже создавалась ранее)\n\n"
                        "📝 Введите следующую тему или нажмите «❌ Отмена» для возврата в меню."
                    )
                )
                db_manager.increment_presentations(user_id)
                db_manager.add_presentation_history(user_id, request)
                # Оставляем пользователя в режиме ввода
                user_states[user_id] = "waiting_for_prompt"
                return
        except Exception as e:
            logger.error(f"Ошибка при проверке кэша: {e}")
            # Продолжаем выполнение, если кэш недоступен
        
        # Шаг 1: Генерация кода через Mistral
        status_msg = await message.answer("🔄 Генерирую код презентации...")
        
        def sync_mistral_call():
            return mistral_client.chat.complete(
                model="mistral-large-latest",
                messages=[
                    {"role": "system", "content": INSTRUCTION},
                    {"role": "user", "content": request}
                ]
            )
        
        try:
            chat_response = await asyncio.get_running_loop().run_in_executor(
                executor,
                sync_mistral_call
            )
            await status_msg.delete()
        except Exception as e:
            logger.error(f"Ошибка Mistral API: {traceback.format_exc()}")
            await status_msg.edit_text("❌ Ошибка при обращении к ИИ. Попробуйте позже.")
            return
        
        generated_code = chat_response.choices[0].message.content
        python_code = extract_python_code(generated_code)
        
        if not python_code:
            await message.answer("❌ Не удалось извлечь код из ответа ИИ")
            return
        
        # Создаём временную папку
        work_dir = os.path.join(TEMP_DIR, f"{user_id}_{uuid.uuid4().hex}")
        os.makedirs(work_dir, exist_ok=True)
        
        # Шаг 2: Выполнение кода (АСИНХРОННО)
        await message.answer("⚙️ Выполняю код и создаю презентацию...")
        
        async with subprocess_semaphore:
            result_file, error = await safe_execute_code_async(python_code, work_dir)
        
        if error:
            logger.error(f"Ошибка выполнения кода для пользователя {user_id}: {error}")
            # Сохраняем данные для возможного retry
            retry_data[user_id] = {
                "original_request": request,
                "error_text": error
            }
            retry_keyboard = InlineKeyboardMarkup(
                inline_keyboard=[
                    [InlineKeyboardButton(text="🔧 Исправить", callback_data=f"retry_fix:{user_id}")]
                ]
            )
            await message.answer(
                "❌ Произошла ошибка при создании презентации.\n\n"
                "Нажмите «🔧 Исправить», чтобы ИИ попытался исправить код автоматически.",
                reply_markup=retry_keyboard
            )
            # Пользователь остаётся в waiting_for_prompt — может ввести новый промпт сам
            user_states[user_id] = "waiting_for_prompt"
            return
        
        # Шаг 3: Сохраняем файл
        final_filename = f"{user_id}_{uuid.uuid4().hex}.pptx"
        final_path = os.path.join(COMPLETED_DIR, final_filename)
        
        try:
            move_file_with_retry(result_file, final_path)
        except Exception as e:
            logger.error(f"Ошибка перемещения файла: {e}")
            await message.answer("❌ Ошибка при сохранении презентации")
            return
        
        # Добавляем в кэш
        try:
            presentation_cache.add(request, final_path)
        except Exception as e:
            logger.error(f"Ошибка добавления в кэш: {e}")
        
        # Увеличиваем счетчик
        try:
            db_manager.increment_presentations(user_id)
            db_manager.add_presentation_history(user_id, request)
        except Exception as e:
            logger.error(f"Ошибка обновления статистики: {e}")
        
        # Отправляем пользователю
        await message.answer_document(
            FSInputFile(final_path),
            caption=(
                "✅ Ваша презентация готова!\n\n"
                "📝 Введите следующую тему или нажмите «❌ Отмена» для возврата в меню."
            )
        )
        # Оставляем пользователя в режиме ввода — не выбрасываем в главное меню
        user_states[user_id] = "waiting_for_prompt"
        
    except Exception as e:
        logger.error(f"Непредвиденная ошибка в _process_presentation_request_internal: {traceback.format_exc()}")
        try:
            await message.answer(f"❌ Произошла непредвиденная ошибка. Пожалуйста, попробуйте позже.")
        except:
            pass
    finally:
        # Очистка временной папки (отложенная)
        if work_dir and os.path.exists(work_dir):
            async def delayed_cleanup():
                await asyncio.sleep(300)
                try:
                    shutil.rmtree(work_dir, ignore_errors=True)
                    logger.info(f"Очищена временная папка: {work_dir}")
                except Exception as e:
                    logger.error(f"Ошибка при очистке {work_dir}: {e}")
            
            asyncio.create_task(delayed_cleanup())

# ==================== ХЕНДЛЕРЫ ====================
@dp.message(Command("start"))
async def cmd_start(message: Message):
    user_id = message.from_user.id
    user_states[user_id] = None
    
    try:
        db_manager.register_user(
            user_id, 
            message.from_user.username,
            message.from_user.first_name
        )
        db_manager.update_activity(user_id)
    except Exception as e:
        logger.error(f"Ошибка регистрации пользователя: {e}")
    
    await message.answer(
        "👋 Привет! Я бот для создания презентаций с помощью ИИ.\n\n"
        "📌 Нажми на кнопку ➕ Создать, чтобы начать создание презентации\n"
        "👑 Нажми на кнопку Привилегии, чтобы узнать о Premium доступе\n"
        "👤 Нажми на кнопку Мой профиль, чтобы узнать свой статус\n\n"
        "🔧 Доступные команды:\n"
        "/start - главное меню",
        reply_markup=get_main_keyboard()
    )

@dp.message(lambda message: message.text == "➕ Создать")
async def create_presentation_button(message: Message):
    user_id = message.from_user.id
    
    # Анти-спам
    can_click, wait_seconds = check_spam(user_id)
    if not can_click:
        await message.answer(
            f"⏳ Пожалуйста, подождите {wait_seconds} секунд перед следующим запросом.",
            reply_markup=get_main_keyboard()
        )
        return
    
    if not await check_sub(message):
        return
    
    user_states[user_id] = "waiting_for_prompt"
    db_manager.update_activity(user_id)
    
    can_generate, limit_message = db_manager.can_generate_presentation(user_id)
    
    if can_generate:
        if "premium" not in limit_message:
            await message.answer(
                f"📝 Введите тему презентации\n\n"
                f"ℹ️ У вас осталось {limit_message}\n\n"
                f"Опишите, какую презентацию вы хотите получить.\n"
                f"Например:\n"
                f"Создай презентацию 'Биография Пушкина' на 4 слайда\n"
                f"❌ Нажмите 'Отмена', чтобы вернуться",
                reply_markup=get_cancel_keyboard()
            )
        else:
            await message.answer(
                "📝 Введите тему презентации\n\n"
                "Опишите, какую презентацию вы хотите получить.\n"
                "❌ Нажмите 'Отмена', чтобы вернуться",
                reply_markup=get_cancel_keyboard()
            )
    else:
        await message.answer(
            f"⚠️ Лимит генераций исчерпан!\n\n"
            f"⏰ {limit_message}\n\n"
            f"💎 Приобретите Premium: /start → 👑 Привилегии",
            reply_markup=get_main_keyboard()
        )
        user_states[user_id] = None

@dp.message(lambda message: message.text == "👑 Привилегии")
async def privileges_button(message: Message):
    premium_keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="💎 Премиум", callback_data="buy_premium")]
        ]
    )
    
    await message.answer(
        "✨ Привилегии пользования ботом ✨\n\n"
        "Premium навсегда - 100 ₽\n\n"
        "Преимущества Premium:\n"
        "• 💲 Очень выгодное решение!\n"
        "• 📈 Неограниченное количество презентаций\n"
        "• ⚡ Приоритетная обработка запросов\n"
        "• 🚀 Без лимитов и ограничений\n\n"
        "Нажмите на кнопку ниже, чтобы приобрести Premium:",
        reply_markup=premium_keyboard
    )

@dp.message(lambda message: message.text == "👤 Мой профиль")
async def my_profile_button(message: Message):
    user_id = message.from_user.id
    db_manager.update_activity(user_id)
    
    user_stats = db_manager.get_user_stats(user_id)
    is_premium = db_manager.is_premium(user_id)
    is_moder = db_manager.is_moderator(user_id)
    is_admin_user = is_admin(user_id)
    
    presentations_last_5h = db_manager.get_presentations_count_since(user_id, 5)
    remaining = 3 - presentations_last_5h
    
    profile_text = f"👤 Ваш профиль\n\n"
    profile_text += f"🆔 ID: {user_id}\n"
    
    if user_stats:
        username = user_stats.get('username')
        if username:
            profile_text += f"📝 Username: @{username}\n"
        registered_at = user_stats.get('registered_at', 'неизвестно')
        profile_text += f"📅 В боте с: {registered_at[:10] if registered_at else 'неизвестно'}\n"
        profile_text += f"📊 Создано презентаций: {user_stats.get('total_presentations', 0)}\n"
    
    if is_admin_user:
        profile_text += f"\n👑 Статус: ГЛАВНЫЙ АДМИНИСТРАТОР\n⭐ Премиум: ДА (бессрочно)\n🎯 Лимиты: отсутствуют\n"
    elif is_moder:
        profile_text += f"\n🛡️ Статус: МОДЕРАТОР\n⭐ Премиум: ДА (бесплатно)\n🎯 Лимиты: отсутствуют\n"
    elif is_premium:
        profile_text += f"\n👑 Статус: PREMIUM\n🎯 Лимиты: отсутствуют\n✨ Спасибо за поддержку!"
    else:
        profile_text += f"\n⭐ Статус: Обычный пользователь\n\n🎯 Лимит: {max(0, remaining)} из 3 презентаций\n"
        if remaining <= 0:
            reset_time = db_manager.get_next_reset_time(user_id, 5)
            if reset_time:
                wait_minutes = int((reset_time - datetime.now()).total_seconds() / 60)
                hours_left = wait_minutes // 60
                minutes_left = wait_minutes % 60
                if hours_left > 0:
                    profile_text += f"⏰ Лимит обновится через: {hours_left} ч {minutes_left} мин\n"
                else:
                    profile_text += f"⏰ Лимит обновится через: {minutes_left} мин\n"
        profile_text += f"\n💰 Приобретите Premium за 100 ₽\nНажмите '👑 Привилегии'"
    
    await message.answer(profile_text, reply_markup=get_main_keyboard())

@dp.message(lambda message: message.text == "❌ Отмена")
async def cancel_button(message: Message):
    user_id = message.from_user.id
    user_states[user_id] = None
    await message.answer("✅ Действие отменено", reply_markup=get_main_keyboard())

@dp.message()
async def handle_presentation_prompt(message: Message):
    user_id = message.from_user.id
    
    if user_states.get(user_id) == "waiting_for_prompt":
        await process_presentation_request_from_message(message)
    else:
        await message.answer(
            "❓ Чтобы создать презентацию, нажмите кнопку ➕ Создать",
            reply_markup=get_main_keyboard()
        )

async def process_presentation_request_from_message(message: Message):
    global active_tasks
    
    user_id = message.from_user.id
    request = message.text.strip()
    # Не сбрасываем user_states[user_id] здесь — состояние управляется внутри генерации
    
    if len(request) > 500:
        await message.answer(
            "❌ Слишком длинный запрос. Максимум 500 символов.\n\n"
            "Попробуйте сформулировать тему короче или нажмите «❌ Отмена».",
            reply_markup=get_cancel_keyboard()
        )
        return
    
    if len(request) < 5:
        await message.answer(
            "❌ Слишком короткий запрос. Опишите тему подробнее.\n\n"
            "Или нажмите «❌ Отмена» для возврата в меню.",
            reply_markup=get_cancel_keyboard()
        )
        return
    
    # Проверка лимитов
    can_generate, limit_message = db_manager.can_generate_presentation(user_id)
    if not can_generate:
        await message.answer(
            f"⚠️ Лимит генераций исчерпан!\n\n⏰ {limit_message}\n\n💎 Приобретите Premium",
            reply_markup=get_main_keyboard()
        )
        user_states[user_id] = None
        return
    
    status_msg = await message.answer(
        f"⏳ Запрос принят. Создаю презентацию на тему: {request[:50]}...\n"
        f"Позиция в очереди: {task_queue.qsize() + 1}\n"
        f"Активных задач: {active_tasks}/{max_concurrent_tasks}"
    )
    
    async with active_tasks_lock:
        if active_tasks >= max_concurrent_tasks:
            await task_queue.put((user_id, request, message))
            await status_msg.edit_text(
                f"⏳ Ваш запрос добавлен в очередь.\n"
                f"Позиция: {task_queue.qsize()}\n"
                f"Активных задач: {active_tasks}/{max_concurrent_tasks}"
            )
        else:
            active_tasks += 1
            asyncio.create_task(process_presentation_request(user_id, request, message))
    
    await asyncio.sleep(3)
    try:
        await status_msg.delete()
    except:
        pass

# ==================== RETRY (ИСПРАВИТЬ КОД) ====================
@dp.callback_query(lambda c: c.data and c.data.startswith("retry_fix:"))
async def retry_fix_callback(callback_query: types.CallbackQuery):
    """
    Обработчик кнопки 'Исправить'.
    Берёт оригинальный запрос и текст ошибки из retry_data,
    формирует расширенный промпт и запускает повторную генерацию.
    """
    global active_tasks

    # Достаём user_id из callback_data (формат "retry_fix:<user_id>")
    try:
        target_user_id = int(callback_query.data.split(":")[1])
    except (IndexError, ValueError):
        await callback_query.answer("❌ Неверные данные для исправления.", show_alert=True)
        return

    caller_id = callback_query.from_user.id

    # Разрешаем нажимать только самому пользователю
    if caller_id != target_user_id:
        await callback_query.answer("❌ Эта кнопка не для вас.", show_alert=True)
        return

    data = retry_data.get(target_user_id)
    if not data:
        await callback_query.answer("❌ Данные для исправления устарели. Попробуйте создать презентацию заново.", show_alert=True)
        return

    original_request = data["original_request"]
    error_text = data["error_text"]

    # Удаляем устаревшие retry-данные
    retry_data.pop(target_user_id, None)

    # Убираем кнопку из сообщения об ошибке
    try:
        await callback_query.message.edit_reply_markup(reply_markup=None)
    except Exception:
        pass

    await callback_query.answer()

    # Формируем расширенный промпт с текстом ошибки
    retry_request = (
        f"{original_request}\n\n"
        f"[ВАЖНО: предыдущая попытка завершилась ошибкой — исправь её]\n"
        f"Ошибка выполнения:\n{error_text[:800]}"
    )

    # Проверяем лимиты
    can_generate, limit_message = db_manager.can_generate_presentation(target_user_id)
    if not can_generate:
        await callback_query.message.answer(
            f"⚠️ Лимит генераций исчерпан!\n\n⏰ {limit_message}\n\n💎 Приобретите Premium",
            reply_markup=get_main_keyboard()
        )
        user_states[target_user_id] = None
        return

    await callback_query.message.answer(
        f"🔧 Запускаю исправление...\n"
        f"Тема: {original_request[:60]}{'...' if len(original_request) > 60 else ''}"
    )

    async with active_tasks_lock:
        if active_tasks >= max_concurrent_tasks:
            await task_queue.put((target_user_id, retry_request, callback_query.message))
            await callback_query.message.answer(
                f"⏳ Задача на исправление добавлена в очередь.\n"
                f"Позиция: {task_queue.qsize()}\n"
                f"Активных задач: {active_tasks}/{max_concurrent_tasks}"
            )
        else:
            active_tasks += 1
            asyncio.create_task(
                process_presentation_request(target_user_id, retry_request, callback_query.message)
            )


# ==================== GRACEFUL SHUTDOWN ====================
async def shutdown():
    """Корректное завершение работы бота"""
    global shutdown_flag
    shutdown_flag = True
    logger.info("Завершение работы бота...")
    
    # Закрываем executor
    try:
        executor.shutdown(wait=True)
    except Exception as e:
        logger.error(f"Ошибка при закрытии executor: {e}")
    
    # Закрываем сессию бота
    try:
        await bot.session.close()
    except Exception as e:
        logger.error(f"Ошибка при закрытии сессии бота: {e}")
    
    logger.info("Бот остановлен")

def signal_handler():
    """Обработчик сигналов для корректного завершения"""
    loop = asyncio.get_running_loop()
    if loop.is_running():
        loop.create_task(shutdown())
    else:
        asyncio.run(shutdown())

# ==================== ЗАПУСК ====================
async def main():
    # Проверка зависимостей
    if not check_dependencies():
        logger.error("Критическая ошибка: отсутствуют необходимые зависимости")
        logger.error("Установите: pip install python-pptx")
        return
    
    # Очистка старых временных папок
    cleanup_old_temp_dirs()
    
    logger.info("🤖 Бот запущен на ПК...")
    logger.info(f"ОС: {sys.platform}")
    logger.info(f"Python: {sys.version}")
    logger.info(f"Максимум задач: {max_concurrent_tasks}")
    logger.info(f"Тип кэша: {CACHE_TYPE.upper()}")
    logger.info(f"Администраторы: {ADMIN_IDS}")
    logger.info(f"Subprocess: asyncio.create_subprocess_exec")
    
    # Настройка обработчиков сигналов
    for sig in [signal.SIGTERM, signal.SIGINT]:
        try:
            asyncio.get_running_loop().add_signal_handler(sig, signal_handler)
        except NotImplementedError:
            # Windows не поддерживает signal handlers в asyncio
            pass
    
    try:
        await dp.start_polling(bot)
    finally:
        await shutdown()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Бот остановлен пользователем")
    except Exception as e:
        logger.error(f"Критическая ошибка: {traceback.format_exc()}")
