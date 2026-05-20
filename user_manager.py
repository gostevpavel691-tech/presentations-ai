import sqlite3
import logging
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict, List
import config

logger = logging.getLogger(__name__)

class DatabaseManager:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._init_db()
    
    def _init_db(self):
        """Инициализация базы данных"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Таблица всех пользователей
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    tg_id INTEGER PRIMARY KEY,
                    username TEXT,
                    first_name TEXT,
                    registered_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    total_presentations INTEGER DEFAULT 0,
                    last_activity TIMESTAMP,
                    last_presentation TIMESTAMP
                )
            ''')
            
            # Таблица модераторов
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS moderators (
                    tg_id INTEGER PRIMARY KEY,
                    added_by INTEGER,
                    added_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (tg_id) REFERENCES users (tg_id)
                )
            ''')
            
            # Таблица премиум пользователей
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS premium_users (
                    tg_id INTEGER PRIMARY KEY,
                    premium_start TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    premium_end TIMESTAMP,
                    premium_type TEXT DEFAULT 'forever',
                    given_by INTEGER,
                    given_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (tg_id) REFERENCES users (tg_id)
                )
            ''')
            
            # Таблица истории презентаций
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS presentations_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tg_id INTEGER,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    prompt TEXT,
                    FOREIGN KEY (tg_id) REFERENCES users (tg_id)
                )
            ''')
            
            # Таблица для настройки модератора по умолчанию для связи
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS bot_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            ''')
            
            conn.commit()
            logger.info("База данных инициализирована")
        except Exception as e:
            logger.error(f"Ошибка инициализации базы данных: {e}")
        finally:
            if conn:
                conn.close()
    
    # ============ РАБОТА С ПОЛЬЗОВАТЕЛЯМИ ============
    def register_user(self, tg_id: int, username: str = None, first_name: str = None):
        """Регистрация нового пользователя"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO users (tg_id, username, first_name, last_activity)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(tg_id) DO UPDATE SET
                    username = excluded.username,
                    first_name = excluded.first_name,
                    last_activity = CURRENT_TIMESTAMP
            ''', (tg_id, username, first_name))
            conn.commit()
            return True
        except Exception as e:
            logger.error(f"Ошибка регистрации пользователя {tg_id}: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def update_activity(self, tg_id: int):
        """Обновить время последней активности"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE users 
                SET last_activity = CURRENT_TIMESTAMP
                WHERE tg_id = ?
            ''', (tg_id,))
            conn.commit()
        except Exception as e:
            logger.error(f"Ошибка обновления активности: {e}")
        finally:
            if conn:
                conn.close()
    
    def increment_presentations(self, tg_id: int):
        """Увеличить счетчик созданных презентаций"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE users 
                SET total_presentations = total_presentations + 1,
                    last_activity = CURRENT_TIMESTAMP,
                    last_presentation = CURRENT_TIMESTAMP
                WHERE tg_id = ?
            ''', (tg_id,))
            conn.commit()
        except Exception as e:
            logger.error(f"Ошибка увеличения счетчика презентаций: {e}")
        finally:
            if conn:
                conn.close()
    
    def get_user_stats(self, tg_id: int) -> Optional[Dict]:
        """Получить статистику пользователя"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('''
                SELECT tg_id, username, first_name, 
                       registered_at, total_presentations, last_activity, last_presentation
                FROM users 
                WHERE tg_id = ?
            ''', (tg_id,))
            result = cursor.fetchone()
            return dict(result) if result else None
        except Exception as e:
            logger.error(f"Ошибка получения статистики пользователя: {e}")
            return None
        finally:
            if conn:
                conn.close()
    
    def get_all_users(self, limit: int = 100) -> List[Dict]:
        """Получить список всех пользователей"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('''
                SELECT tg_id, username, first_name, total_presentations, registered_at, last_activity
                FROM users 
                ORDER BY registered_at DESC
                LIMIT ?
            ''', (limit,))
            results = cursor.fetchall()
            return [dict(row) for row in results]
        except Exception as e:
            logger.error(f"Ошибка получения списка пользователей: {e}")
            return []
        finally:
            if conn:
                conn.close()
    
    def get_tg_id_by_username(self, username: str) -> Optional[int]:
        """Найти tg_id пользователя по username (без @)"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute(
                'SELECT tg_id FROM users WHERE LOWER(username) = LOWER(?)',
                (username.lstrip('@'),)
            )
            result = cursor.fetchone()
            return result[0] if result else None
        except Exception as e:
            logger.error(f"Ошибка поиска пользователя по username: {e}")
            return None
        finally:
            if conn:
                conn.close()

    def get_total_users(self) -> int:
        """Получить общее количество пользователей"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM users')
            count = cursor.fetchone()[0]
            return count
        except Exception as e:
            logger.error(f"Ошибка подсчета пользователей: {e}")
            return 0
        finally:
            if conn:
                conn.close()
    
    def get_total_presentations(self) -> int:
        """Получить общее количество созданных презентаций"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT SUM(total_presentations) FROM users')
            total = cursor.fetchone()[0]
            return total or 0
        except Exception as e:
            logger.error(f"Ошибка подсчета презентаций: {e}")
            return 0
        finally:
            if conn:
                conn.close()
    
    # ============ ЛИМИТЫ ГЕНЕРАЦИЙ ============
    def get_presentations_count_since(self, tg_id: int, hours: int = 5) -> int:
        """Получить количество презентаций, созданных за последние N часов"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                SELECT COUNT(*) FROM presentations_history 
                WHERE tg_id = ? AND created_at > datetime('now', ?)
            ''', (tg_id, f'-{hours} hours'))
            count = cursor.fetchone()[0]
            return count
        except Exception as e:
            logger.error(f"Ошибка подсчета презентаций за период: {e}")
            return 0
        finally:
            if conn:
                conn.close()
    
    def add_presentation_history(self, tg_id: int, prompt: str = None):
        """Добавить запись о создании презентации в историю"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO presentations_history (tg_id, prompt, created_at)
                VALUES (?, ?, CURRENT_TIMESTAMP)
            ''', (tg_id, prompt))
            conn.commit()
            return True
        except Exception as e:
            logger.error(f"Ошибка добавления истории презентации: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def get_next_reset_time(self, tg_id: int, hours: int = 5) -> Optional[datetime]:
        """Получить время, когда обновится лимит"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                SELECT created_at FROM presentations_history 
                WHERE tg_id = ? 
                ORDER BY created_at DESC 
                LIMIT 1
            ''', (tg_id,))
            result = cursor.fetchone()
            
            if result:
                oldest_time = datetime.fromisoformat(result[0]).replace(tzinfo=timezone.utc)
                reset_time = oldest_time + timedelta(hours=hours)
                return reset_time
            return None
        except Exception as e:
            logger.error(f"Ошибка получения времени сброса: {e}")
            return None
        finally:
            if conn:
                conn.close()
    
    def can_generate_presentation(self, tg_id: int, max_per_hours: int = 2, hours: int = 5) -> tuple[bool, str]:
        """Проверить, может ли пользователь генерировать презентацию"""
        # Премиум пользователи, модераторы и админы не имеют лимитов
        if self.is_premium(tg_id):
            return True, "premium"
        
        count = self.get_presentations_count_since(tg_id, hours)
        
        if count < max_per_hours:
            return True, f"осталось {max_per_hours - count} из {max_per_hours}"
        else:
            reset_time = self.get_next_reset_time(tg_id, hours)
            if reset_time:
                wait_minutes = int((reset_time - datetime.now(timezone.utc)).total_seconds() / 60)
                hours_left = wait_minutes // 60
                minutes_left = wait_minutes % 60
                if hours_left > 0:
                    wait_text = f"{hours_left} ч {minutes_left} мин"
                else:
                    wait_text = f"{minutes_left} мин"
                return False, f"Лимит исчерпан. Следующая генерация через {wait_text}"
            return False, "Лимит исчерпан"
    
    # ============ РАБОТА С МОДЕРАТОРАМИ ============
    def add_moderator(self, tg_id: int, admin_id: int) -> bool:
        """Назначить модератора"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO moderators (tg_id, added_by)
                VALUES (?, ?)
            ''', (tg_id, admin_id))
            conn.commit()
            logger.info(f"Админ {admin_id} назначил модератора {tg_id}")
            return True
        except Exception as e:
            logger.error(f"Ошибка назначения модератора: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def remove_moderator(self, tg_id: int) -> bool:
        """Снять модератора"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM moderators WHERE tg_id = ?', (tg_id,))
            conn.commit()
            logger.info(f"Снят модератор {tg_id}")
            return True
        except Exception as e:
            logger.error(f"Ошибка снятия модератора: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def is_moderator(self, tg_id: int) -> bool:
        """Проверить, является ли пользователь модератором"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT 1 FROM moderators WHERE tg_id = ?', (tg_id,))
            result = cursor.fetchone()
            return result is not None
        except Exception as e:
            logger.error(f"Ошибка проверки модератора: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def get_moderators(self) -> List[Dict]:
        """Получить список всех модераторов"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('''
                SELECT u.tg_id, u.username, u.first_name, m.added_at
                FROM moderators m
                JOIN users u ON u.tg_id = m.tg_id
                ORDER BY m.added_at DESC
            ''')
            results = cursor.fetchall()
            return [dict(row) for row in results]
        except Exception as e:
            logger.error(f"Ошибка получения списка модераторов: {e}")
            return []
        finally:
            if conn:
                conn.close()
    
    # ============ НАСТРОЙКИ БОТА ============
    def set_default_moderator(self, moderator_id: int) -> bool:
        """Установить модератора по умолчанию для связи"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO bot_settings (key, value)
                VALUES ('default_moderator', ?)
            ''', (str(moderator_id),))
            conn.commit()
            logger.info(f"Установлен модератор по умолчанию: {moderator_id}")
            return True
        except Exception as e:
            logger.error(f"Ошибка установки модератора по умолчанию: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def remove_default_moderator(self) -> bool:
        """Удалить модератора по умолчанию (будут выбираться рандомные)"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM bot_settings WHERE key = "default_moderator"')
            conn.commit()
            logger.info("Удален модератор по умолчанию")
            return True
        except Exception as e:
            logger.error(f"Ошибка удаления модератора по умолчанию: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def get_default_moderator(self) -> Optional[int]:
        """Получить ID модератора по умолчанию"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT value FROM bot_settings WHERE key = "default_moderator"')
            result = cursor.fetchone()
            if result:
                return int(result[0])
            return None
        except Exception as e:
            logger.error(f"Ошибка получения модератора по умолчанию: {e}")
            return None
        finally:
            if conn:
                conn.close()
    
    def get_contact_moderator(self) -> Optional[int]:
        """Получить ID модератора для связи (приоритет: заданный -> рандомный)"""
        default_mod = self.get_default_moderator()
        
        if default_mod:
            # Проверяем, существует ли еще этот модератор
            if self.is_moderator(default_mod):
                return default_mod
        
        # Если нет заданного, берем рандомного из списка
        moderators = self.get_moderators()
        if moderators:
            import random
            return random.choice(moderators)['tg_id']
        
        return None
    
    # ============ РАБОТА С ПРЕМИУМ ============
    def add_premium(self, tg_id: int, admin_id: int, duration: str = 'forever') -> bool:
        """Выдать премиум пользователю"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            now = datetime.now(timezone.utc)
            end_date = None
            
            if duration == 'month':
                end_date = now + timedelta(days=30)
            elif duration == 'year':
                end_date = now + timedelta(days=365)
            
            cursor.execute('''
                INSERT OR REPLACE INTO premium_users 
                (tg_id, premium_start, premium_end, premium_type, given_by)
                VALUES (?, ?, ?, ?, ?)
            ''', (tg_id, now, end_date, duration, admin_id))
            
            conn.commit()
            logger.info(f"Админ {admin_id} выдал премиум пользователю {tg_id} (тип: {duration})")
            return True
        except Exception as e:
            logger.error(f"Ошибка выдачи премиум: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def remove_premium(self, tg_id: int) -> bool:
        """Снять премиум с пользователя"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM premium_users WHERE tg_id = ?', (tg_id,))
            conn.commit()
            logger.info(f"Снят премиум с пользователя {tg_id}")
            return True
        except Exception as e:
            logger.error(f"Ошибка снятия премиум: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def is_premium(self, tg_id: int) -> bool:
        """Проверить, есть ли у пользователя активный премиум"""
        # Модераторы и админы имеют бесплатный премиум
        if self.is_moderator(tg_id) or tg_id in config.ADMIN_IDS:
            return True
        
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                SELECT premium_end FROM premium_users WHERE tg_id = ?
            ''', (tg_id,))
            result = cursor.fetchone()
            
            if not result:
                return False
            
            end_date = result[0]
            if end_date:
                try:
                    end_date_obj = datetime.fromisoformat(end_date).replace(tzinfo=timezone.utc)
                    if datetime.now(timezone.utc) > end_date_obj:
                        self.remove_premium(tg_id)
                        return False
                except:
                    pass
            
            return True
        except Exception as e:
            logger.error(f"Ошибка проверки премиум: {e}")
            return False
        finally:
            if conn:
                conn.close()
    
    def get_premium_info(self, tg_id: int) -> Optional[Dict]:
        """Получить информацию о премиум статусе пользователя"""
        # Модераторы и админы имеют вечный премиум
        if self.is_moderator(tg_id) or tg_id in config.ADMIN_IDS:
            return {
                'premium_type': 'moderator',
                'premium_end': None,
                'premium_start': datetime.now(timezone.utc).isoformat(),
                'given_at': datetime.now(timezone.utc).isoformat()
            }
        
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('''
                SELECT premium_start, premium_end, premium_type, given_at
                FROM premium_users WHERE tg_id = ?
            ''', (tg_id,))
            result = cursor.fetchone()
            return dict(result) if result else None
        except Exception as e:
            logger.error(f"Ошибка получения информации о премиум: {e}")
            return None
        finally:
            if conn:
                conn.close()
    
    def get_premium_users(self) -> List[Dict]:
        """Получить список всех премиум пользователей"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('''
                SELECT u.tg_id, u.username, u.first_name, u.total_presentations, 
                       p.premium_type, p.premium_end, p.given_at
                FROM premium_users p
                JOIN users u ON u.tg_id = p.tg_id
                ORDER BY p.given_at DESC
            ''')
            results = cursor.fetchall()
            return [dict(row) for row in results]
        except Exception as e:
            logger.error(f"Ошибка получения списка премиум: {e}")
            return []
        finally:
            if conn:
                conn.close()

    def get_users_page(self, page: int = 1, per_page: int = 10) -> List[Dict]:
        """Возвращает список пользователей для указанной страницы."""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            offset = (page - 1) * per_page
            cursor.execute('''
                SELECT tg_id, username, first_name, total_presentations, registered_at, last_activity
                FROM users 
                ORDER BY registered_at DESC
                LIMIT ? OFFSET ?
            ''', (per_page, offset))
            results = cursor.fetchall()
            return [dict(row) for row in results]
        except Exception as e:
            logger.error(f"Ошибка получения страницы пользователей: {e}")
            return []
        finally:
            if conn:
                conn.close()
    
    def get_premium_count(self) -> int:
        """Получить количество премиум пользователей"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM premium_users')
            count = cursor.fetchone()[0]
            return count
        except Exception as e:
            logger.error(f"Ошибка подсчета премиум: {e}")
            return 0
        finally:
            if conn:
                conn.close()
