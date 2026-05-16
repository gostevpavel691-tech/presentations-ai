"""
Модуль для работы с кэшем презентаций
Поддерживает JSON и SQLite хранилища
"""

import os
import json
import sqlite3
import hashlib
import logging
from abc import ABC, abstractmethod
from datetime import datetime, timezone

logger = logging.getLogger(__name__)


class CacheStorage(ABC):
    """Абстрактный класс для хранения кэша"""
    
    @abstractmethod
    def get(self, key):
        pass
    
    @abstractmethod
    def add(self, key, filepath):
        pass
    
    @abstractmethod
    def delete(self, key):
        pass
    
    @abstractmethod
    def clear(self):
        pass
    
    @abstractmethod
    def get_all(self):
        pass
    
    @abstractmethod
    def size(self):
        pass


class JSONCacheStorage(CacheStorage):
    """Хранение кэша в JSON файле"""
    
    def __init__(self, cache_dir):
        self.cache_dir = cache_dir
        self.cache_file = os.path.join(cache_dir, "cache.json")
        self.cache = self._load_cache()
        logger.info("Используется JSON кэш")
    
    def _load_cache(self):
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"Ошибка загрузки JSON кэша: {e}")
        return {}
    
    def _save_cache(self):
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"Ошибка сохранения JSON кэша: {e}")
    
    def get(self, key):
        if key in self.cache:
            path = self.cache[key]
            if os.path.exists(path):
                return path
            else:
                del self.cache[key]
                self._save_cache()
        return None
    
    def add(self, key, filepath):
        self.cache[key] = filepath
        self._save_cache()
    
    def delete(self, key):
        if key in self.cache:
            del self.cache[key]
            self._save_cache()
    
    def clear(self):
        self.cache = {}
        self._save_cache()
    
    def get_all(self):
        return self.cache.copy()
    
    def size(self):
        return len(self.cache)


class SQLiteCacheStorage(CacheStorage):
    """Хранение кэша в SQLite базе данных"""
    
    def __init__(self, cache_dir):
        self.cache_dir = cache_dir
        self.db_path = os.path.join(cache_dir, "cache.db")
        self._init_db()
        logger.info("Используется SQLite кэш")
    
    def _init_db(self):
        """Создает таблицу кэша если её нет"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cache (
                    key TEXT PRIMARY KEY,
                    filepath TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    last_accessed TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    access_count INTEGER DEFAULT 0
                )
            ''')
            conn.commit()
        except Exception as e:
            logger.error(f"Ошибка инициализации SQLite кэша: {e}")
        finally:
            if conn:
                conn.close()
    
    def _update_access(self, key):
        """Обновляет время последнего доступа и счетчик"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE cache 
                SET last_accessed = CURRENT_TIMESTAMP, 
                    access_count = access_count + 1 
                WHERE key = ?
            ''', (key,))
            conn.commit()
        except Exception as e:
            logger.error(f"Ошибка обновления статистики доступа: {e}")
        finally:
            if conn:
                conn.close()
    
    def get(self, key):
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT filepath FROM cache WHERE key = ?', (key,))
            result = cursor.fetchone()
            
            if result:
                filepath = result[0]
                if os.path.exists(filepath):
                    self._update_access(key)
                    return filepath
                else:
                    self.delete(key)
            return None
        except Exception as e:
            logger.error(f"Ошибка получения из SQLite кэша: {e}")
            return None
        finally:
            if conn:
                conn.close()
    
    def add(self, key, filepath):
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO cache (key, filepath) 
                VALUES (?, ?)
            ''', (key, filepath))
            conn.commit()
        except Exception as e:
            logger.error(f"Ошибка добавления в SQLite кэш: {e}")
        finally:
            if conn:
                conn.close()
    
    def delete(self, key):
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM cache WHERE key = ?', (key,))
            conn.commit()
        except Exception as e:
            logger.error(f"Ошибка удаления из SQLite кэша: {e}")
        finally:
            if conn:
                conn.close()
    
    def clear(self):
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM cache')
            conn.commit()
        except Exception as e:
            logger.error(f"Ошибка очистки SQLite кэша: {e}")
        finally:
            if conn:
                conn.close()
    
    def get_all(self):
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT key, filepath FROM cache')
            results = cursor.fetchall()
            return {key: filepath for key, filepath in results}
        except Exception as e:
            logger.error(f"Ошибка получения всех записей из SQLite: {e}")
            return {}
        finally:
            if conn:
                conn.close()
    
    def size(self):
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM cache')
            count = cursor.fetchone()[0]
            return count
        except Exception as e:
            logger.error(f"Ошибка подсчета записей в SQLite: {e}")
            return 0
        finally:
            if conn:
                conn.close()
    
    def get_stats(self):
        """Получить статистику кэша"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                SELECT 
                    COUNT(*) as total,
                    SUM(access_count) as total_access,
                    AVG(access_count) as avg_access,
                    MAX(access_count) as max_access
                FROM cache
            ''')
            stats = cursor.fetchone()
            return {
                'total': stats[0] or 0,
                'total_access': stats[1] or 0,
                'avg_access': round(stats[2] or 0, 2),
                'max_access': stats[3] or 0
            }
        except Exception as e:
            logger.error(f"Ошибка получения статистики: {e}")
            return {'total': 0, 'total_access': 0, 'avg_access': 0, 'max_access': 0}
        finally:
            if conn:
                conn.close()
    
    def clean_old(self, days=7):
        """Удаляет записи старше N дней"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                DELETE FROM cache 
                WHERE last_accessed < datetime('now', ?)
            ''', (f'-{days} days',))
            deleted = cursor.rowcount
            conn.commit()
            logger.info(f"Удалено {deleted} старых записей из SQLite кэша")
            return deleted
        except Exception as e:
            logger.error(f"Ошибка очистки старых записей: {e}")
            return 0
        finally:
            if conn:
                conn.close()


class PresentationCache:
    """Обертка для кэша с возможностью выбора типа хранения"""
    
    def __init__(self, cache_dir, cache_type="json"):
        self.cache_dir = cache_dir
        self.cache_type = cache_type
        
        if cache_type == "sqlite":
            self.storage = SQLiteCacheStorage(cache_dir)
        else:
            self.storage = JSONCacheStorage(cache_dir)
    
    def get_key(self, request):
        """Создает ключ из запроса"""
        return hashlib.md5(request.lower().strip().encode()).hexdigest()
    
    def get(self, request):
        """Получить презентацию из кэша по запросу"""
        key = self.get_key(request)
        return self.storage.get(key)
    
    def add(self, request, filepath):
        """Добавить презентацию в кэш"""
        key = self.get_key(request)
        self.storage.add(key, filepath)
    
    def delete(self, request):
        """Удалить презентацию из кэша"""
        key = self.get_key(request)
        self.storage.delete(key)
    
    def clear(self):
        """Очистить весь кэш"""
        self.storage.clear()
    
    def get_all(self):
        """Получить все записи кэша"""
        return self.storage.get_all()
    
    def size(self):
        """Размер кэша"""
        return self.storage.size()
    
    def get_stats(self):
        """Получить статистику кэша (только для SQLite)"""
        if hasattr(self.storage, 'get_stats'):
            return self.storage.get_stats()
        return {'total': self.size(), 'total_access': 0, 'avg_access': 0, 'max_access': 0}
    
    def clean_old(self, days=7):
        """Очистить старые записи (только для SQLite)"""
        if hasattr(self.storage, 'clean_old'):
            return self.storage.clean_old(days)
        return 0