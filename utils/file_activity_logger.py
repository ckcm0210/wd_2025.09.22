#!/usr/bin/env python3
"""
檔案活動記錄器 - 將檔案開啟/關閉事件記錄到數據庫和 timeline
"""
import os
import sqlite3
import json
from datetime import datetime
from typing import Dict, Optional
import config.settings as settings

class FileActivityLogger:
    """檔案活動記錄器"""
    
    def __init__(self):
        self.db_path = self._get_db_path()
        self._init_database()
    
    def _get_db_path(self) -> str:
        """取得數據庫路徑"""
        log_folder = getattr(settings, 'LOG_FOLDER', '.')
        return os.path.join(log_folder, 'file_activity.sqlite')
    
    def _init_database(self):
        """初始化數據庫表"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute('''
                    CREATE TABLE IF NOT EXISTS file_activity (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        file_path TEXT NOT NULL,
                        file_name TEXT NOT NULL,
                        action TEXT NOT NULL,  -- 'open' or 'close'
                        user_name TEXT,
                        timestamp TEXT NOT NULL,
                        temp_files TEXT,  -- JSON array of temp files
                        duration_seconds REAL,  -- only for close events
                        session_id TEXT  -- to link open/close events
                    )
                ''')
                
                # 創建索引
                conn.execute('''
                    CREATE INDEX IF NOT EXISTS idx_file_activity_timestamp 
                    ON file_activity (timestamp DESC)
                ''')
                conn.execute('''
                    CREATE INDEX IF NOT EXISTS idx_file_activity_file 
                    ON file_activity (file_path, action)
                ''')
                conn.execute('''
                    CREATE INDEX IF NOT EXISTS idx_file_activity_user 
                    ON file_activity (user_name)
                ''')
                
                conn.commit()
        except Exception as e:
            print(f"[file_activity] 數據庫初始化失敗: {e}")
    
    def log_file_open(self, file_path: str, user_name: str, temp_files: set, session_id: str):
        """記錄檔案開啟事件"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute('''
                    INSERT INTO file_activity 
                    (file_path, file_name, action, user_name, timestamp, temp_files, session_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    file_path,
                    os.path.basename(file_path),
                    'open',
                    user_name,
                    datetime.now().isoformat(),
                    json.dumps(list(temp_files)),
                    session_id
                ))
                conn.commit()
        except Exception as e:
            print(f"[file_activity] 記錄檔案開啟失敗: {e}")
    
    def log_file_close(self, file_path: str, user_name: str, duration_seconds: float, session_id: str):
        """記錄檔案關閉事件"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute('''
                    INSERT INTO file_activity 
                    (file_path, file_name, action, user_name, timestamp, duration_seconds, session_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    file_path,
                    os.path.basename(file_path),
                    'close',
                    user_name,
                    datetime.now().isoformat(),
                    duration_seconds,
                    session_id
                ))
                conn.commit()
        except Exception as e:
            print(f"[file_activity] 記錄檔案關閉失敗: {e}")
    
    def get_recent_activities(self, hours: int = 24, limit: int = 100) -> list:
        """取得最近的檔案活動"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.execute('''
                    SELECT * FROM file_activity 
                    WHERE timestamp >= datetime('now', '-{} hours')
                    ORDER BY timestamp DESC 
                    LIMIT ?
                '''.format(hours), (limit,))
                return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            print(f"[file_activity] 查詢活動失敗: {e}")
            return []
    
    def get_user_statistics(self, hours: int = 24) -> dict:
        """取得用戶統計"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.execute('''
                    SELECT 
                        user_name,
                        COUNT(*) as total_actions,
                        COUNT(CASE WHEN action = 'open' THEN 1 END) as open_count,
                        COUNT(CASE WHEN action = 'close' THEN 1 END) as close_count,
                        AVG(CASE WHEN action = 'close' THEN duration_seconds END) as avg_duration
                    FROM file_activity 
                    WHERE timestamp >= datetime('now', '-{} hours')
                    AND user_name IS NOT NULL
                    GROUP BY user_name
                    ORDER BY total_actions DESC
                '''.format(hours))
                return {row['user_name']: dict(row) for row in cursor.fetchall()}
        except Exception as e:
            print(f"[file_activity] 查詢用戶統計失敗: {e}")
            return {}

# 全域實例
_file_activity_logger: Optional[FileActivityLogger] = None

def get_file_activity_logger() -> FileActivityLogger:
    """取得檔案活動記錄器單例"""
    global _file_activity_logger
    if _file_activity_logger is None:
        _file_activity_logger = FileActivityLogger()
    return _file_activity_logger

def log_file_activity(action: str, file_path: str, user_name: str, **kwargs):
    """記錄檔案活動的便利函數"""
    logger = get_file_activity_logger()
    
    if action == 'open':
        temp_files = kwargs.get('temp_files', set())
        session_id = kwargs.get('session_id', '')
        logger.log_file_open(file_path, user_name, temp_files, session_id)
    elif action == 'close':
        duration_seconds = kwargs.get('duration_seconds', 0.0)
        session_id = kwargs.get('session_id', '')
        logger.log_file_close(file_path, user_name, duration_seconds, session_id)