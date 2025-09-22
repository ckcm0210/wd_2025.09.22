import os
import time
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import config.settings as settings
import logging
from datetime import datetime

class ActivePollingHandler:
    """
    主動輪詢處理器，採用新的智慧輪詢邏輯 + 穩定窗口/冷靜期
    """
    def __init__(self):
        self.polling_tasks = {}
        self.lock = threading.Lock()
        self.stop_event = threading.Event()
        # 狀態表（每檔案）
        # { file_path: {"last_mtime":float, "last_size":int, "stable":int, "cooldown_until":float} }
        self.state = {}

    def start_polling(self, file_path, event_number):
        """
        根據檔案大小決定輪詢策略（用 mtime/size 穩定檢查，不再用與 baseline 的差異判斷）
        """
        # 停止中不再啟動新輪詢
        try:
            if self.stop_event.is_set() or getattr(settings, 'force_stop', False):
                return
        except Exception:
            pass
        try:
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        except (FileNotFoundError, PermissionError, OSError) as e:
            logging.warning(f"獲取檔案大小失敗: {file_path}, 錯誤: {e}")
            file_size_mb = 0

        interval = settings.DENSE_POLLING_INTERVAL_SEC if file_size_mb < settings.POLLING_SIZE_THRESHOLD_MB else settings.SPARSE_POLLING_INTERVAL_SEC
        polling_type = "密集" if file_size_mb < settings.POLLING_SIZE_THRESHOLD_MB else "稀疏"
        
        print(f"[輪詢] 檔案: {os.path.basename(file_path)}（{polling_type}輪詢，每 {interval}s 檢查一次；首次檢查 {interval}s 後）")
        # 初始化 last_mtime/size 與狀態
        try:
            last_mtime = os.path.getmtime(file_path)
        except Exception:
            last_mtime = 0
        try:
            last_size = os.path.getsize(file_path)
        except Exception:
            last_size = -1
        with self.lock:
            self.state[file_path] = {"last_mtime": last_mtime, "last_size": last_size, "stable": 0, "cooldown_until": 0.0}
        self._start_adaptive_polling(file_path, event_number, interval, last_mtime)

    def _start_adaptive_polling(self, file_path, event_number, interval, last_mtime):
        """
        開始自適應輪詢
        """
        # 停止中不再排程
        if self.stop_event.is_set() or getattr(settings, 'force_stop', False):
            return
        with self.lock:
            if file_path in self.polling_tasks:
                self.polling_tasks[file_path]['timer'].cancel()

            def task_wrapper():
                if self.stop_event.is_set() or getattr(settings, 'force_stop', False):
                    return
                self._poll_for_stability(file_path, event_number, interval, last_mtime)

            timer = threading.Timer(interval, task_wrapper)
            self.polling_tasks[file_path] = {'timer': timer}
            timer.start()

    def _poll_for_stability(self, file_path, event_number, interval, last_mtime):
        """
        執行輪詢檢查：使用 mtime/size 的穩定窗口策略，並包含冷靜期與暫存鎖檔判斷
        """
        if self.stop_event.is_set():
            return

        # 冷靜期判斷
        st = self.state.get(file_path, {})
        now = time.time()
        if st and now < st.get("cooldown_until", 0):
            print(f"    [cooldown] {os.path.basename(file_path)} 尚在冷靜期，略過本次。")
            # 重新排程
            with self.lock:
                if file_path in self.polling_tasks:
                    new_timer = threading.Timer(interval, lambda: self._poll_for_stability(file_path, event_number, interval, last_mtime))
                    self.polling_tasks[file_path]['timer'] = new_timer
                    new_timer.start()
            return

        # 檢測暫存鎖檔 (~$)
        if getattr(settings, 'SKIP_WHEN_TEMP_LOCK_PRESENT', True):
            tmp_lock = os.path.join(os.path.dirname(file_path), "~$" + os.path.basename(file_path))
            try:
                if os.path.exists(tmp_lock):
                    print(f"    [鎖檔] 偵測到 {os.path.basename(tmp_lock)}，延後檢查。")
                    with self.lock:
                        if file_path in self.polling_tasks:
                            new_timer = threading.Timer(interval, lambda: self._poll_for_stability(file_path, event_number, interval, last_mtime))
                            self.polling_tasks[file_path]['timer'] = new_timer
                            new_timer.start()
                    return
            except Exception:
                pass

        print(f"    [輪詢檢查] 正在檢查 {os.path.basename(file_path)} 的變更...")

        # 以 mtime/size 穩定判斷
        try:
            cur_mtime = os.path.getmtime(file_path)
        except Exception:
            cur_mtime = last_mtime
        try:
            cur_size = os.path.getsize(file_path)
        except Exception:
            cur_size = self.state.get(file_path, {}).get('last_size', -1)

        changed = False
        st = self.state.get(file_path, {})
        if st:
            if cur_mtime != st.get('last_mtime') or cur_size != st.get('last_size'):
                changed = True
                st['last_mtime'] = cur_mtime
                st['last_size'] = cur_size
                st['stable'] = 0
                print(f"    [輪詢] 檢測到變動，等待穩定窗口（{getattr(settings,'POLLING_STABLE_CHECKS',3)} 次）…")
            else:
                st['stable'] = st.get('stable', 0) + 1
        
        has_changes = False
        if st and st.get('stable', 0) >= getattr(settings, 'POLLING_STABLE_CHECKS', 3):
            from core.comparison import compare_excel_changes, set_current_event_number
            set_current_event_number(event_number)
            print(f"    [輪詢] 已穩定，開始比較…")
            has_changes = compare_excel_changes(file_path, silent=False, event_number=event_number, is_polling=True)

        with self.lock:
            if file_path not in self.polling_tasks:
                return

            if has_changes:
                try:
                    _sz_mb = os.path.getsize(file_path)/(1024*1024)
                    _sz_str = f"{_sz_mb:.2f}MB"
                except Exception:
                    _sz_str = "N/A"
                print(f"    [輪詢] 變更仍持續（事件 #{event_number}，大小 {_sz_str}），啟動冷靜期，{getattr(settings,'POLLING_COOLDOWN_SEC',20)} 秒後再次檢查。")
                st['cooldown_until'] = time.time() + float(getattr(settings, 'POLLING_COOLDOWN_SEC', 20))
                st['stable'] = 0
                new_timer = threading.Timer(interval, lambda: self._poll_for_stability(file_path, event_number, interval, last_mtime))
                self.polling_tasks[file_path]['timer'] = new_timer
                new_timer.start()
            else:
                # 若尚未達穩定次數，或剛檢測到變動，繼續等待；若已穩定且無變更，結束輪詢
                if st and st.get('stable', 0) < getattr(settings, 'POLLING_STABLE_CHECKS', 3):
                    new_timer = threading.Timer(interval, lambda: self._poll_for_stability(file_path, event_number, interval, last_mtime))
                    self.polling_tasks[file_path]['timer'] = new_timer
                    new_timer.start()
                elif changed:
                    new_timer = threading.Timer(interval, lambda: self._poll_for_stability(file_path, event_number, interval, last_mtime))
                    self.polling_tasks[file_path]['timer'] = new_timer
                    new_timer.start()
                else:
                    print(f"    [輪詢結束] {os.path.basename(file_path)} 檔案已穩定。")
                    self.polling_tasks.pop(file_path, None)
                    self.state.pop(file_path, None)

    def stop(self):
        """
        停止所有輪詢任務
        """
        self.stop_event.set()
        with self.lock:
            for task in self.polling_tasks.values():
                task['timer'].cancel()
            self.polling_tasks.clear()

class ExcelFileEventHandler(FileSystemEventHandler):
    """
    Excel 檔案事件處理器
    """
    def __init__(self, polling_handler):
        self.polling_handler = polling_handler
        self.last_event_times = {}
        self.event_counter = 0
        # 檔案開啟/關閉狀態追蹤
        self.file_open_status = {}  # {file_path: {'is_open': bool, 'temp_files': set(), 'opened_at': timestamp, 'last_author': str}}
        self.temp_file_tracking = {}  # {temp_file_path: 'original_file_path'}
        
    def _is_cache_ignored(self, path: str) -> bool:
        try:
            if getattr(settings, 'IGNORE_CACHE_FOLDER', False) and getattr(settings, 'CACHE_FOLDER', None):
                p = os.path.abspath(path)
                c = os.path.abspath(settings.CACHE_FOLDER)
                return os.path.commonpath([p, c]) == c
        except Exception:
            pass
        return False

    def _is_log_ignored(self, path: str) -> bool:
        try:
            if getattr(settings, 'IGNORE_LOG_FOLDER', False) and getattr(settings, 'LOG_FOLDER', None):
                p = os.path.abspath(path)
                l = os.path.abspath(settings.LOG_FOLDER)
                return os.path.commonpath([p, l]) == l
        except Exception:
            pass
        return False

    def _is_temp_file(self, file_path: str) -> str:
        """
        檢查是否為 Excel 臨時檔案，返回對應的原始檔案路徑
        Excel 臨時檔案類型：
        1. ~$filename.xlsx (鎖檔)
        2. filename.tmp
        3. ~WRLxxxx.tmp
        4. 其他 Office 臨時檔案模式
        """
        filename = os.path.basename(file_path)
        dirname = os.path.dirname(file_path)
        
        # 類型1: ~$檔名.xlsx (最常見的 Excel 鎖檔)
        if filename.startswith('~$'):
            original_name = filename[2:]  # 移除 ~$ 前綴
            original_path = os.path.join(dirname, original_name)
            if os.path.exists(original_path) and original_path.lower().endswith(settings.SUPPORTED_EXTS):
                return original_path
        
        # 類型2: 檔名.tmp
        if filename.lower().endswith('.tmp'):
            # 檢查同目錄下是否有同名的 Excel 檔案
            base_name = filename[:-4]  # 移除 .tmp
            for ext in settings.SUPPORTED_EXTS:
                potential_original = os.path.join(dirname, base_name + ext)
                if os.path.exists(potential_original):
                    return potential_original
        
        # 類型3: ~WRLxxxx.tmp (Word/Excel 臨時檔案)
        if filename.startswith('~WRL') and filename.lower().endswith('.tmp'):
            # 這種類型較難直接對應，需要根據時間和目錄推測
            for file in os.listdir(dirname):
                if file.lower().endswith(settings.SUPPORTED_EXTS):
                    potential_path = os.path.join(dirname, file)
                    # 檢查該檔案是否最近被修改
                    try:
                        mtime = os.path.getmtime(potential_path)
                        if time.time() - mtime < 60:  # 1分鐘內修改過
                            return potential_path
                    except Exception:
                        continue
        
        return None

    def _track_file_open(self, original_file: str, temp_file: str, author: str = None):
        """追蹤檔案開啟狀態"""
        current_time = datetime.now()
        
        # 生成 session ID
        import uuid
        session_id = str(uuid.uuid4())[:8]
        
        if original_file not in self.file_open_status:
            self.file_open_status[original_file] = {
                'is_open': True,
                'temp_files': set(),
                'opened_at': current_time,
                'last_author': author,
                'session_id': session_id
            }
            print(f"📂 檔案開啟: {os.path.basename(original_file)}")
            if author:
                print(f"   👤 使用者: {author}")
            print(f"   🕒 開啟時間: {current_time.strftime('%H:%M:%S')}")
            
            # 記錄到數據庫和 timeline
            try:
                from utils.file_activity_logger import log_file_activity
                log_file_activity(
                    action='open',
                    file_path=original_file,
                    user_name=author or '未知用戶',
                    temp_files={temp_file},
                    session_id=session_id
                )
            except Exception as e:
                print(f"   ⚠️ 記錄開啟事件失敗: {e}")
        else:
            # 更新狀態
            self.file_open_status[original_file]['is_open'] = True
            if author:
                self.file_open_status[original_file]['last_author'] = author
        
        # 追蹤臨時檔案
        self.file_open_status[original_file]['temp_files'].add(temp_file)
        self.temp_file_tracking[temp_file] = original_file

    def _track_file_close(self, temp_file: str):
        """追蹤檔案關閉狀態"""
        if temp_file in self.temp_file_tracking:
            original_file = self.temp_file_tracking[temp_file]
            
            if original_file in self.file_open_status:
                # 移除臨時檔案
                self.file_open_status[original_file]['temp_files'].discard(temp_file)
                
                # 如果沒有剩餘的臨時檔案，表示檔案已關閉
                if not self.file_open_status[original_file]['temp_files']:
                    current_time = datetime.now()
                    opened_at = self.file_open_status[original_file]['opened_at']
                    duration = current_time - opened_at
                    author = self.file_open_status[original_file].get('last_author', '未知')
                    session_id = self.file_open_status[original_file].get('session_id', '')
                    
                    print(f"📁 檔案關閉: {os.path.basename(original_file)}")
                    print(f"   👤 使用者: {author}")
                    print(f"   🕒 關閉時間: {current_time.strftime('%H:%M:%S')}")
                    print(f"   ⏱️ 開啟時長: {str(duration).split('.')[0]}")
                    
                    # 記錄到數據庫和 timeline
                    try:
                        from utils.file_activity_logger import log_file_activity
                        log_file_activity(
                            action='close',
                            file_path=original_file,
                            user_name=author,
                            duration_seconds=duration.total_seconds(),
                            session_id=session_id
                        )
                    except Exception as e:
                        print(f"   ⚠️ 記錄關閉事件失敗: {e}")
                    
                    # 標記為關閉
                    self.file_open_status[original_file]['is_open'] = False
                    
                    # 可選：清理舊狀態（保留一段時間供查詢）
                    # del self.file_open_status[original_file]
            
            # 清理臨時檔案追蹤
            del self.temp_file_tracking[temp_file]

    def get_file_status(self, file_path: str) -> dict:
        """取得檔案開啟狀態"""
        return self.file_open_status.get(file_path, {'is_open': False})

    def on_created(self, event):
        try:
            import time as _t
            settings.LAST_DISPATCH_TS = _t.time()
        except Exception:
            pass
        """
        檔案建立事件處理
        """
        if event.is_directory:
            return
        # 停止中：忽略新事件
        try:
            if getattr(settings, 'force_stop', False):
                return
        except Exception:
            pass

        file_path = event.src_path

        # 首先檢查是否為臨時檔案
        original_file = self._is_temp_file(file_path)
        if original_file:
            # 這是一個臨時檔案，追蹤檔案開啟
            try:
                # 嘗試取得最後作者
                author = None
                try:
                    if getattr(settings, 'ENABLE_LAST_AUTHOR_LOOKUP', True):
                        from core.excel_parser import get_excel_last_author
                        author = get_excel_last_author(original_file)
                except Exception:
                    pass
                
                self._track_file_open(original_file, file_path, author)
            except Exception as e:
                print(f"   ⚠️ 追蹤檔案開啟失敗: {e}")
            return

        # 檢查是否為支援的 Excel 檔案
        if not file_path.lower().endswith(settings.SUPPORTED_EXTS):
            return

        print(f"\n✨ 發現新檔案: {os.path.basename(file_path)}")
        print(f"📊 正在建立基準線...")

        # 等源檔短暫穩定再建 baseline（避免剛複製完即讀取）
        try:
            checks = max(1, int(getattr(settings, 'COPY_STABILITY_CHECKS', 3)))
            interval = max(0.0, float(getattr(settings, 'COPY_STABILITY_INTERVAL_SEC', 1.0)))
            max_wait = float(getattr(settings, 'COPY_STABILITY_MAX_WAIT_SEC', 5.0))
            from utils.cache import _wait_for_stable_mtime as _stable
            _ = _stable(file_path, checks, interval, max_wait)
        except Exception:
            pass

        from core.baseline import create_baseline_for_files_robust
        create_baseline_for_files_robust([file_path])

        print(f"✅ 基準線建立完成，已納入監控: {os.path.basename(file_path)}")

    def on_deleted(self, event):
        try:
            import time as _t
            settings.LAST_DISPATCH_TS = _t.time()
        except Exception:
            pass
        """
        檔案刪除事件處理 - 主要用於檢測臨時檔案刪除（檔案關閉）
        """
        if event.is_directory:
            return
        
        # 停止中：忽略新事件
        try:
            if getattr(settings, 'force_stop', False):
                return
        except Exception:
            pass

        file_path = event.src_path
        
        # 檢查是否為被追蹤的臨時檔案
        if file_path in self.temp_file_tracking:
            try:
                self._track_file_close(file_path)
            except Exception as e:
                print(f"   ⚠️ 追蹤檔案關閉失敗: {e}")

    def _is_in_watch_folders(self, path: str) -> bool:
        try:
            p = os.path.normcase(os.path.abspath(path))
            for root in (settings.WATCH_FOLDERS or []):
                r = os.path.normcase(os.path.abspath(root))
                try:
                    if os.path.commonpath([p, r]) == r:
                        # 排除清單
                        for ex in (getattr(settings, 'WATCH_EXCLUDE_FOLDERS', []) or []):
                            exa = os.path.normcase(os.path.abspath(ex))
                            if os.path.commonpath([p, exa]) == exa:
                                return False
                        return True
                except Exception:
                    # commonpath 可能在不同磁碟時丟例外，退回 startswith 判斷
                    if p.startswith(r.rstrip('\\/') + os.sep):
                        return True
        except Exception:
            pass
        return False

    def _is_monitor_only(self, path: str) -> bool:
        # WATCH_FOLDERS 優先於 MONITOR_ONLY_FOLDERS
        if self._is_in_watch_folders(path):
            return False
        try:
            p = os.path.abspath(path)
            for root in (settings.MONITOR_ONLY_FOLDERS or []):
                r = os.path.abspath(root)
                if os.path.commonpath([p, r]) == r:
                    # 排除清單
                    for ex in (getattr(settings, 'MONITOR_ONLY_EXCLUDE_FOLDERS', []) or []):
                        exa = os.path.abspath(ex)
                        if os.path.commonpath([p, exa]) == exa:
                            return False
                    return True
        except Exception:
            pass
        return False

    def on_modified(self, event):
        try:
            import time as _t
            settings.LAST_DISPATCH_TS = _t.time()
        except Exception:
            pass
        """
        檔案修改事件處理
        """
        if event.is_directory:
            return
        # 停止中：忽略事件
        try:
            if getattr(settings, 'force_stop', False):
                return
        except Exception:
            pass
            
        file_path = event.src_path

        # 路由判斷（加入 debug）：WATCH 優先於 MONITOR-ONLY
        in_watch = self._is_in_watch_folders(file_path)
        in_mononly = self._is_monitor_only(file_path) if not in_watch else False
        try:
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print(f"   [route] in_watch={in_watch} in_mononly={in_mononly} path={file_path}")
        except Exception:
            pass
        
        # 忽略 cache 與 log 目錄下的所有事件
        if self._is_cache_ignored(file_path) or self._is_log_ignored(file_path):
            return
        
        # 檢查是否為支援的 Excel 檔案
        if not file_path.lower().endswith(settings.SUPPORTED_EXTS):
            return
            
        # 檢查是否為臨時檔案（已由 on_created/on_deleted 處理）
        original_file = self._is_temp_file(file_path)
        if original_file:
            # 臨時檔案的修改事件，可能需要更新作者信息
            try:
                author = None
                try:
                    if getattr(settings, 'ENABLE_LAST_AUTHOR_LOOKUP', True):
                        from core.excel_parser import get_excel_last_author
                        author = get_excel_last_author(original_file)
                except Exception:
                    pass
                
                if original_file in self.file_open_status and author:
                    self.file_open_status[original_file]['last_author'] = author
            except Exception:
                pass
            return
            
        # 防抖動處理
        current_time = time.time()
        if file_path in self.last_event_times:
            if current_time - self.last_event_times[file_path] < settings.DEBOUNCE_INTERVAL_SEC:
                return
                
        self.last_event_times[file_path] = current_time
        self.event_counter += 1
        
        # 獲取檔案最後作者
        try:
            last_author = None
            try:
                if getattr(settings, 'ENABLE_LAST_AUTHOR_LOOKUP', True):
                    from core.excel_parser import get_excel_last_author
                    last_author = get_excel_last_author(file_path)
            except Exception:
                last_author = None
            author_info = f" (最後儲存者: {last_author})" if last_author else ""
        except Exception as e:
            author_info = ""

        # 統一事件標題
        print(f"\n🔔 檔案變更偵測: {os.path.basename(file_path)} (事件 #{self.event_counter}){author_info}")

        # Monitor-only 首次事件：直接建立基準線（僅當不屬於 WATCH_FOLDERS 時）
        if (not in_watch) and in_mononly:
            try:
                from core.baseline import get_baseline_file_with_extension, save_baseline
                from core.excel_parser import dump_excel_cells_with_timeout, hash_excel_content
                from utils.helpers import _baseline_key_for_path, get_file_mtime
                base_key = _baseline_key_for_path(file_path)
                baseline_exists = bool(get_baseline_file_with_extension(base_key))
                if not baseline_exists:
                    mtime = get_file_mtime(file_path)
                    print(f"    [MONITOR-ONLY] {file_path}\n       - 最後修改時間: {mtime}\n       - 最後儲存者: {last_author}")
                    cur = dump_excel_cells_with_timeout(file_path)
                    if cur:
                        bdata = {
                            "last_author": last_author,
                            "content_hash": hash_excel_content(cur),
                            "cells": cur,
                            "timestamp": datetime.now().isoformat()
                        }
                        # 補上 source_mtime/size 供快速跳過判斷使用
                        try:
                            bdata["source_mtime"] = os.path.getmtime(file_path)
                            bdata["source_size"] = os.path.getsize(file_path)
                        except Exception:
                            pass
                        save_baseline(base_key, bdata)
                        print("    [MONITOR-ONLY] 已建立首次基準線（本次不比較）。")
                        return
            except Exception as e:
                logging.warning(f"monitor-only 初始化失敗: {e}")
                return
        
        # 🔥 佇列化比對流程
        from core.comparison import compare_excel_changes, set_current_event_number
        from utils.task_queue import get_compare_queue
        set_current_event_number(self.event_counter)

        # 首次即時比較（維持體感）
        has_changes = False
        if getattr(settings, 'IMMEDIATE_COMPARE_ON_FIRST_EVENT', True):
            if file_path in self.polling_handler.polling_tasks:
                st = self.polling_handler.state.get(file_path, {})
                if not st.get('has_shown_initial_compare', False):
                    print(f"📊 立即檢查變更（輪詢中首次）...")
                    has_changes = compare_excel_changes(file_path, silent=False, event_number=self.event_counter, is_polling=False)
                    st['has_shown_initial_compare'] = True
                    self.polling_handler.state[file_path] = st
                else:
                    if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                        print(f"    [偵測] {os.path.basename(file_path)} 正在輪詢中，已顯示過即時比較，改交佇列處理。")
            else:
                print(f"📊 立即檢查變更...")
                has_changes = compare_excel_changes(file_path, silent=False, event_number=self.event_counter, is_polling=False)

        # 將後續比較任務交由佇列，限制並行與去重
        q = get_compare_queue(lambda p, evt: compare_excel_changes(p, silent=False, event_number=evt, is_polling=False))
        q.submit(file_path, self.event_counter)

        if has_changes:
            print(f"✅ 偵測到變更，啟動輪詢以監控後續活動...")
        else:
            print(f"ℹ️  佇列已接手後續比較與輸出…")

        # 開始輪詢
        self.polling_handler.start_polling(file_path, self.event_counter)

# 創建全局輪詢處理器實例
active_polling_handler = ActivePollingHandler()

# 啟動 Idle-GC Scheduler（背景執行），避免在臨界區期間觸發 GC
try:
    from utils.idle_gc_scheduler import start_idle_gc
    start_idle_gc()
except Exception:
    pass