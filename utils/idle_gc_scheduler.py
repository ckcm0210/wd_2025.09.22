"""
Idle-GC Scheduler: run GC only when the system is idle enough.
- Idle condition: (guard_count == 0) and (now - last_activity_ts >= calm_sec)
- Run in a background thread; start/stop via API.
"""
import threading
import time
import gc
import logging
import threading as _th
import sys
from typing import Optional

try:
    import config.settings as settings
except Exception:
    class settings:
        IDLE_GC_ENABLED = True
        IDLE_GC_CALM_SEC = 8
        IDLE_GC_TICK_SEC = 2
        IDLE_GC_COLLECT_GENERATION = -1  # -1 for full; 0 for gen0
        IDLE_GC_MAX_COLLECT_PER_MIN = 6
        SHOW_DEBUG_MESSAGES = False

from utils.gc_guard import get_guard_count, get_last_activity_ts


class IdleGCScheduler:
    def __init__(self,
                 calm_sec: Optional[float] = None,
                 tick_sec: Optional[float] = None,
                 collect_gen: Optional[int] = None,
                 max_collect_per_min: Optional[int] = None):
        self.calm_sec = float(calm_sec if calm_sec is not None else getattr(settings, 'IDLE_GC_CALM_SEC', 8))
        self.tick_sec = float(tick_sec if tick_sec is not None else getattr(settings, 'IDLE_GC_TICK_SEC', 2))
        self.collect_gen = int(collect_gen if collect_gen is not None else getattr(settings, 'IDLE_GC_COLLECT_GENERATION', -1))
        self.max_collect_per_min = int(max_collect_per_min if max_collect_per_min is not None else getattr(settings, 'IDLE_GC_MAX_COLLECT_PER_MIN', 6))
        self._stop = threading.Event()
        self._th: Optional[threading.Thread] = None
        self._last_collect_min = int(time.time() // 60)
        self._collect_count_in_min = 0

    def start(self):
        if not getattr(settings, 'IDLE_GC_ENABLED', True):
            return False
        if self._th and self._th.is_alive():
            return True
        self._stop.clear()
        self._th = threading.Thread(target=self._run, name='IdleGCScheduler', daemon=True)
        self._th.start()
        if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
            print(f"[idle-gc] started: calm={self.calm_sec}s tick={self.tick_sec}s gen={self.collect_gen} cap/min={self.max_collect_per_min}")
        return True

    def stop(self, wait: bool = False):
        self._stop.set()
        if wait and self._th:
            try:
                self._th.join(timeout=2.0)
            except Exception:
                pass

    def _is_tk_active(self) -> bool:
        try:
            # 檢查是否有任何 Tkinter 相關的線程或實例
            import sys
            
            # 檢查是否有 tkinter 模組被載入
            if 'tkinter' in sys.modules:
                # 檢查 black_console 是否存在且運行中
                from ui.console import black_console as _bc
                if _bc and getattr(_bc, 'running', False):
                    return True
                
                # 檢查是否有 Tk 根視窗存在
                try:
                    import tkinter as tk
                    root = tk._default_root
                    if root is not None:
                        return True
                except Exception:
                    pass
            
            return False
        except Exception:
            # 發生任何錯誤時，保守起見認為 Tk 可能活躍
            return True

    def _collect_safely(self):
        """Perform GC without calling any Tk APIs from non-Tk threads.
        If Tk is active and settings ask to skip, we skip. Otherwise, we collect with selected generation.
        """
        if self._is_tk_active() and getattr(settings, 'IDLE_GC_SKIP_WHEN_TK', True):
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print("[idle-gc] skip collect (Tk active)")
            return False
        
        try:
            # 額外的安全檢查：確保沒有正在進行的檔案操作
            import threading
            current_thread = threading.current_thread()
            if current_thread.name != 'IdleGCScheduler':
                if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                    print(f"[idle-gc] skip collect (wrong thread: {current_thread.name})")
                return False
            
            # 執行垃圾回收前再次確認 Tk 狀態
            if self._is_tk_active():
                if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                    print("[idle-gc] skip collect (Tk became active)")
                return False
            
            gen = self.collect_gen
            if gen in (0, 1, 2):
                gc.collect(gen)
            else:
                gc.collect()
            return True
        except Exception as e:
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print(f"[idle-gc] collect failed safely: {e}")
            return False

    def _run(self):
        while not self._stop.is_set():
            try:
                # limit per minute
                now_min = int(time.time() // 60)
                if now_min != self._last_collect_min:
                    self._last_collect_min = now_min
                    self._collect_count_in_min = 0
                # idle check
                g = get_guard_count()
                idle_for = time.time() - float(get_last_activity_ts() or 0)
                if g == 0 and idle_for >= self.calm_sec and self._collect_count_in_min < self.max_collect_per_min:
                    if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                        print(f"[idle-gc] collect gen={self.collect_gen} idle_for={idle_for:.1f}s g={g}")
                    try:
                        self._collect_safely()
                        self._collect_count_in_min += 1
                    except Exception as e:
                        logging.debug(f"[idle-gc] collect failed: {e}")
            except Exception as e:
                logging.debug(f"[idle-gc] loop err: {e}")
            finally:
                self._stop.wait(self.tick_sec)


# module-level singleton helpers
_scheduler: Optional[IdleGCScheduler] = None

def start_idle_gc():
    global _scheduler
    if _scheduler is None:
        _scheduler = IdleGCScheduler()
    return _scheduler.start()

def stop_idle_gc(wait: bool = False):
    global _scheduler
    if _scheduler:
        _scheduler.stop(wait=wait)
