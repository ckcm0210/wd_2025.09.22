"""
GC 守護工具：在指定區塊內暫停循環垃圾回收（引用計數仍然生效），離開時恢復並可選擇觸發一次收集。
用法：

from utils.gc_guard import gc_guard
with gc_guard(enabled=True, do_collect=True):
    # openpyxl 解析或其他需避免 GC 插入的臨界區
    ...
"""
from contextlib import contextmanager
import gc
import threading
import time

# Global re-entrant, cross-thread GC guard state
_guard_lock = threading.RLock()
_guard_count = 0
# Track whether this guard disabled GC when entering from 0 -> 1
_guard_disabled_by_us = False
# Last activity timestamp for idle-GC scheduler
_last_activity_ts = time.time()

@contextmanager
def gc_guard(enabled: bool = True, do_collect: bool = True):
    """
    在 with 期間暫停循環 GC；離開時恢復原狀，並可選擇觸發一次 gc.collect()。
    注意：此為進程層級設定；請將守護區塊保持盡可能短。
    """
    if not enabled:
        yield
        return
    # 僅在主執行緒才暫停 GC，避免與 Tk/Tcl 內部的 async handler 發生意外交叉
    if threading.current_thread() is not threading.main_thread():
        yield
        return
    was_enabled = gc.isenabled()
    try:
        if was_enabled:
            gc.disable()
        yield
    finally:
        try:
            if was_enabled:
                gc.enable()
            # 移除 gc.collect() 以避免在 XML 解析後立即觸發 GC 導致 0x80000003 崩潰
            # if do_collect:
            #     try:
            #         gc.collect()
            #     except Exception:
            #         pass
        except Exception:
            pass

def get_guard_count() -> int:
    try:
        with _guard_lock:
            return int(_guard_count)
    except Exception:
        return 0


def get_last_activity_ts() -> float:
    try:
        with _guard_lock:
            return float(_last_activity_ts)
    except Exception:
        return 0.0


def note_activity() -> None:
    try:
        with _guard_lock:
            global _last_activity_ts
            _last_activity_ts = time.time()
    except Exception:
        pass


@contextmanager
def gc_guard_any_thread(enabled: bool = True, do_collect: bool = False):
    """
    可重入、跨執行緒的 GC guard：第一個進入者關閉循環 GC，最後一個離開者再恢復。
    - 僅用於極短、明確的關鍵區段（例如 openpyxl load/iter_rows、ElementTree feed）。
    - 預設不在離開時強制 gc.collect()。
    """
    global _guard_count, _guard_disabled_by_us, _last_activity_ts
    if not enabled:
        yield
        return

    entered = False
    try:
        with _guard_lock:
            if _guard_count == 0:
                # 記錄目前 GC 是否啟用，僅在啟用時由我們關閉
                _guard_disabled_by_us = gc.isenabled()
                if _guard_disabled_by_us:
                    gc.disable()
            _guard_count += 1
            entered = True
            # note activity on enter
            try:
                global _last_activity_ts
                _last_activity_ts = time.time()
            except Exception:
                pass
        # 臨界區
        yield
    finally:
        if entered:
            try:
                with _guard_lock:
                    _guard_count -= 1
                    # mark activity on exit
                    _last_activity_ts = time.time()
                    if _guard_count == 0:
                        if _guard_disabled_by_us:
                            gc.enable()
                            if do_collect:
                                try:
                                    gc.collect()
                                except Exception:
                                    pass
                        _guard_disabled_by_us = False
            except Exception:
                pass
