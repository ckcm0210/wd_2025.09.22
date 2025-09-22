"""
統一的子進程管理器
整合所有子進程任務，提供統一的接口
"""
import os
import sys
import time
import json
import subprocess
import threading
from typing import Dict, Any, Optional, List, Union
from concurrent.futures import ThreadPoolExecutor, Future, TimeoutError
import config.settings as settings

class SubprocessManager:
    """
    統一子進程管理器
    整合 XML、Excel、基準線等所有子進程任務
    """
    
    def __init__(self):
        self.max_workers = max(1, int(getattr(settings, 'SUBPROCESS_MAX_WORKERS', 2)))
        self.timeout_sec = float(getattr(settings, 'SUBPROCESS_TIMEOUT_SEC', 30))
        self.safe_retry = bool(getattr(settings, 'SUBPROCESS_SAFE_RETRY', True))
        self.enabled = bool(getattr(settings, 'USE_SUBPROCESS', True))
        
        self.executor = None
        self.worker_counter = 0
        self.lock = threading.Lock()
        
        # 效能統計
        self.stats = {
            'total_tasks': 0,
            'successful_tasks': 0,
            'failed_tasks': 0,
            'timeout_tasks': 0,
            'safe_retry_tasks': 0,
            'total_time': 0.0,
            'avg_time': 0.0,
            'tasks_by_type': {},
            'last_reset': time.time()
        }
        
        if self.enabled:
            self.executor = ThreadPoolExecutor(
                max_workers=self.max_workers, 
                thread_name_prefix="subprocess-mgr"
            )
            self._debug(f"initialized max_workers={self.max_workers} timeout={self.timeout_sec}s")
    
    def _debug(self, message: str, level: int = 1):
        """輸出 debug 訊息"""
        try:
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False) and getattr(settings, 'DEBUG_LEVEL', 1) >= level:
                print(f"[subprocess-mgr] {message}")
        except Exception:
            pass
    
    def _update_stats(self, task_type: str, success: bool, elapsed: float, timeout: bool = False, safe_retry: bool = False):
        """更新統計資訊"""
        try:
            with self.lock:
                self.stats['total_tasks'] += 1
                self.stats['total_time'] += elapsed
                
                if success:
                    self.stats['successful_tasks'] += 1
                else:
                    self.stats['failed_tasks'] += 1
                
                if timeout:
                    self.stats['timeout_tasks'] += 1
                
                if safe_retry:
                    self.stats['safe_retry_tasks'] += 1
                
                # 按任務類型統計
                if task_type not in self.stats['tasks_by_type']:
                    self.stats['tasks_by_type'][task_type] = {'count': 0, 'success': 0, 'failed': 0}
                
                self.stats['tasks_by_type'][task_type]['count'] += 1
                if success:
                    self.stats['tasks_by_type'][task_type]['success'] += 1
                else:
                    self.stats['tasks_by_type'][task_type]['failed'] += 1
                
                # 計算平均時間
                if self.stats['total_tasks'] > 0:
                    self.stats['avg_time'] = self.stats['total_time'] / self.stats['total_tasks']
                
                # 定期輸出統計
                if self.stats['total_tasks'] % 20 == 0:
                    self._print_stats()
        except Exception:
            pass
    
    def _print_stats(self):
        """輸出統計資訊"""
        try:
            total = self.stats['total_tasks']
            success = self.stats['successful_tasks']
            failed = self.stats['failed_tasks']
            timeout = self.stats['timeout_tasks']
            avg_time = self.stats['avg_time']
            
            success_rate = (success / total * 100) if total > 0 else 0
            
            self._debug(f"stats: total={total} success={success}({success_rate:.1f}%) failed={failed} timeout={timeout} avg={avg_time:.2f}s")
            
            # 按任務類型顯示
            for task_type, stats in self.stats['tasks_by_type'].items():
                rate = (stats['success'] / stats['count'] * 100) if stats['count'] > 0 else 0
                self._debug(f"  {task_type}: {stats['count']} tasks, {rate:.1f}% success")
        except Exception:
            pass
    
    def get_stats(self) -> Dict[str, Any]:
        """取得統計資訊"""
        with self.lock:
            return self.stats.copy()
    
    def _get_worker_script_path(self) -> str:
        """取得統一子進程工作腳本路徑"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(current_dir, 'unified_subprocess_worker.py')
    
    def _run_subprocess_task(self, task_type: str, task_data: Dict[str, Any], 
                           worker_id: int, safe_mode: bool = False) -> Dict[str, Any]:
        """執行子進程任務"""
        worker_script = self._get_worker_script_path()
        if not os.path.exists(worker_script):
            raise FileNotFoundError(f"子進程工作腳本不存在: {worker_script}")
        
        # 準備任務輸入
        task_input = {
            'task_type': task_type,
            'task_data': task_data,
            'safe_mode': safe_mode,
            'worker_id': worker_id
        }
        
        file_path = task_data.get('file_path', 'unknown')
        mode_str = 'safe' if safe_mode else 'normal'
        
        self._debug(f"start worker_id={worker_id} task={task_type} file={os.path.basename(file_path)} mode={mode_str}")
        
        start_time = time.time()
        
        try:
            # 啟動子進程
            env = os.environ.copy()
            env.setdefault('PYTHONUTF8', '1')
            env.setdefault('PYTHONIOENCODING', 'utf-8')
            
            process = subprocess.Popen(
                [sys.executable, worker_script],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='replace',
                env=env,
                bufsize=0,  # 無緩衝，立即輸出
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0) if os.name == 'nt' else 0
            )
            
            # 發送任務資料
            input_json = json.dumps(task_input, ensure_ascii=False)
            stdout, stderr = process.communicate(input=input_json, timeout=self.timeout_sec)
            
            elapsed = time.time() - start_time
            
            if process.returncode == 0:
                # 成功完成
                try:
                    if not stdout or not stdout.strip():
                        raise RuntimeError("子進程無輸出")
                    result = json.loads(stdout.strip())
                    result_size = len(stdout.encode('utf-8', errors='ignore'))
                    self._debug(f"ok worker_id={worker_id} elapsed={elapsed:.1f}s result_size={result_size}bytes")
                    return result
                except json.JSONDecodeError as e:
                    self._debug(f"failed worker_id={worker_id} json_parse_error: {e}")
                    raise RuntimeError(f"子進程返回無效 JSON: {e}")
            else:
                # 子進程失敗
                self._debug(f"failed worker_id={worker_id} exit_code={process.returncode} elapsed={elapsed:.1f}s")
                if stderr:
                    self._debug(f"stderr worker_id={worker_id}: {stderr[:500]}")
                raise RuntimeError(f"子進程失敗 (exit_code={process.returncode})")
                
        except subprocess.TimeoutExpired:
            # 超時
            elapsed = time.time() - start_time
            self._debug(f"timeout worker_id={worker_id} after {elapsed:.1f}s")
            try:
                process.kill()
                process.wait(timeout=3)
            except Exception:
                pass
            raise TimeoutError(f"子進程超時 ({self.timeout_sec}s)")
        
        except Exception as e:
            elapsed = time.time() - start_time
            self._debug(f"error worker_id={worker_id} elapsed={elapsed:.1f}s error={type(e).__name__}: {e}")
            raise
    
    def execute_task(self, task_type: str, task_data: Dict[str, Any], 
                    safe_mode: Optional[bool] = None) -> Dict[str, Any]:
        """
        執行子進程任務
        
        Args:
            task_type: 任務類型
            task_data: 任務資料
            safe_mode: 是否使用安全模式 (None=自動判斷)
            
        Returns:
            任務結果
        """
        if not self.enabled:
            raise RuntimeError("子進程未啟用")
        
        if not self.executor:
            raise RuntimeError("子進程執行器未初始化")
        
        # 自動判斷安全模式
        if safe_mode is None:
            safe_mode = self._should_use_safe_mode(task_type, task_data)
        
        with self.lock:
            self.worker_counter += 1
            worker_id = self.worker_counter
        
        file_path = task_data.get('file_path', 'unknown')
        start_time = time.time()
        
        # 第一次嘗試：正常模式
        try:
            future = self.executor.submit(
                self._run_subprocess_task, 
                task_type, 
                task_data, 
                worker_id, 
                safe_mode=False
            )
            result = future.result(timeout=self.timeout_sec + 5)
            
            elapsed = time.time() - start_time
            self._update_stats(task_type, success=True, elapsed=elapsed)
            
            return result
            
        except TimeoutError as e:
            elapsed = time.time() - start_time
            self._debug(f"normal_mode_timeout worker_id={worker_id} elapsed={elapsed:.1f}s")
            self._update_stats(task_type, success=False, elapsed=elapsed, timeout=True)
            
            # 第二次嘗試：安全模式
            if self.safe_retry:
                return self._try_safe_mode(task_type, task_data, worker_id, file_path, start_time, [e])
            else:
                raise e
                
        except Exception as e:
            elapsed = time.time() - start_time
            self._debug(f"normal_mode_failed worker_id={worker_id} error={type(e).__name__}: {e}")
            self._update_stats(task_type, success=False, elapsed=elapsed)
            
            # 第二次嘗試：安全模式
            if self.safe_retry:
                return self._try_safe_mode(task_type, task_data, worker_id, file_path, start_time, [e])
            else:
                raise e
    
    def _try_safe_mode(self, task_type: str, task_data: Dict[str, Any], original_worker_id: int, 
                      file_path: str, original_start_time: float, errors: List[Exception]):
        """嘗試安全模式重試"""
        with self.lock:
            self.worker_counter += 1
            safe_worker_id = self.worker_counter
        
        self._debug(f"safe_retry_start worker_id={safe_worker_id} file={os.path.basename(file_path)}")
        
        safe_start_time = time.time()
        
        try:
            future = self.executor.submit(
                self._run_subprocess_task, 
                task_type, 
                task_data, 
                safe_worker_id, 
                safe_mode=True
            )
            result = future.result(timeout=self.timeout_sec + 10)
            
            safe_elapsed = time.time() - safe_start_time
            total_elapsed = time.time() - original_start_time
            self._update_stats(task_type, success=True, elapsed=total_elapsed, safe_retry=True)
            self._debug(f"safe_retry_ok worker_id={safe_worker_id} elapsed={safe_elapsed:.1f}s")
            
            return result
            
        except Exception as safe_e:
            safe_elapsed = time.time() - safe_start_time
            total_elapsed = time.time() - original_start_time
            
            is_timeout = isinstance(safe_e, TimeoutError)
            self._update_stats(task_type, success=False, elapsed=total_elapsed, timeout=is_timeout, safe_retry=True)
            
            self._debug(f"safe_retry_failed worker_id={safe_worker_id} error={type(safe_e).__name__}: {safe_e}")
            
            # 保存崩潰資訊
            self._save_crash_dump(task_type, task_data, errors + [safe_e])
            
            # 重新拋出原始錯誤
            raise errors[0]
    
    def _should_use_safe_mode(self, task_type: str, task_data: Dict[str, Any]) -> bool:
        """判斷是否應該使用安全模式"""
        
        # 檢查檔案大小
        file_path = task_data.get('file_path')
        if file_path and os.path.exists(file_path):
            try:
                file_size = os.path.getsize(file_path) / (1024 * 1024)  # MB
                if file_size > getattr(settings, 'SUBPROCESS_SAFE_MODE_THRESHOLD_MB', 50):
                    return True
            except Exception:
                pass
        
        # 檢查任務類型
        high_risk_tasks = [
            'full_excel_scan', 
            'extract_all_formulas', 
            'load_baseline'
        ]
        if task_type in high_risk_tasks:
            return True
        
        return False
    
    def _save_crash_dump(self, task_type: str, task_data: Dict[str, Any], errors: List[Exception]):
        """保存崩潰資訊"""
        try:
            log_folder = getattr(settings, 'LOG_FOLDER', '.')
            error_log_dir = os.path.join(log_folder, 'subprocess_crashes')
            os.makedirs(error_log_dir, exist_ok=True)
            
            timestamp = time.strftime('%Y%m%d_%H%M%S')
            crash_file = os.path.join(error_log_dir, f'subprocess_crash_{timestamp}.log')
            
            with open(crash_file, 'w', encoding='utf-8') as f:
                f.write(f"子進程崩潰報告\n")
                f.write(f"時間: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"任務類型: {task_type}\n")
                f.write(f"檔案路徑: {task_data.get('file_path', 'unknown')}\n")
                f.write(f"設定: max_workers={self.max_workers}, timeout={self.timeout_sec}s\n")
                f.write(f"\n錯誤詳情:\n")
                
                for i, error in enumerate(errors, 1):
                    f.write(f"\n嘗試 {i}: {type(error).__name__}\n")
                    f.write(f"訊息: {str(error)}\n")
            
            self._debug(f"crash_dump_saved path={crash_file}")
            
        except Exception as e:
            self._debug(f"crash_dump_save_failed error={e}")
    
    def shutdown(self):
        """關閉子進程管理器"""
        if self.executor:
            self._debug("shutting_down")
            self.executor.shutdown(wait=True, cancel_futures=True)
            self.executor = None
    
    # ============ 高級接口 ============
    
    def load_excel_safe(self, file_path: str, **kwargs) -> Dict[str, Any]:
        """安全載入 Excel 檔案 - 暫時不支援，回退到原有方法"""
        raise NotImplementedError("Excel 載入功能暫時不支援，請使用原有的 dump_excel_cells_with_timeout")
    
    def extract_formulas_safe(self, file_path: str, batch_size: int = 10000) -> Dict[str, Any]:
        """安全提取 Excel 公式 - 暫時不支援，回退到原有方法"""
        raise NotImplementedError("公式提取功能暫時不支援，請使用原有的 XML 子進程")
    
    def scan_excel_complete(self, file_path: str, include_formulas: bool = True, 
                           include_values: bool = True, batch_size: int = 10000) -> Dict[str, Any]:
        """完整 Excel 掃描（子進程）"""
        return self.execute_task('full_excel_scan', {
            'file_path': file_path,
            'include_formulas': include_formulas,
            'include_values': include_values,
            'batch_size': batch_size
        })
    
    def load_baseline_safe(self, baseline_path: str) -> Dict[str, Any]:
        """安全載入基準線（子進程）"""
        return self.execute_task('load_baseline', {
            'baseline_path': baseline_path
        })
    
    def save_baseline_safe(self, baseline_path: str, baseline_data: Dict[str, Any], 
                          compression_format: str = 'lz4') -> bool:
        """安全儲存基準線（子進程）"""
        result = self.execute_task('save_baseline', {
            'baseline_path': baseline_path,
            'baseline_data': baseline_data,
            'compression_format': compression_format
        })
        return result.get('success', False)
    
    def compare_baseline_safe(self, old_baseline: Dict[str, Any], 
                             new_data: Dict[str, Any]) -> Dict[str, Any]:
        """安全比較基準線（子進程）"""
        return self.execute_task('compare_baseline', {
            'old_baseline': old_baseline,
            'new_data': new_data
        })

# 全域單例
_subprocess_manager: Optional[SubprocessManager] = None

def get_subprocess_manager() -> SubprocessManager:
    """取得子進程管理器單例"""
    global _subprocess_manager
    if _subprocess_manager is None:
        _subprocess_manager = SubprocessManager()
    return _subprocess_manager

def is_subprocess_enabled() -> bool:
    """檢查子進程是否啟用"""
    return bool(getattr(settings, 'USE_SUBPROCESS', True))

def shutdown_subprocess_manager():
    """關閉子進程管理器"""
    global _subprocess_manager
    if _subprocess_manager:
        _subprocess_manager.shutdown()
        _subprocess_manager = None