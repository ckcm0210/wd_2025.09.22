"""
安全的基準線管理器 - 所有 JSON 和壓縮操作都通過子進程執行
替代原來的 core/baseline.py 中的危險函數
"""
import os
import time
import logging
from datetime import datetime
from typing import Dict, Any, Optional
import config.settings as settings
from utils.subprocess_manager import get_subprocess_manager, is_subprocess_enabled

def load_baseline_safe(base_key: str) -> Optional[Dict[str, Any]]:
    """
    安全的基準線載入 (子進程版本)
    替代原來的 load_baseline
    
    Args:
        base_key: 基準線鍵值
        
    Returns:
        基準線資料，失敗返回 None
    """
    try:
        if not is_subprocess_enabled():
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print("[baseline_safe] 子進程未啟用，無法安全載入基準線")
            return None
        
        # 獲取基準線檔案路徑
        from core.baseline import baseline_file_path
        baseline_path = baseline_file_path(base_key)
        
        if not os.path.exists(baseline_path):
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print(f"[baseline_safe] 基準線檔案不存在: {baseline_path}")
            return None
        
        # 透過子進程載入
        manager = get_subprocess_manager()
        result = manager.load_baseline_safe(baseline_path)
        
        if result.get('success'):
            baseline_data = result.get('baseline_data', {})
            
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                cells_count = 0
                if 'cells' in baseline_data:
                    for ws_data in baseline_data['cells'].values():
                        if isinstance(ws_data, dict):
                            cells_count += len(ws_data)
                print(f"[baseline_safe] 基準線載入成功: {cells_count} 儲存格")
            
            return baseline_data
        
        if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
            error = result.get('error', '未知錯誤')
            print(f"[baseline_safe] 基準線載入失敗: {error}")
        
        return None
        
    except Exception as e:
        logging.error(f"安全基準線載入失敗: {base_key}, 錯誤: {e}")
        return None

def save_baseline_safe(base_key: str, baseline_data: Dict[str, Any]) -> bool:
    """
    安全的基準線儲存 (子進程版本)
    替代原來的 save_baseline
    
    Args:
        base_key: 基準線鍵值
        baseline_data: 基準線資料
        
    Returns:
        是否成功
    """
    try:
        if not is_subprocess_enabled():
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print("[baseline_safe] 子進程未啟用，無法安全儲存基準線")
            return False
        
        # 獲取基準線檔案路徑
        from core.baseline import baseline_file_path
        baseline_path = baseline_file_path(base_key)
        
        # 確保目錄存在
        os.makedirs(os.path.dirname(baseline_path), exist_ok=True)
        
        # 透過子進程儲存
        manager = get_subprocess_manager()
        compression_format = getattr(settings, 'DEFAULT_COMPRESSION_FORMAT', 'lz4')
        
        success = manager.save_baseline_safe(baseline_path, baseline_data, compression_format)
        
        if success:
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                cells_count = 0
                if 'cells' in baseline_data:
                    for ws_data in baseline_data['cells'].values():
                        if isinstance(ws_data, dict):
                            cells_count += len(ws_data)
                print(f"[baseline_safe] 基準線儲存成功: {cells_count} 儲存格")
        else:
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print(f"[baseline_safe] 基準線儲存失敗")
        
        return success
        
    except Exception as e:
        logging.error(f"安全基準線儲存失敗: {base_key}, 錯誤: {e}")
        return False

def create_baseline_for_files_safe(file_paths: list) -> bool:
    """
    安全的批量基準線建立 (子進程版本)
    替代原來的 create_baseline_for_files_robust
    
    Args:
        file_paths: Excel 檔案路徑列表
        
    Returns:
        是否全部成功
    """
    if not file_paths:
        return True
    
    if not is_subprocess_enabled():
        print("❌ 子進程未啟用，無法安全建立基準線")
        return False
    
    print(f"📊 安全建立基準線：{len(file_paths)} 個檔案")
    
    success_count = 0
    total_files = len(file_paths)
    
    for i, file_path in enumerate(file_paths, 1):
        try:
            print(f"   ({i}/{total_files}) 處理: {os.path.basename(file_path)}")
            
            if not os.path.exists(file_path):
                print(f"      ❌ 檔案不存在")
                continue
            
            # 透過安全的 Excel 解析器取得資料
            from core.excel_parser_safe import dump_excel_cells_safe, get_excel_last_author_safe
            
            current_data = dump_excel_cells_safe(file_path, show_sheet_detail=False, silent=True)
            if current_data is None:
                print(f"      ❌ 無法讀取檔案")
                continue
            
            # 取得最後作者
            last_author = get_excel_last_author_safe(file_path)
            
            # 建立基準線資料
            from utils.helpers import _baseline_key_for_path, get_file_mtime
            from core.excel_parser import hash_excel_content
            
            base_key = _baseline_key_for_path(file_path)
            file_mtime = get_file_mtime(file_path)
            
            baseline_data = {
                "last_author": last_author,
                "content_hash": hash_excel_content(current_data),
                "cells": current_data,
                "timestamp": datetime.now().isoformat(),
                "file_mtime_str": file_mtime,
                "source_mtime": os.path.getmtime(file_path),
                "source_size": os.path.getsize(file_path)
            }
            
            # 安全儲存基準線
            if save_baseline_safe(base_key, baseline_data):
                success_count += 1
                print(f"      ✅ 基準線建立成功")
            else:
                print(f"      ❌ 基準線儲存失敗")
            
        except Exception as e:
            print(f"      ❌ 處理失敗: {e}")
            logging.error(f"基準線建立失敗: {file_path}, 錯誤: {e}")
            continue
    
    print(f"✅ 基準線建立完成：{success_count}/{total_files} 成功")
    return success_count == total_files

def compare_with_baseline_safe(file_path: str, base_key: str) -> Optional[Dict[str, Any]]:
    """
    安全的基準線比較 (子進程版本)
    
    Args:
        file_path: Excel 檔案路徑
        base_key: 基準線鍵值
        
    Returns:
        比較結果
    """
    try:
        if not is_subprocess_enabled():
            return None
        
        # 載入基準線
        old_baseline = load_baseline_safe(base_key)
        if not old_baseline:
            return None
        
        # 讀取當前資料
        from core.excel_parser_safe import dump_excel_cells_safe
        current_data = dump_excel_cells_safe(file_path, show_sheet_detail=False, silent=True)
        if current_data is None:
            return None
        
        # 透過子進程比較
        manager = get_subprocess_manager()
        result = manager.compare_baseline_safe(old_baseline, current_data)
        
        if result.get('success'):
            return result.get('comparison_result')
        
        return None
        
    except Exception as e:
        logging.error(f"安全基準線比較失敗: {file_path}, 錯誤: {e}")
        return None

def validate_baseline_safe(base_key: str) -> Optional[Dict[str, Any]]:
    """
    安全的基準線驗證 (子進程版本)
    
    Args:
        base_key: 基準線鍵值
        
    Returns:
        驗證結果
    """
    try:
        if not is_subprocess_enabled():
            return None
        
        # 載入基準線
        baseline_data = load_baseline_safe(base_key)
        if not baseline_data:
            return {
                'is_valid': False,
                'errors': ['基準線載入失敗'],
                'warnings': [],
                'statistics': {}
            }
        
        # 透過子進程驗證
        manager = get_subprocess_manager()
        result = manager.execute_task('validate_baseline', {
            'baseline_data': baseline_data
        })
        
        if result.get('success'):
            return result.get('validation_result')
        
        return {
            'is_valid': False,
            'errors': [result.get('error', '驗證失敗')],
            'warnings': [],
            'statistics': {}
        }
        
    except Exception as e:
        logging.error(f"安全基準線驗證失敗: {base_key}, 錯誤: {e}")
        return {
            'is_valid': False,
            'errors': [f'驗證過程出錯: {e}'],
            'warnings': [],
            'statistics': {}
        }

def get_baseline_statistics_safe(base_key: str) -> Optional[Dict[str, Any]]:
    """
    安全的基準線統計資訊取得
    
    Args:
        base_key: 基準線鍵值
        
    Returns:
        統計資訊
    """
    try:
        validation_result = validate_baseline_safe(base_key)
        if validation_result:
            return validation_result.get('statistics', {})
        return None
        
    except Exception as e:
        logging.error(f"基準線統計取得失敗: {base_key}, 錯誤: {e}")
        return None

def cleanup_old_baselines_safe(days_old: int = 30) -> int:
    """
    安全的舊基準線清理
    
    Args:
        days_old: 清理多少天前的基準線
        
    Returns:
        清理的檔案數量
    """
    try:
        # 這個函數本身是安全的，不涉及 JSON 解析
        # 只是檔案系統操作
        from core.baseline import get_cache_folder
        
        cache_folder = get_cache_folder()
        if not os.path.exists(cache_folder):
            return 0
        
        import time
        cutoff_time = time.time() - (days_old * 24 * 60 * 60)
        cleaned_count = 0
        
        for root, dirs, files in os.walk(cache_folder):
            for file in files:
                if file.endswith(('.lz4', '.zst', '.gz', '.json')):
                    file_path = os.path.join(root, file)
                    try:
                        if os.path.getmtime(file_path) < cutoff_time:
                            os.remove(file_path)
                            cleaned_count += 1
                    except Exception:
                        continue
        
        if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
            print(f"[baseline_safe] 清理了 {cleaned_count} 個舊基準線檔案")
        
        return cleaned_count
        
    except Exception as e:
        logging.error(f"舊基準線清理失敗: {e}")
        return 0

# ============ 主進程替換接口 ============

def replace_dangerous_baseline_functions():
    """
    替換主進程中的危險基準線函數為安全版本
    這個函數應該在程式啟動時調用
    """
    try:
        # 替換 core.baseline 中的危險函數
        import core.baseline as baseline
        
        # 備份原函數 (供除錯使用)
        if not hasattr(baseline, '_original_load_baseline'):
            baseline._original_load_baseline = baseline.load_baseline
            baseline._original_save_baseline = baseline.save_baseline
            baseline._original_create_baseline_for_files_robust = baseline.create_baseline_for_files_robust
        
        # 替換為安全版本
        baseline.load_baseline = load_baseline_safe
        baseline.save_baseline = save_baseline_safe
        baseline.create_baseline_for_files_robust = create_baseline_for_files_safe
        
        if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
            print("[baseline_safe] 危險基準線函數已替換為安全版本")
        
    except Exception as e:
        logging.error(f"替換危險基準線函數失敗: {e}")

def restore_original_baseline_functions():
    """
    恢復原始基準線函數 (供除錯使用)
    """
    try:
        import core.baseline as baseline
        
        if hasattr(baseline, '_original_load_baseline'):
            baseline.load_baseline = baseline._original_load_baseline
            baseline.save_baseline = baseline._original_save_baseline
            baseline.create_baseline_for_files_robust = baseline._original_create_baseline_for_files_robust
            
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print("[baseline_safe] 已恢復原始基準線函數")
        
    except Exception as e:
        logging.error(f"恢復原始基準線函數失敗: {e}")

# ============ 兼容性函數 ============

def get_baseline_file_with_extension_safe(base_key: str) -> Optional[str]:
    """
    安全版本的基準線檔案查找
    這個函數本身是安全的，只涉及檔案系統操作
    """
    try:
        from core.baseline import get_baseline_file_with_extension
        return get_baseline_file_with_extension(base_key)
    except Exception as e:
        logging.error(f"基準線檔案查找失敗: {base_key}, 錯誤: {e}")
        return None

def baseline_file_path_safe(base_key: str) -> str:
    """
    安全版本的基準線檔案路徑取得
    這個函數本身是安全的
    """
    from core.baseline import baseline_file_path
    return baseline_file_path(base_key)