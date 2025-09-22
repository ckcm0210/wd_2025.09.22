"""
安全的 Excel 解析器 - 所有危險操作都通過子進程執行
替代原來的 core/excel_parser.py 中的危險函數
"""
import os
import time
import logging
from datetime import datetime
from typing import Dict, Any, Optional, List
import config.settings as settings
from utils.subprocess_manager import get_subprocess_manager, is_subprocess_enabled
from utils.cache import copy_to_cache

def dump_excel_cells_safe(path: str, show_sheet_detail: bool = True, silent: bool = False) -> Optional[Dict[str, Any]]:
    """
    安全的 Excel 儲存格資料提取 (子進程版本)
    替代原來的 dump_excel_cells_with_timeout
    
    Args:
        path: Excel 檔案路徑
        show_sheet_detail: 是否顯示工作表詳情
        silent: 是否靜默模式
        
    Returns:
        Excel 資料 {worksheet_name: {address: cell_data}}
    """
    if not is_subprocess_enabled():
        if not silent:
            print("❌ 子進程未啟用，無法安全處理 Excel 檔案")
        return None
    
    # 更新全局變數
    settings.current_processing_file = path
    settings.processing_start_time = time.time()
    
    try:
        if not silent:
            try:
                file_size_mb = os.path.getsize(path) / (1024 * 1024)
                print(f"   📊 檔案大小: {file_size_mb:.1f} MB")
            except Exception:
                pass
        
        # 複製到快取
        local_path = copy_to_cache(path, silent=silent)
        if not local_path or not os.path.exists(local_path):
            if not silent:
                print("   ❌ 無法建立快取副本，跳過此檔案")
            return None
        
        if not silent:
            print(f"   🚀 使用子進程安全模式處理")
        
        # 透過子進程管理器執行完整掃描
        manager = get_subprocess_manager()
        
        # 判斷是否為公式專用模式
        formula_only = bool(getattr(settings, 'FORMULA_ONLY_MODE', False))
        
        result = manager.scan_excel_complete(
            file_path=local_path,
            include_formulas=True,
            include_values=not formula_only,  # 公式模式下不包含值
            batch_size=int(getattr(settings, 'EXCEL_BATCH_SIZE', 10000))
        )
        
        excel_data = result.get('excel_data', {})
        
        if not excel_data:
            if not silent:
                print("   ⚠️ 子進程未返回資料")
            return {}
        
        # 統計資訊
        total_cells = sum(len(ws_data) for ws_data in excel_data.values())
        total_formulas = sum(
            sum(1 for cell in ws_data.values() if isinstance(cell, dict) and cell.get('formula'))
            for ws_data in excel_data.values()
        )
        
        if not silent:
            print(f"   ✅ Excel 讀取完成（子進程安全模式）")
            print(f"   📈 工作表: {len(excel_data)}, 儲存格: {total_cells}, 公式: {total_formulas}")
            
            if show_sheet_detail:
                for ws_name, ws_data in excel_data.items():
                    cell_count = len(ws_data)
                    formula_count = sum(1 for cell in ws_data.values() 
                                      if isinstance(cell, dict) and cell.get('formula'))
                    print(f"      工作表: {ws_name}（{cell_count} 儲存格，{formula_count} 公式）")
        
        return excel_data
        
    except Exception as e:
        if not silent:
            print(f"   ❌ 子進程處理失敗: {e}")
        logging.error(f"安全 Excel 處理失敗: {path}, 錯誤: {e}")
        return None

def get_excel_last_author_safe(path: str) -> Optional[str]:
    """
    安全的 Excel 最後作者取得 (子進程版本)
    替代原來的 get_excel_last_author
    
    Args:
        path: Excel 檔案路徑
        
    Returns:
        最後作者名稱，失敗返回 None
    """
    try:
        if not getattr(settings, 'ENABLE_LAST_AUTHOR_LOOKUP', True):
            return None
        
        if not is_subprocess_enabled():
            return None
        
        # 複製到快取
        local_path = copy_to_cache(path, silent=True)
        if not local_path or not os.path.exists(local_path):
            return None
        
        # 透過子進程讀取 metadata
        manager = get_subprocess_manager()
        result = manager.execute_task('read_meta', {
            'file_path': local_path
        })
        
        if result.get('success'):
            meta = result.get('meta', {})
            return meta.get('last_author')
        
        return None
        
    except Exception as e:
        logging.warning(f"安全取得最後作者失敗: {path}, 錯誤: {e}")
        return None

def extract_external_refs_safe(xlsx_path: str) -> Dict[int, str]:
    """
    安全的外部參照提取 (子進程版本)
    替代原來的 extract_external_refs
    
    Args:
        xlsx_path: Excel 檔案路徑
        
    Returns:
        外部參照映射 {index: path}
    """
    try:
        if not is_subprocess_enabled():
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print(f"[excel_parser_safe] 子進程未啟用，返回空映射")
            return {}
        
        manager = get_subprocess_manager()
        result = manager.execute_task('extract_refs', {
            'file_path': xlsx_path
        })
        
        if result.get('success'):
            external_refs = result.get('external_refs', {})
            
            # 修復鍵類型 (JSON 會將整數鍵轉為字符串)
            if external_refs and isinstance(external_refs, dict):
                try:
                    fixed_refs = {}
                    for k, v in external_refs.items():
                        try:
                            int_key = int(k)
                            fixed_refs[int_key] = v
                        except (ValueError, TypeError):
                            fixed_refs[k] = v
                    return fixed_refs
                except Exception:
                    return external_refs
            
            return external_refs
        
        return {}
        
    except Exception as e:
        if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
            print(f"[excel_parser_safe] 外部參照提取失敗: {e}")
        return {}

def load_excel_safe(file_path: str, **kwargs) -> Optional[Dict[str, Any]]:
    """
    安全的 Excel 載入 (子進程版本)
    僅返回工作簿基本資訊，不進行危險的主進程操作
    
    Args:
        file_path: Excel 檔案路徑
        **kwargs: openpyxl 載入參數
        
    Returns:
        工作簿資訊
    """
    try:
        if not is_subprocess_enabled():
            return None
        
        # 複製到快取
        local_path = copy_to_cache(file_path, silent=True)
        if not local_path or not os.path.exists(local_path):
            return None
        
        manager = get_subprocess_manager()
        result = manager.load_excel_safe(local_path, **kwargs)
        
        if result.get('success'):
            return result.get('workbook_info')
        
        return None
        
    except Exception as e:
        logging.error(f"安全 Excel 載入失敗: {file_path}, 錯誤: {e}")
        return None

def extract_formulas_only_safe(file_path: str, batch_size: int = 10000) -> Optional[Dict[str, Any]]:
    """
    安全的公式提取 (子進程版本)
    專用於公式專用模式
    
    Args:
        file_path: Excel 檔案路徑
        batch_size: 批次大小
        
    Returns:
        公式資料
    """
    try:
        if not is_subprocess_enabled():
            return None
        
        # 複製到快取
        local_path = copy_to_cache(file_path, silent=True)
        if not local_path or not os.path.exists(local_path):
            return None
        
        manager = get_subprocess_manager()
        result = manager.extract_formulas_safe(local_path, batch_size)
        
        if result.get('success'):
            return result.get('formulas_data')
        
        return None
        
    except Exception as e:
        logging.error(f"安全公式提取失敗: {file_path}, 錯誤: {e}")
        return None

def get_cell_values_safe(file_path: str, target_addresses: Dict[str, List[str]], 
                        use_data_only: bool = True) -> Optional[Dict[str, Dict[str, Any]]]:
    """
    安全的儲存格值提取 (子進程版本)
    
    Args:
        file_path: Excel 檔案路徑
        target_addresses: {worksheet_name: [address_list]}
        use_data_only: 是否使用 data_only 模式
        
    Returns:
        儲存格值 {worksheet_name: {address: value}}
    """
    try:
        if not is_subprocess_enabled():
            return None
        
        if not target_addresses:
            return {}
        
        # 複製到快取
        local_path = copy_to_cache(file_path, silent=True)
        if not local_path or not os.path.exists(local_path):
            return None
        
        manager = get_subprocess_manager()
        result = manager.execute_task('extract_cell_values', {
            'file_path': local_path,
            'target_addresses': target_addresses,
            'use_data_only': use_data_only
        })
        
        if result.get('success'):
            return result.get('cell_values')
        
        return None
        
    except Exception as e:
        logging.error(f"安全儲存格值提取失敗: {file_path}, 錯誤: {e}")
        return None

def pretty_formula_safe(formula, ref_map=None):
    """
    安全的公式美化函數 (保持原有邏輯，但用於子進程返回的資料)
    這個函數本身是安全的，不涉及 XML 解析
    """
    # 導入原有的 pretty_formula 實現
    from core.excel_parser import pretty_formula
    return pretty_formula(formula, ref_map)

def serialize_cell_value_safe(value):
    """
    安全的儲存格值序列化 (保持原有邏輯)
    這個函數本身是安全的
    """
    # 導入原有的實現
    from core.excel_parser import serialize_cell_value
    return serialize_cell_value(value)

# ============ 主進程替換接口 ============

def replace_dangerous_functions():
    """
    替換主進程中的危險函數為安全版本
    這個函數應該在程式啟動時調用
    """
    try:
        # 替換 core.excel_parser 中的危險函數
        import core.excel_parser as excel_parser
        
        # 備份原函數 (供除錯使用)
        if not hasattr(excel_parser, '_original_dump_excel_cells'):
            excel_parser._original_dump_excel_cells = excel_parser.dump_excel_cells_with_timeout
            excel_parser._original_get_excel_last_author = excel_parser.get_excel_last_author
            excel_parser._original_extract_external_refs = excel_parser.extract_external_refs
        
        # 替換為安全版本
        excel_parser.dump_excel_cells_with_timeout = dump_excel_cells_safe
        excel_parser.get_excel_last_author = get_excel_last_author_safe
        excel_parser.extract_external_refs = extract_external_refs_safe
        
        if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
            print("[excel_parser_safe] 危險函數已替換為安全版本")
        
    except Exception as e:
        logging.error(f"替換危險函數失敗: {e}")

def restore_original_functions():
    """
    恢復原始函數 (供除錯使用)
    """
    try:
        import core.excel_parser as excel_parser
        
        if hasattr(excel_parser, '_original_dump_excel_cells'):
            excel_parser.dump_excel_cells_with_timeout = excel_parser._original_dump_excel_cells
            excel_parser.get_excel_last_author = excel_parser._original_get_excel_last_author
            excel_parser.extract_external_refs = excel_parser._original_extract_external_refs
            
            if getattr(settings, 'SHOW_DEBUG_MESSAGES', False):
                print("[excel_parser_safe] 已恢復原始函數")
        
    except Exception as e:
        logging.error(f"恢復原始函數失敗: {e}")