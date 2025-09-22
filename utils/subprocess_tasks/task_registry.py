"""
子進程任務註冊表 - 統一管理所有任務類型
已更新為使用統一子進程管理器，此文件保留以兼容舊版本
"""

def get_task_handler(task_type: str):
    """
    根據任務類型返回對應的處理函數
    
    Args:
        task_type: 任務類型
        
    Returns:
        處理函數
    """
    # 導入所有任務模組
    from .xml_tasks import extract_external_refs_task, read_meta_task
    from .compression_tasks import decompress_json_task
    
    # 動態導入 excel_tasks (新增)
    try:
        from .excel_tasks import (
            safe_load_workbook_task, 
            extract_all_formulas_task, 
            extract_cell_values_task,
            full_excel_scan_task
        )
    except ImportError:
        # 兼容性處理
        safe_load_workbook_task = None
        extract_all_formulas_task = None
        extract_cell_values_task = None
        full_excel_scan_task = None
    
    # 動態導入 baseline_tasks (新增)
    try:
        from .baseline_tasks import (
            load_baseline_task,
            save_baseline_task,
            compare_baseline_task,
            validate_baseline_task
        )
    except ImportError:
        load_baseline_task = None
        save_baseline_task = None
        compare_baseline_task = None
        validate_baseline_task = None
    
    task_handlers = {
        # XML 相關任務
        'extract_refs': extract_external_refs_task,
        'read_meta': read_meta_task,
        'decompress_json': decompress_json_task,
        
        # Excel 相關任務 (新增)
        'load_excel': safe_load_workbook_task,
        'extract_formulas': extract_all_formulas_task,
        'extract_cell_values': extract_cell_values_task,
        'full_excel_scan': full_excel_scan_task,
        
        # 基準線相關任務 (新增)
        'load_baseline': load_baseline_task,
        'save_baseline': save_baseline_task,
        'compare_baseline': compare_baseline_task,
        'validate_baseline': validate_baseline_task,
    }
    
    handler = task_handlers.get(task_type)
    if not handler:
        raise ValueError(f"不支援的任務類型: {task_type}")
    
    return handler

def get_available_task_types():
    """取得所有可用的任務類型"""
    return [
        # XML 相關
        'extract_refs',
        'read_meta', 
        'decompress_json',
        
        # Excel 相關
        'load_excel',
        'extract_formulas',
        'extract_cell_values', 
        'full_excel_scan',
        
        # 基準線相關
        'load_baseline',
        'save_baseline',
        'compare_baseline',
        'validate_baseline',
    ]