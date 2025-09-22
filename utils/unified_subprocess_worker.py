"""
統一子進程工作腳本
整合所有子進程任務的執行入口
"""
import os
import sys
import json
import traceback
from typing import Dict, Any

def debug_print(message: str, worker_id: int = 0):
    """輸出 debug 訊息到 stderr"""
    print(f"[unified-worker-{worker_id}] {message}", file=sys.stderr, flush=True)

def execute_task(task_input: Dict[str, Any]) -> Dict[str, Any]:
    """
    統一任務執行入口
    
    Args:
        task_input: 任務輸入資料
        
    Returns:
        任務結果
    """
    task_type = task_input.get('task_type')
    task_data = task_input.get('task_data', {})
    safe_mode = task_input.get('safe_mode', False)
    worker_id = task_input.get('worker_id', 0)
    
    debug_print(f"execute_task start type={task_type} safe_mode={safe_mode}", worker_id)
    
    try:
        # ============ XML 相關任務 ============
        if task_type == 'extract_refs':
            # 使用同目錄的 xml_subproc_worker
            import xml_subproc_worker as xmlw
            file_path = task_data['file_path']
            external_refs = xmlw.extract_external_refs_task(file_path, safe_mode, worker_id)
            return {
                'success': True,
                'external_refs': external_refs,
                'worker_id': worker_id
            }
        
        elif task_type == 'read_meta':
            # 使用原有的 XML 子進程工作器中的函數  
            from utils.xml_subproc_worker import execute_task as xml_execute_task
            # 構建 XML 子進程的任務格式
            xml_task_input = {
                'task_type': 'read_meta',
                'task_data': task_data,
                'safe_mode': safe_mode,
                'worker_id': worker_id
            }
            result = xml_execute_task(xml_task_input)
            return result
        
        elif task_type == 'read_values':
            # 根據 engine 分發到不同的處理函數
            engine = task_data.get('engine', 'xml')
            file_path = task_data['file_path']
            
            if engine == 'openpyxl_scan':
                from subprocess_tasks.excel_tasks import extract_all_formulas_task
                scan = extract_all_formulas_task(file_path, safe_mode, worker_id)
                return {
                    'success': True,
                    'openpyxl_scan': scan,
                    'worker_id': worker_id
                }
            elif engine == 'data_only_values':
                from subprocess_tasks.excel_tasks import extract_cell_values_task
                coords = task_data.get('coords_by_sheet', {})
                data_vals = extract_cell_values_task(file_path, coords, safe_mode, worker_id, use_data_only=True)
                return {
                    'success': True,
                    'data_only_values': data_vals,
                    'worker_id': worker_id
                }
            elif engine in ('xml', 'polars_xml'):
                # 使用原有的 XML 子進程工作器中的函數
                from utils.xml_subproc_worker import read_values_task
                values_by_sheet = read_values_task(file_path, engine, safe_mode, worker_id)
                return {
                    'success': True,
                    'values_by_sheet': values_by_sheet,
                    'worker_id': worker_id
                }
            else:
                raise ValueError(f"不支援的值引擎: {engine}")
        
        # ============ Excel 相關任務 ============
        elif task_type == 'full_excel_scan':
            from subprocess_tasks.excel_tasks import full_excel_scan_task
            file_path = task_data['file_path']
            include_formulas = task_data.get('include_formulas', True)
            include_values = task_data.get('include_values', True)
            batch_size = task_data.get('batch_size', 10000)
            excel_data = full_excel_scan_task(file_path, safe_mode, worker_id, include_formulas, include_values, batch_size)
            return {
                'success': True,
                'excel_data': excel_data,
                'worker_id': worker_id
            }
        
        # ============ 基準線相關任務 ============
        elif task_type == 'load_baseline':
            from subprocess_tasks.baseline_tasks import load_baseline_task
            baseline_path = task_data['baseline_path']
            baseline_data = load_baseline_task(baseline_path, safe_mode, worker_id)
            return {
                'success': True,
                'baseline_data': baseline_data,
                'worker_id': worker_id
            }
        elif task_type == 'save_baseline':
            from subprocess_tasks.baseline_tasks import save_baseline_task
            baseline_path = task_data['baseline_path']
            baseline_data = task_data['baseline_data']
            compression_format = task_data.get('compression_format', 'lz4')
            success = save_baseline_task(baseline_path, baseline_data, compression_format, safe_mode, worker_id)
            return {
                'success': success,
                'worker_id': worker_id
            }
        elif task_type == 'compare_baseline':
            from subprocess_tasks.baseline_tasks import compare_baseline_task
            old_baseline = task_data['old_baseline']
            new_data = task_data['new_data']
            comparison_result = compare_baseline_task(old_baseline, new_data, safe_mode, worker_id)
            return {
                'success': True,
                'comparison_result': comparison_result,
                'worker_id': worker_id
            }
        elif task_type == 'validate_baseline':
            from subprocess_tasks.baseline_tasks import validate_baseline_task
            baseline_data = task_data['baseline_data']
            validation_result = validate_baseline_task(baseline_data, safe_mode, worker_id)
            return {
                'success': True,
                'validation_result': validation_result,
                'worker_id': worker_id
            }
        
        # ============ 壓縮相關任務 ============
        elif task_type == 'decompress_json':
            try:
                from subprocess_tasks.compression_tasks import decompress_json_task
                file_path = task_data['file_path']
                decompressed_data = decompress_json_task(file_path, safe_mode, worker_id)
                return {
                    'success': True,
                    'decompressed_data': decompressed_data,
                    'worker_id': worker_id
                }
            except ImportError:
                # 如果新的壓縮任務模組不存在，使用基本解壓縮
                import os
                import json
                import gzip
                file_path = task_data['file_path']
                
                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"檔案不存在: {file_path}")
                
                try:
                    with gzip.open(file_path, 'rt', encoding='utf-8') as f:
                        data = json.load(f)
                    return {
                        'success': True,
                        'decompressed_data': data,
                        'worker_id': worker_id
                    }
                except Exception as e:
                    raise RuntimeError(f"解壓縮失敗: {e}")
        
        else:
            raise ValueError(f"不支援的任務類型: {task_type}")
    
    except Exception as e:
        debug_print(f"execute_task failed: {type(e).__name__}: {e}", worker_id)
        return {
            'success': False,
            'error': str(e),
            'error_type': type(e).__name__,
            'traceback': traceback.format_exc(),
            'worker_id': worker_id
        }

def main():
    """子進程主要入口點"""
    try:
        # 強制 I/O 編碼為 UTF-8
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='strict')
        except Exception:
            pass
        try:
            sys.stderr.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
        
        # 從 stdin 讀取任務輸入
        input_data = sys.stdin.read()
        if not input_data.strip():
            raise ValueError("未收到任務輸入")
        
        task_input = json.loads(input_data)
        
        # 執行任務
        result = execute_task(task_input)
        
        # 輸出結果到 stdout
        output = json.dumps(result, ensure_ascii=True, indent=None, separators=(',', ':'))
        print(output, flush=True)
        
        # 根據執行結果決定退出碼
        if result.get('success', False):
            sys.exit(0)
        else:
            sys.exit(1)
        
    except json.JSONDecodeError as e:
        error_result = {
            'success': False,
            'error': f'JSON 解析失敗: {e}',
            'error_type': 'JSONDecodeError',
            'traceback': traceback.format_exc()
        }
        print(json.dumps(error_result, ensure_ascii=True), flush=True)
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            'success': False,
            'error': str(e),
            'error_type': type(e).__name__,
            'traceback': traceback.format_exc()
        }
        print(json.dumps(error_result, ensure_ascii=True), flush=True)
        sys.exit(1)

if __name__ == '__main__':
    main()