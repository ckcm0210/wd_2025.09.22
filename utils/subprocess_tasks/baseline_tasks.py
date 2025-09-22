"""
基準線相關的子進程任務
將 JSON 解壓縮和基準線處理移到子進程
"""
import os
import sys
import json
import gzip
# 延遲載入壓縮模組，避免在沒有安裝時於 import 階段失敗
# 於函數內按需載入 lz4.frame / zstandard
from typing import Dict, Any, Optional

def debug_print(message: str, worker_id: int = 0):
    """輸出 debug 訊息到 stderr"""
    print(f"[baseline-worker-{worker_id}] {message}", file=sys.stderr, flush=True)

def load_baseline_task(baseline_path: str, safe_mode: bool = False, worker_id: int = 0) -> Dict[str, Any]:
    """
    載入基準線檔案 (子進程版本)
    支援各種壓縮格式的安全解壓縮
    
    Args:
        baseline_path: 基準線檔案路徑
        safe_mode: 是否使用安全模式
        worker_id: 工作者 ID
        
    Returns:
        基準線資料
    """
    debug_print(f"load_baseline start file={os.path.basename(baseline_path)}", worker_id)
    
    if not os.path.exists(baseline_path):
        debug_print("baseline file not found", worker_id)
        return {}
    
    try:
        # 檢測壓縮格式
        compression_format = _detect_compression_format(baseline_path)
        debug_print(f"detected format: {compression_format}", worker_id)
        
        # 讀取並解壓縮
        with open(baseline_path, 'rb') as f:
            compressed_data = f.read()
        
        if not compressed_data:
            debug_print("empty baseline file", worker_id)
            return {}
        
        # 解壓縮
        try:
            if compression_format == 'lz4':
                try:
                    import lz4.frame as _lz4f
                except ImportError as _ie:
                    if safe_mode:
                        debug_print("lz4 not installed; safe_mode -> return empty", worker_id)
                        return {}
                    raise RuntimeError("LZ4 模組未安裝") from _ie
                json_data = _lz4f.decompress(compressed_data).decode('utf-8')
            elif compression_format == 'zstd':
                try:
                    import zstandard as _zstd
                except ImportError as _ie:
                    if safe_mode:
                        debug_print("zstandard not installed; safe_mode -> return empty", worker_id)
                        return {}
                    raise RuntimeError("Zstandard 模組未安裝") from _ie
                decompressor = _zstd.ZstdDecompressor()
                json_data = decompressor.decompress(compressed_data).decode('utf-8')
            elif compression_format == 'gzip':
                json_data = gzip.decompress(compressed_data).decode('utf-8')
            else:
                # 未壓縮的 JSON
                json_data = compressed_data.decode('utf-8')
                
        except Exception as decomp_e:
            debug_print(f"decompression failed: {decomp_e}", worker_id)
            if safe_mode:
                return {}
            raise RuntimeError(f"解壓縮失敗: {decomp_e}")
        
        # 解析 JSON
        try:
            baseline_data = json.loads(json_data)
            debug_print(f"JSON parsed successfully, size: {len(json_data)} chars", worker_id)
            
            # 驗證基準線資料結構
            if not isinstance(baseline_data, dict):
                debug_print("invalid baseline structure (not dict)", worker_id)
                if safe_mode:
                    return {}
                raise ValueError("基準線格式錯誤：不是字典結構")
            
            # 統計資訊
            cells_count = 0
            if 'cells' in baseline_data and isinstance(baseline_data['cells'], dict):
                for ws_data in baseline_data['cells'].values():
                    if isinstance(ws_data, dict):
                        cells_count += len(ws_data)
            
            debug_print(f"baseline loaded successfully, cells: {cells_count}", worker_id)
            return baseline_data
            
        except json.JSONDecodeError as json_e:
            debug_print(f"JSON parsing failed: {json_e}", worker_id)
            if safe_mode:
                return {}
            raise RuntimeError(f"JSON 解析失敗: {json_e}")
            
    except Exception as e:
        debug_print(f"load_baseline failed: {type(e).__name__}: {e}", worker_id)
        if safe_mode:
            return {}
        else:
            raise RuntimeError(f"基準線載入失敗: {e}")

def save_baseline_task(baseline_path: str, baseline_data: Dict[str, Any], 
                      compression_format: str = 'lz4', safe_mode: bool = False, 
                      worker_id: int = 0) -> bool:
    """
    儲存基準線檔案 (子進程版本)
    
    Args:
        baseline_path: 基準線檔案路徑
        baseline_data: 基準線資料
        compression_format: 壓縮格式 ('lz4', 'zstd', 'gzip', 'none')
        safe_mode: 是否使用安全模式
        worker_id: 工作者 ID
        
    Returns:
        是否成功
    """
    debug_print(f"save_baseline start file={os.path.basename(baseline_path)} format={compression_format}", worker_id)
    
    try:
        # 確保目錄存在
        os.makedirs(os.path.dirname(baseline_path), exist_ok=True)
        
        # 序列化為 JSON
        try:
            json_data = json.dumps(baseline_data, ensure_ascii=False, separators=(',', ':'))
            debug_print(f"JSON serialized, size: {len(json_data)} chars", worker_id)
        except Exception as json_e:
            debug_print(f"JSON serialization failed: {json_e}", worker_id)
            if safe_mode:
                return False
            raise RuntimeError(f"JSON 序列化失敗: {json_e}")
        
        # 壓縮
        try:
            if compression_format == 'lz4':
                try:
                    import lz4.frame as _lz4f
                except ImportError as _ie:
                    if safe_mode:
                        debug_print("lz4 not installed; safe_mode -> return False", worker_id)
                        return False
                    raise RuntimeError("LZ4 模組未安裝") from _ie
                compressed_data = _lz4f.compress(json_data.encode('utf-8'))
            elif compression_format == 'zstd':
                try:
                    import zstandard as _zstd
                except ImportError as _ie:
                    if safe_mode:
                        debug_print("zstandard not installed; safe_mode -> return False", worker_id)
                        return False
                    raise RuntimeError("Zstandard 模組未安裝") from _ie
                compressor = _zstd.ZstdCompressor()
                compressed_data = compressor.compress(json_data.encode('utf-8'))
            elif compression_format == 'gzip':
                compressed_data = gzip.compress(json_data.encode('utf-8'))
            else:
                # 不壓縮
                compressed_data = json_data.encode('utf-8')
                
            debug_print(f"compressed, ratio: {len(compressed_data)}/{len(json_data.encode('utf-8'))}", worker_id)
            
        except Exception as comp_e:
            debug_print(f"compression failed: {comp_e}", worker_id)
            if safe_mode:
                return False
            raise RuntimeError(f"壓縮失敗: {comp_e}")
        
        # 寫入檔案
        try:
            # 先寫入臨時檔案，再移動 (原子操作)
            temp_path = baseline_path + '.tmp'
            with open(temp_path, 'wb') as f:
                f.write(compressed_data)
                f.flush()
                os.fsync(f.fileno())  # 強制寫入磁碟
            
            # 原子移動
            os.replace(temp_path, baseline_path)
            debug_print("baseline saved successfully", worker_id)
            return True
            
        except Exception as write_e:
            debug_print(f"file write failed: {write_e}", worker_id)
            # 清理臨時檔案
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
            if safe_mode:
                return False
            raise RuntimeError(f"檔案寫入失敗: {write_e}")
            
    except Exception as e:
        debug_print(f"save_baseline failed: {type(e).__name__}: {e}", worker_id)
        if safe_mode:
            return False
        else:
            raise RuntimeError(f"基準線儲存失敗: {e}")

def compare_baseline_task(old_baseline: Dict[str, Any], new_data: Dict[str, Any], 
                         safe_mode: bool = False, worker_id: int = 0) -> Dict[str, Any]:
    """
    比較基準線和新資料 (子進程版本)
    
    Args:
        old_baseline: 舊基準線資料
        new_data: 新的 Excel 資料
        safe_mode: 是否使用安全模式
        worker_id: 工作者 ID
        
    Returns:
        比較結果
    """
    debug_print("compare_baseline start", worker_id)
    
    try:
        old_cells = old_baseline.get('cells', {})
        new_cells = new_data or {}
        
        # 統計資訊
        old_sheets = set(old_cells.keys())
        new_sheets = set(new_cells.keys())
        all_sheets = old_sheets | new_sheets
        
        comparison_result = {
            'has_changes': False,
            'sheets_added': list(new_sheets - old_sheets),
            'sheets_removed': list(old_sheets - new_sheets),
            'sheets_modified': [],
            'total_changes': 0,
            'changes_by_sheet': {}
        }
        
        debug_print(f"comparing {len(all_sheets)} sheets", worker_id)
        
        # 比較每個工作表
        for sheet_name in all_sheets:
            old_sheet = old_cells.get(sheet_name, {})
            new_sheet = new_cells.get(sheet_name, {})
            
            sheet_changes = _compare_worksheet(old_sheet, new_sheet, sheet_name, worker_id, safe_mode)
            
            if sheet_changes['change_count'] > 0:
                comparison_result['has_changes'] = True
                comparison_result['sheets_modified'].append(sheet_name)
                comparison_result['total_changes'] += sheet_changes['change_count']
                comparison_result['changes_by_sheet'][sheet_name] = sheet_changes
        
        # 移除工作表也算變更
        if comparison_result['sheets_added'] or comparison_result['sheets_removed']:
            comparison_result['has_changes'] = True
        
        debug_print(f"comparison completed, changes: {comparison_result['total_changes']}", worker_id)
        return comparison_result
        
    except Exception as e:
        debug_print(f"compare_baseline failed: {type(e).__name__}: {e}", worker_id)
        if safe_mode:
            return {'has_changes': False, 'error': str(e)}
        else:
            raise RuntimeError(f"基準線比較失敗: {e}")

def _compare_worksheet(old_sheet: Dict[str, Any], new_sheet: Dict[str, Any], 
                      sheet_name: str, worker_id: int, safe_mode: bool) -> Dict[str, Any]:
    """比較單個工作表"""
    
    try:
        old_addrs = set(old_sheet.keys())
        new_addrs = set(new_sheet.keys())
        all_addrs = old_addrs | new_addrs
        
        changes = {
            'change_count': 0,
            'cells_added': [],
            'cells_removed': [],
            'cells_modified': [],
            'formula_changes': 0,
            'value_changes': 0
        }
        
        for addr in all_addrs:
            old_cell = old_sheet.get(addr, {})
            new_cell = new_sheet.get(addr, {})
            
            # 儲存格新增
            if not old_cell and new_cell:
                changes['cells_added'].append(addr)
                changes['change_count'] += 1
                continue
            
            # 儲存格刪除
            if old_cell and not new_cell:
                changes['cells_removed'].append(addr)
                changes['change_count'] += 1
                continue
            
            # 儲存格修改
            if old_cell != new_cell:
                change_detail = _analyze_cell_change(old_cell, new_cell, addr)
                changes['cells_modified'].append(change_detail)
                changes['change_count'] += 1
                
                if change_detail.get('formula_changed'):
                    changes['formula_changes'] += 1
                if change_detail.get('value_changed'):
                    changes['value_changes'] += 1
        
        return changes
        
    except Exception as e:
        debug_print(f"worksheet comparison failed for '{sheet_name}': {e}", worker_id)
        if safe_mode:
            return {'change_count': 0, 'error': str(e)}
        else:
            raise

def _analyze_cell_change(old_cell: Dict[str, Any], new_cell: Dict[str, Any], addr: str) -> Dict[str, Any]:
    """分析儲存格變更詳情"""
    
    change_detail = {
        'address': addr,
        'formula_changed': False,
        'value_changed': False,
        'old_formula': old_cell.get('formula'),
        'new_formula': new_cell.get('formula'),
        'old_value': old_cell.get('value'),
        'new_value': new_cell.get('value')
    }
    
    # 檢查公式變更
    if old_cell.get('formula') != new_cell.get('formula'):
        change_detail['formula_changed'] = True
    
    # 檢查值變更
    old_val = old_cell.get('cached_value') if old_cell.get('cached_value') is not None else old_cell.get('value')
    new_val = new_cell.get('cached_value') if new_cell.get('cached_value') is not None else new_cell.get('value')
    
    if old_val != new_val:
        change_detail['value_changed'] = True
    
    return change_detail

def _detect_compression_format(file_path: str) -> str:
    """檢測檔案的壓縮格式"""
    
    try:
        with open(file_path, 'rb') as f:
            header = f.read(16)  # 讀取前16位元組
        
        if not header:
            return 'none'
        
        # LZ4 magic number: 0x184D2204
        if header.startswith(b'\x04"M\x18'):
            return 'lz4'
        
        # Zstandard magic number: 0xFD2FB5?? (28 B5 2F FD)
        if header.startswith(b'\x28\xb5\x2f\xfd'):
            return 'zstd'
        
        # Gzip magic number: 1f 8b
        if header.startswith(b'\x1f\x8b'):
            return 'gzip'
        
        # 嘗試解析為 JSON (未壓縮)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                f.read(1)  # 嘗試讀取一個字符
            return 'none'
        except Exception:
            pass
        
        return 'unknown'
        
    except Exception:
        return 'unknown'

def validate_baseline_task(baseline_data: Dict[str, Any], safe_mode: bool = False, 
                          worker_id: int = 0) -> Dict[str, Any]:
    """
    驗證基準線資料完整性 (子進程版本)
    
    Args:
        baseline_data: 基準線資料
        safe_mode: 是否使用安全模式
        worker_id: 工作者 ID
        
    Returns:
        驗證結果
    """
    debug_print("validate_baseline start", worker_id)
    
    try:
        validation_result = {
            'is_valid': True,
            'errors': [],
            'warnings': [],
            'statistics': {}
        }
        
        # 檢查基本結構
        if not isinstance(baseline_data, dict):
            validation_result['is_valid'] = False
            validation_result['errors'].append("基準線不是字典結構")
            return validation_result
        
        # 檢查必要欄位
        required_fields = ['cells', 'timestamp']
        for field in required_fields:
            if field not in baseline_data:
                validation_result['warnings'].append(f"缺少欄位: {field}")
        
        # 檢查 cells 結構
        cells = baseline_data.get('cells', {})
        if not isinstance(cells, dict):
            validation_result['is_valid'] = False
            validation_result['errors'].append("cells 不是字典結構")
            return validation_result
        
        # 統計資訊
        total_sheets = len(cells)
        total_cells = 0
        formula_cells = 0
        value_cells = 0
        
        for sheet_name, sheet_data in cells.items():
            if not isinstance(sheet_data, dict):
                validation_result['errors'].append(f"工作表 '{sheet_name}' 資料不是字典結構")
                continue
            
            for addr, cell_data in sheet_data.items():
                total_cells += 1
                
                if isinstance(cell_data, dict):
                    if cell_data.get('formula'):
                        formula_cells += 1
                    if cell_data.get('value') is not None:
                        value_cells += 1
                else:
                    validation_result['warnings'].append(f"儲存格 {sheet_name}!{addr} 資料格式異常")
        
        validation_result['statistics'] = {
            'total_sheets': total_sheets,
            'total_cells': total_cells,
            'formula_cells': formula_cells,
            'value_cells': value_cells
        }
        
        debug_print(f"validation completed, cells: {total_cells}, valid: {validation_result['is_valid']}", worker_id)
        return validation_result
        
    except Exception as e:
        debug_print(f"validate_baseline failed: {type(e).__name__}: {e}", worker_id)
        if safe_mode:
            return {
                'is_valid': False,
                'errors': [f"驗證過程出錯: {e}"],
                'warnings': [],
                'statistics': {}
            }
        else:
            raise RuntimeError(f"基準線驗證失敗: {e}")