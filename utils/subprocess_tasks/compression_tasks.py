"""
壓縮相關的子進程任務
"""
import os
import sys
import json
import gzip

def debug_print(message: str, worker_id: int = 0):
    """輸出 debug 訊息到 stderr"""
    print(f"[compression-worker-{worker_id}] {message}", file=sys.stderr, flush=True)

def decompress_json_task(file_path: str, safe_mode: bool = False, worker_id: int = 0):
    """
    子進程JSON解壓與解析任務
    
    Args:
        file_path: 壓縮檔案路徑
        safe_mode: 是否使用安全模式
        worker_id: 工作者ID
        
    Returns:
        解析後的JSON資料
    """
    debug_print(f"decompress_json start file={os.path.basename(file_path)}", worker_id)
    
    try:
        # 檢測壓縮格式
        format_type = None
        if file_path.endswith('.lz4'):
            format_type = 'lz4'
        elif file_path.endswith('.zst'):
            format_type = 'zstd'
        elif file_path.endswith('.gz'):
            format_type = 'gzip'
        
        # 讀取壓縮檔案
        with open(file_path, 'rb') as f:
            compressed_data = f.read()
        
        # 解壓縮
        if format_type == 'lz4':
            try:
                import lz4.frame
                json_data = lz4.frame.decompress(compressed_data).decode('utf-8')
            except ImportError:
                if safe_mode:
                    debug_print("LZ4 模組未安裝，安全模式返回空", worker_id)
                    return {}
                raise RuntimeError("LZ4 模組未安裝")
        elif format_type == 'zstd':
            try:
                import zstandard as zstd
                decompressor = zstd.ZstdDecompressor()
                json_data = decompressor.decompress(compressed_data).decode('utf-8')
            except ImportError:
                if safe_mode:
                    debug_print("Zstandard 模組未安裝，安全模式返回空", worker_id)
                    return {}
                raise RuntimeError("Zstandard 模組未安裝")
        else:  # gzip 或其他
            json_data = gzip.decompress(compressed_data).decode('utf-8')
        
        # JSON 解析
        result_data = json.loads(json_data)
        debug_print(f"decompress_json completed size={len(json_data)} chars", worker_id)
        
        return result_data
        
    except Exception as je:
        debug_print(f"decompress_json failed: {je}", worker_id)
        if safe_mode:
            debug_print("安全模式返回空結果", worker_id)
            return {}
        else:
            raise