"""
Console 壓縮檔案查看器
為 HTML timeline 提供壓縮檔案的在線查看功能
"""
import os
import json
from typing import Optional
from utils.console_compressor import get_console_compressor

def create_viewer_endpoint(compressed_file_path: str) -> str:
    """
    為壓縮檔案創建查看器端點
    返回可在 HTML timeline 中使用的 URL
    """
    try:
        # 生成 HTML 查看器檔案
        viewer_path = compressed_file_path.replace('.json.gz', '_viewer.html') \
                                         .replace('.json.lz4', '_viewer.html') \
                                         .replace('.json.zst', '_viewer.html') \
                                         .replace('.json', '_viewer.html')
        
        compressor = get_console_compressor()
        html_content = compressor.generate_html_viewer(compressed_file_path)
        
        with open(viewer_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # 返回相對路徑（用於 timeline HTML）
        return os.path.relpath(viewer_path)
        
    except Exception as e:
        print(f"[查看器生成失敗] {e}")
        return None

def serve_compressed_console(file_path: str) -> str:
    """
    直接服務壓縮的 console 內容（用於 web server）
    """
    try:
        compressor = get_console_compressor()
        return compressor.generate_html_viewer(file_path)
    except Exception as e:
        return f"<p>讀取失敗: {e}</p>"