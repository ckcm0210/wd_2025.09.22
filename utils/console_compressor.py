"""
Console 詳情檔案壓縮工具
大幅減少 txt 檔案大小，同時保持 HTML timeline 的點擊查看功能
"""
import os
import json
import gzip
import time
from datetime import datetime
from typing import Dict, List, Any, Optional
import config.settings as settings

class ConsoleCompressor:
    """Console 檔案壓縮器"""
    
    def __init__(self):
        self.compression_format = getattr(settings, 'CONSOLE_COMPRESSION_FORMAT', 'gzip')  # gzip, lz4, zstd
        self.compress_threshold_kb = int(getattr(settings, 'CONSOLE_COMPRESS_THRESHOLD_KB', 50))
        self.enable_compression = bool(getattr(settings, 'ENABLE_CONSOLE_COMPRESSION', True))
        
    def compress_console_data(self, file_path: str, event_number: int, worksheet_changes: List[Dict]) -> str:
        """
        壓縮 console 資料為結構化格式
        
        Args:
            file_path: Excel 檔案路徑
            event_number: 事件編號
            worksheet_changes: 工作表變更列表
            
        Returns:
            壓縮檔案路徑
        """
        try:
            # 構建結構化資料
            structured_data = {
                'meta': {
                    'file_path': file_path,
                    'filename': os.path.basename(file_path),
                    'event_number': event_number,
                    'timestamp': datetime.now().isoformat(),
                    'compression_version': '1.0'
                },
                'changes': []
            }
            
            # 轉換變更資料為緊湊格式
            for ws_data in worksheet_changes:
                ws_name, display_old, display_new, baseline_time, current_time, old_author, new_author = ws_data
                
                ws_changes = {
                    'worksheet': ws_name,
                    'baseline_time': baseline_time,
                    'current_time': current_time,
                    'old_author': old_author,
                    'new_author': new_author,
                    'cells': []
                }
                
                # 只儲存有變化的儲存格
                all_addresses = set(display_old.keys()) | set(display_new.keys())
                for addr in sorted(all_addresses):
                    old_cell = display_old.get(addr, {})
                    new_cell = display_new.get(addr, {})
                    
                    if old_cell != new_cell:
                        cell_change = {
                            'addr': addr,
                            'old': self._compress_cell_data(old_cell),
                            'new': self._compress_cell_data(new_cell)
                        }
                        ws_changes['cells'].append(cell_change)
                
                structured_data['changes'].append(ws_changes)
            
            # 寫入壓縮檔案
            output_path = self._get_compressed_file_path(file_path, event_number, new_author)
            
            compressed_file = None
            if self.enable_compression:
                if self.compression_format == 'gzip':
                    compressed_file = self._write_gzip_file(output_path, structured_data)
                elif self.compression_format == 'lz4':
                    compressed_file = self._write_lz4_file(output_path, structured_data)
                elif self.compression_format == 'zstd':
                    compressed_file = self._write_zstd_file(output_path, structured_data)
            
            if not compressed_file:
                # 後備：寫入 JSON 檔案
                json_path = output_path.replace('.gz', '.json').replace('.lz4', '.json').replace('.zst', '.json')
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(structured_data, f, ensure_ascii=False, indent=2)
                compressed_file = json_path
            
            # 🔑 關鍵：同時生成 HTML 查看器檔案
            html_viewer_path = self._create_html_viewer_file(compressed_file, structured_data)
            
            # 返回 HTML 查看器路徑，而不是壓縮檔案路徑
            return html_viewer_path
            
        except Exception as e:
            print(f"[壓縮失敗] {e}")
            return None
    
    def _compress_cell_data(self, cell_data: Dict) -> Dict:
        """壓縮儲存格資料"""
        if not cell_data:
            return {}
        
        compressed = {}
        
        # 只保留有值的欄位
        if 'formula' in cell_data and cell_data['formula']:
            compressed['f'] = cell_data['formula']
        
        # 優先使用 cached_value，其次 value
        value = cell_data.get('cached_value')
        if value is None:
            value = cell_data.get('value')
        
        if value is not None and value != '':
            compressed['v'] = value
        
        return compressed
    
    def _get_compressed_file_path(self, file_path: str, event_number: int, author: str) -> str:
        """生成壓縮檔案路徑"""
        out_dir = getattr(settings, 'PER_EVENT_CONSOLE_DIR', None) or os.path.join(settings.LOG_FOLDER, 'console_details')
        compressed_dir = os.path.join(out_dir, 'compressed')
        os.makedirs(compressed_dir, exist_ok=True)
        
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        author_safe = self._sanitize_filename(author or 'Unknown')
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        if self.compression_format == 'gzip':
            ext = '.json.gz'
        elif self.compression_format == 'lz4':
            ext = '.json.lz4'
        elif self.compression_format == 'zstd':
            ext = '.json.zst'
        else:
            ext = '.json'
        
        filename = f"{base_name}_{author_safe}_{ts}_evt{event_number}{ext}"
        return os.path.join(compressed_dir, filename)
    
    def _write_gzip_file(self, output_path: str, data: Dict) -> str:
        """寫入 GZIP 壓縮檔案"""
        json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
        
        with gzip.open(output_path, 'wt', encoding='utf-8') as f:
            f.write(json_str)
        
        # 計算壓縮比
        original_size = len(json_str.encode('utf-8'))
        compressed_size = os.path.getsize(output_path)
        ratio = (1 - compressed_size / original_size) * 100
        
        if getattr(settings, 'SHOW_COMPRESSION_STATS', False):
            print(f"[壓縮] {os.path.basename(output_path)} - 壓縮率: {ratio:.1f}% ({original_size//1024}KB -> {compressed_size//1024}KB)")
        
        return output_path
    
    def _write_lz4_file(self, output_path: str, data: Dict) -> str:
        """寫入 LZ4 壓縮檔案"""
        try:
            import lz4.frame
            json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
            json_bytes = json_str.encode('utf-8')
            
            with open(output_path, 'wb') as f:
                f.write(lz4.frame.compress(json_bytes))
            
            compressed_size = os.path.getsize(output_path)
            ratio = (1 - compressed_size / len(json_bytes)) * 100
            
            if getattr(settings, 'SHOW_COMPRESSION_STATS', False):
                print(f"[壓縮] {os.path.basename(output_path)} - LZ4 壓縮率: {ratio:.1f}%")
            
            return output_path
        except ImportError:
            # LZ4 不可用，回退到 GZIP
            return self._write_gzip_file(output_path.replace('.lz4', '.gz'), data)
    
    def _write_zstd_file(self, output_path: str, data: Dict) -> str:
        """寫入 ZSTD 壓縮檔案"""
        try:
            import zstandard as zstd
            json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
            json_bytes = json_str.encode('utf-8')
            
            cctx = zstd.ZstdCompressor(level=int(getattr(settings, 'ZSTD_COMPRESSION_LEVEL', 3)))
            
            with open(output_path, 'wb') as f:
                f.write(cctx.compress(json_bytes))
            
            compressed_size = os.path.getsize(output_path)
            ratio = (1 - compressed_size / len(json_bytes)) * 100
            
            if getattr(settings, 'SHOW_COMPRESSION_STATS', False):
                print(f"[壓縮] {os.path.basename(output_path)} - ZSTD 壓縮率: {ratio:.1f}%")
            
            return output_path
        except ImportError:
            # ZSTD 不可用，回退到 GZIP
            return self._write_gzip_file(output_path.replace('.zst', '.gz'), data)
    
    def decompress_console_data(self, compressed_path: str) -> Optional[Dict]:
        """解壓縮 console 資料"""
        try:
            if not os.path.exists(compressed_path):
                return None
            
            if compressed_path.endswith('.gz'):
                with gzip.open(compressed_path, 'rt', encoding='utf-8') as f:
                    return json.load(f)
            elif compressed_path.endswith('.lz4'):
                try:
                    import lz4.frame
                    with open(compressed_path, 'rb') as f:
                        compressed_data = f.read()
                    json_str = lz4.frame.decompress(compressed_data).decode('utf-8')
                    return json.loads(json_str)
                except ImportError:
                    return None
            elif compressed_path.endswith('.zst'):
                try:
                    import zstandard as zstd
                    dctx = zstd.ZstdDecompressor()
                    with open(compressed_path, 'rb') as f:
                        compressed_data = f.read()
                    json_str = dctx.decompress(compressed_data).decode('utf-8')
                    return json.loads(json_str)
                except ImportError:
                    return None
            elif compressed_path.endswith('.json'):
                with open(compressed_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            
        except Exception as e:
            print(f"[解壓縮失敗] {compressed_path}: {e}")
            return None
    
    def generate_html_viewer(self, compressed_path: str) -> str:
        """為壓縮檔案生成 HTML 查看器"""
        data = self.decompress_console_data(compressed_path)
        if not data:
            return "<p>無法讀取檔案</p>"
        
        html_content = []
        html_content.append("<html><head><meta charset='utf-8'><title>Excel 變更詳情</title>")
        html_content.append("<style>")
        html_content.append("body { font-family: 'Consolas', monospace; background: #1e1e1e; color: #d4d4d4; }")
        html_content.append("table { border-collapse: collapse; width: 100%; margin: 20px 0; }")
        html_content.append("th, td { border: 1px solid #444; padding: 8px; text-align: left; }")
        html_content.append("th { background: #333; }")
        html_content.append(".formula { color: #00bcd4; }")
        html_content.append(".value { color: #ffd54f; }")
        html_content.append("</style></head><body>")
        
        # 檔案資訊
        meta = data.get('meta', {})
        html_content.append(f"<h2>檔案: {meta.get('filename', 'Unknown')}</h2>")
        html_content.append(f"<p>事件編號: #{meta.get('event_number', 'N/A')}</p>")
        html_content.append(f"<p>時間: {meta.get('timestamp', 'N/A')}</p>")
        
        # 變更詳情
        for change in data.get('changes', []):
            html_content.append(f"<h3>工作表: {change.get('worksheet', 'Unknown')}</h3>")
            html_content.append("<table>")
            html_content.append("<tr><th>地址</th><th>類型</th><th>舊值</th><th>新值</th></tr>")
            
            for cell in change.get('cells', []):
                addr = cell.get('addr', '')
                old_data = cell.get('old', {})
                new_data = cell.get('new', {})
                
                # 公式行
                old_formula = old_data.get('f', '(No formula)')
                new_formula = new_data.get('f', '(No formula)')
                html_content.append(f"<tr><td>{addr}</td><td class='formula'>formula</td><td>{old_formula}</td><td>{new_formula}</td></tr>")
                
                # 值行
                old_value = repr(old_data.get('v', '(Empty)'))
                new_value = repr(new_data.get('v', '(Empty)'))
                html_content.append(f"<tr><td></td><td class='value'>value</td><td>{old_value}</td><td>{new_value}</td></tr>")
            
            html_content.append("</table>")
        
        html_content.append("</body></html>")
        return ''.join(html_content)
    
    def _create_html_viewer_file(self, compressed_path: str, data: Dict) -> str:
        """創建可直接在瀏覽器中查看的 HTML 檔案"""
        try:
            # 生成 HTML 檔案路徑
            html_path = compressed_path.replace('.json.gz', '.html') \
                                     .replace('.json.lz4', '.html') \
                                     .replace('.json.zst', '.html') \
                                     .replace('.json', '.html')
            
            # 生成完整的 HTML 內容
            html_content = self._generate_full_html_content(data)
            
            # 寫入 HTML 檔案
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            return html_path
            
        except Exception as e:
            print(f"[HTML查看器生成失敗] {e}")
            return compressed_path  # 回退到壓縮檔案路徑
    
    def _generate_full_html_content(self, data: Dict) -> str:
        """生成完整的 HTML 內容，模擬原始 txt 格式"""
        html_lines = []
        
        # HTML 頭部
        html_lines.append("""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Excel 變更詳情</title>
    <style>
        body { 
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace; 
            background: #1e1e1e; 
            color: #d4d4d4; 
            padding: 20px;
            margin: 0;
            line-height: 1.4;
        }
        .header { 
            color: #569cd6; 
            font-weight: bold; 
            margin: 20px 0 10px 0;
        }
        .separator { 
            color: #808080; 
            margin: 5px 0;
        }
        .event-info {
            background: #2d2d30;
            padding: 10px;
            border-left: 4px solid #007acc;
            margin-bottom: 20px;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin: 10px 0;
            font-family: inherit;
        }
        th, td { 
            padding: 6px 12px; 
            text-align: left; 
            border: 1px solid #444;
            font-family: inherit;
        }
        th { 
            background: #333; 
            color: #ffffff;
        }
        .formula { color: #00bcd4; }
        .value { color: #ffd54f; }
        .address { color: #c586c0; font-weight: bold; }
        pre { margin: 0; white-space: pre-wrap; }
    </style>
</head>
<body>""")
        
        # 檔案資訊
        meta = data.get('meta', {})
        html_lines.append(f'<div class="event-info">')
        html_lines.append(f'<h2>📊 {meta.get("filename", "Unknown")}</h2>')
        html_lines.append(f'<p><strong>事件編號:</strong> #{meta.get("event_number", "N/A")}</p>')
        html_lines.append(f'<p><strong>處理時間:</strong> {meta.get("timestamp", "N/A")}</p>')
        html_lines.append(f'</div>')
        
        # 為每個工作表生成表格，模擬原始格式
        for change in data.get('changes', []):
            worksheet = change.get('worksheet', 'Unknown')
            baseline_time = change.get('baseline_time', 'N/A')
            current_time = change.get('current_time', 'N/A') 
            old_author = change.get('old_author', 'N/A')
            new_author = change.get('new_author', 'N/A')
            
            # 工作表標題（模擬原始格式）
            html_lines.append('<div class="separator">' + '=' * 120 + '</div>')
            html_lines.append(f'<div class="header">(事件#{meta.get("event_number", "N/A")}) {meta.get("filename", "Unknown")} [Worksheet: {worksheet}]</div>')
            html_lines.append('<div class="separator">' + '=' * 120 + '</div>')
            
            # 表格標題
            html_lines.append('<table>')
            html_lines.append('<tr>')
            html_lines.append('<th style="width: 80px;">Address</th>')
            html_lines.append('<th style="width: 80px;">Type</th>')
            html_lines.append(f'<th>Baseline ({baseline_time} by {old_author})</th>')
            html_lines.append(f'<th>Current ({current_time} by {new_author})</th>')
            html_lines.append('</tr>')
            
            # 變更內容
            if not change.get('cells'):
                html_lines.append('<tr><td colspan="4" style="text-align: center; color: #808080;">(No cell changes)</td></tr>')
            else:
                for cell in change.get('cells', []):
                    addr = cell.get('addr', '')
                    old_data = cell.get('old', {})
                    new_data = cell.get('new', {})
                    
                    # 公式行
                    old_formula = old_data.get('f') or '(No formula)'
                    new_formula = new_data.get('f') or '(No formula)'
                    
                    html_lines.append('<tr>')
                    html_lines.append(f'<td class="address">{addr}</td>')
                    html_lines.append('<td class="formula">formula</td>')
                    html_lines.append(f'<td><pre>{old_formula}</pre></td>')
                    html_lines.append(f'<td><pre>{new_formula}</pre></td>')
                    html_lines.append('</tr>')
                    
                    # 值行
                    old_value = old_data.get('v')
                    new_value = new_data.get('v')
                    
                    old_display = '(Empty)' if old_value is None or old_value == '' else repr(old_value)
                    new_display = '(Empty)' if new_value is None or new_value == '' else repr(new_value)
                    
                    html_lines.append('<tr>')
                    html_lines.append('<td></td>')  # 空的地址欄
                    html_lines.append('<td class="value">value</td>')
                    html_lines.append(f'<td><pre>{old_display}</pre></td>')
                    html_lines.append(f'<td><pre>{new_display}</pre></td>')
                    html_lines.append('</tr>')
                    
                    # 分隔線（除了最後一行）
                    html_lines.append('<tr><td colspan="4" style="border-top: 1px solid #666; padding: 0;"></td></tr>')
            
            html_lines.append('</table>')
            html_lines.append('<div class="separator">' + '=' * 120 + '</div>')
            html_lines.append('<br>')
        
        # HTML 結尾
        html_lines.append('</body></html>')
        
        return '\n'.join(html_lines)

    def _sanitize_filename(self, name: str) -> str:
        """清理檔案名稱"""
        import re
        name = re.sub(r'[<>:"/\\|?*]', '_', str(name))
        return name[:50]  # 限制長度

# 全域實例
_console_compressor = None

def get_console_compressor() -> ConsoleCompressor:
    """取得壓縮器實例"""
    global _console_compressor
    if _console_compressor is None:
        _console_compressor = ConsoleCompressor()
    return _console_compressor