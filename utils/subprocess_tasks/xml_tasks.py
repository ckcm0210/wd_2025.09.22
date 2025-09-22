"""
XML相關的子進程任務
"""
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

def debug_print(message: str, worker_id: int = 0):
    """輸出 debug 訊息到 stderr"""
    print(f"[xml-worker-{worker_id}] {message}", file=sys.stderr, flush=True)

def extract_external_refs_task(file_path: str, safe_mode: bool = False, worker_id: int = 0):
    """提取 Excel 外部參照"""
    debug_print(f"extract_external_refs start file={os.path.basename(file_path)}", worker_id)
    
    ref_map = {}
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            rels_xml = z.read('xl/_rels/workbook.xml.rels')
            rels = ET.fromstring(rels_xml)
            
            # 解析邏輯...（移動自原 xml_subproc_worker.py）
            
    except Exception as e:
        debug_print(f"extract_external_refs error: {e}", worker_id)
        if safe_mode:
            return {}
        else:
            raise
    
    return ref_map

def read_meta_task(file_path: str, safe_mode: bool = False, worker_id: int = 0):
    """讀取 Excel metadata"""
    debug_print(f"read_meta start file={os.path.basename(file_path)}", worker_id)
    
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            core_xml = z.read('docProps/core.xml')
            root = ET.fromstring(core_xml)
            
            # 解析邏輯...（移動自原 xml_subproc_worker.py）
            
    except Exception as e:
        debug_print(f"read_meta failed: {e}", worker_id)
        if safe_mode:
            return {"last_author": None}
        else:
            raise
    
    return {"last_author": "extracted_author"}