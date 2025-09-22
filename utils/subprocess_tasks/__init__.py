"""
子進程任務模組
集中暴露目前穩定可用的子進程任務，避免在套件初始化時載入不存在或非必要的成員
"""

from .xml_tasks import extract_external_refs_task, read_meta_task
from .compression_tasks import decompress_json_task
from .task_registry import get_task_handler

__all__ = [
    'extract_external_refs_task',
    'read_meta_task',
    'decompress_json_task',
    'get_task_handler',
]
