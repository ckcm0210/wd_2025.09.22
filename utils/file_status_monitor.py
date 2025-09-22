#!/usr/bin/env python3
"""
檔案開啟/關閉狀態監控工具
用於查詢和顯示 Excel 檔案的開啟狀態
"""
import os
import sys
from datetime import datetime, timedelta
from typing import Dict, List, Optional

def get_watcher_instance():
    """取得 watcher 實例"""
    try:
        from core.watcher import active_polling_handler
        # 假設我們能從某處取得 event_handler 實例
        # 這需要根據實際的程式架構調整
        return None  # 暫時返回 None，需要後續實現
    except Exception:
        return None

def format_duration(duration: timedelta) -> str:
    """格式化時間長度"""
    total_seconds = int(duration.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    
    if hours > 0:
        return f"{hours}小時{minutes}分{seconds}秒"
    elif minutes > 0:
        return f"{minutes}分{seconds}秒"
    else:
        return f"{seconds}秒"

def show_file_status(file_open_status: Dict) -> None:
    """顯示檔案開啟狀態"""
    if not file_open_status:
        print("📊 目前沒有追蹤到任何檔案狀態")
        return
    
    current_time = datetime.now()
    open_files = []
    closed_files = []
    
    for file_path, status in file_open_status.items():
        if status.get('is_open', False):
            open_files.append((file_path, status))
        else:
            closed_files.append((file_path, status))
    
    if open_files:
        print(f"\n📂 目前開啟的檔案 ({len(open_files)} 個):")
        print("=" * 60)
        for file_path, status in open_files:
            filename = os.path.basename(file_path)
            author = status.get('last_author', '未知')
            opened_at = status.get('opened_at', current_time)
            duration = current_time - opened_at
            temp_files = status.get('temp_files', set())
            
            print(f"📄 {filename}")
            print(f"   👤 使用者: {author}")
            print(f"   🕒 開啟時間: {opened_at.strftime('%H:%M:%S')}")
            print(f"   ⏱️ 已開啟: {format_duration(duration)}")
            print(f"   📁 臨時檔案: {len(temp_files)} 個")
            if temp_files and len(temp_files) <= 3:
                for temp_file in temp_files:
                    print(f"      - {os.path.basename(temp_file)}")
            elif len(temp_files) > 3:
                temp_list = list(temp_files)[:3]
                for temp_file in temp_list:
                    print(f"      - {os.path.basename(temp_file)}")
                print(f"      - ... 還有 {len(temp_files) - 3} 個")
            print()
    
    if closed_files:
        print(f"\n📁 最近關閉的檔案 ({len(closed_files)} 個):")
        print("=" * 60)
        # 按關閉時間排序（最新的在前）
        closed_files.sort(key=lambda x: x[1].get('opened_at', datetime.min), reverse=True)
        
        for file_path, status in closed_files[:10]:  # 只顯示最近 10 個
            filename = os.path.basename(file_path)
            author = status.get('last_author', '未知')
            opened_at = status.get('opened_at', current_time)
            
            print(f"📄 {filename}")
            print(f"   👤 使用者: {author}")
            print(f"   🕒 開啟時間: {opened_at.strftime('%H:%M:%S')}")
            print()

def show_status_summary(file_open_status: Dict) -> None:
    """顯示狀態摘要"""
    if not file_open_status:
        print("📊 目前沒有追蹤到任何檔案狀態")
        return
    
    open_count = sum(1 for status in file_open_status.values() if status.get('is_open', False))
    total_count = len(file_open_status)
    closed_count = total_count - open_count
    
    print(f"\n📊 檔案狀態摘要:")
    print(f"   📂 目前開啟: {open_count} 個檔案")
    print(f"   📁 已關閉: {closed_count} 個檔案")
    print(f"   📋 總追蹤: {total_count} 個檔案")
    
    # 統計使用者
    users = set()
    for status in file_open_status.values():
        author = status.get('last_author')
        if author and author != '未知':
            users.add(author)
    
    if users:
        print(f"   👥 活躍使用者: {len(users)} 人")
        for user in sorted(users):
            user_open_count = sum(
                1 for status in file_open_status.values() 
                if status.get('is_open', False) and status.get('last_author') == user
            )
            if user_open_count > 0:
                print(f"      - {user}: {user_open_count} 個檔案")

def monitor_file_activity():
    """監控檔案活動"""
    print("📊 Excel 檔案開啟/關閉狀態監控")
    print("=" * 50)
    print("注意: 此功能需要 watchdog 程式正在運行")
    print("=" * 50)
    
    # 這裡需要實現從 watcher 取得狀態的邏輯
    # 由於架構限制，暫時提供示例
    
    print("\n使用方法:")
    print("1. 確保 watchdog 程式正在運行")
    print("2. 開啟一些 Excel 檔案")
    print("3. 觀察控制台輸出的開啟/關閉訊息")
    print("\n檔案開啟時會顯示:")
    print("   📂 檔案開啟: filename.xlsx")
    print("   👤 使用者: John Doe")
    print("   🕒 開啟時間: 14:30:25")
    print("\n檔案關閉時會顯示:")
    print("   📁 檔案關閉: filename.xlsx")
    print("   👤 使用者: John Doe") 
    print("   🕒 關閉時間: 14:35:10")
    print("   ⏱️ 開啟時長: 0:04:45")

if __name__ == "__main__":
    monitor_file_activity()