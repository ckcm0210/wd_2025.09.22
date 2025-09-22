#!/usr/bin/env python3
"""
取得當前用戶的各種方法比較
"""
import os
import getpass
import subprocess
import platform

def get_user_by_various_methods():
    """比較各種取得用戶的方法"""
    print("=== 用戶識別方法比較 ===")
    
    methods = {
        "1. getpass.getuser()": lambda: getpass.getuser(),
        "2. os.environ['USERNAME']": lambda: os.environ.get('USERNAME', '未知'),
        "3. os.environ['USER']": lambda: os.environ.get('USER', '未知'),
        "4. platform.node()": lambda: platform.node(),
        "5. Excel last_author": "從 Excel 檔案 metadata 取得",
        "6. Windows whoami": lambda: subprocess.check_output(['whoami'], text=True).strip() if platform.system() == 'Windows' else '不支援',
        "7. 網絡用戶名": lambda: os.environ.get('USERDOMAIN', '') + '\\' + os.environ.get('USERNAME', ''),
    }
    
    for method_name, method_func in methods.items():
        try:
            if callable(method_func):
                result = method_func()
                print(f"{method_name}: {result}")
            else:
                print(f"{method_name}: {method_func}")
        except Exception as e:
            print(f"{method_name}: 錯誤 - {e}")
    
    print("\n=== 建議的組合方法 ===")
    print("1. 優先使用 Excel last_author (檔案實際編輯者)")
    print("2. 回退到系統當前用戶 (檔案開啟者)")
    print("3. 結合網絡用戶名 (企業環境)")

if __name__ == "__main__":
    get_user_by_various_methods()