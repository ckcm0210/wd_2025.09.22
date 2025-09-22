"""
子進程相關設定
統一管理所有子進程的配置參數
"""

# ============ 子進程管理器設定 ============

# 是否啟用子進程
USE_SUBPROCESS = True

# 最大子進程工作者數量
SUBPROCESS_MAX_WORKERS = 2

# 子進程超時時間 (秒)
SUBPROCESS_TIMEOUT_SEC = 30

# 是否啟用安全模式重試
SUBPROCESS_SAFE_RETRY = True

# 自動使用安全模式的檔案大小閾值 (MB)
SUBPROCESS_SAFE_MODE_THRESHOLD_MB = 50

# ============ Excel 相關子進程設定 ============

# Excel 批次處理大小
EXCEL_BATCH_SIZE = 10000

# 大檔案檢測閾值 (儲存格數量)
LARGE_FILE_CELL_THRESHOLD = 1000000

# 是否強制所有 Excel 操作都通過子進程
FORCE_EXCEL_SUBPROCESS = True

# ============ XML 相關子進程設定 ============

# 是否使用 XML 子進程
USE_XML_SUBPROCESS = True

# XML 子進程最大工作者數量
XML_SUBPROCESS_MAX_WORKERS = 1

# XML 子進程超時時間 (秒)
XML_SUBPROCESS_TIMEOUT_SEC = 15

# XML 子進程安全模式重試
XML_SUBPROCESS_SAFE_RETRY = True

# ============ 安全模式設定 ============

# 禁用主進程 XML 回退
DISALLOW_MAINPROC_XML_FALLBACK = True

# 禁用主進程 metadata XML 解析
DISALLOW_MAINPROC_META_ET = True

# 僅允許子進程進行最後作者查詢
LAST_AUTHOR_SUBPROCESS_ONLY = True

# 公式專用模式的嚴格安全設定
FORMULA_ONLY_STRICT_SAFE = True

# 啟用 XML GC 保護
ENABLE_XML_GC_GUARD = True

# ============ 基準線相關子進程設定 ============

# 是否強制基準線操作通過子進程
FORCE_BASELINE_SUBPROCESS = True

# 基準線子進程超時時間 (秒)
BASELINE_SUBPROCESS_TIMEOUT_SEC = 45

# ============ 效能調優設定 ============

# 空閒 GC 調度器設定
IDLE_GC_ENABLED = True
IDLE_GC_CALM_SEC = 8
IDLE_GC_TICK_SEC = 2
IDLE_GC_COLLECT_GENERATION = -1  # -1 為完整回收
IDLE_GC_MAX_COLLECT_PER_MIN = 6
IDLE_GC_SKIP_WHEN_TK = True

# ============ 除錯和診斷設定 ============

# 顯示子進程除錯訊息
SHOW_SUBPROCESS_DEBUG = True

# 子進程崩潰報告目錄
SUBPROCESS_CRASH_DIR = "subprocess_crashes"

# 保存子進程統計資訊
ENABLE_SUBPROCESS_STATS = True

# 定期輸出統計資訊的間隔 (任務數)
SUBPROCESS_STATS_INTERVAL = 20

# ============ 容錯設定 ============

# 子進程失敗時的回退策略
SUBPROCESS_FALLBACK_STRATEGY = "safe_empty"  # "safe_empty", "retry", "abort"

# 最大重試次數
SUBPROCESS_MAX_RETRIES = 2

# 重試延遲 (秒)
SUBPROCESS_RETRY_DELAY = 1.0

# ============ 記憶體管理設定 ============

# 子進程記憶體監控
ENABLE_SUBPROCESS_MEMORY_MONITOR = True

# 記憶體使用警告閾值 (MB)
SUBPROCESS_MEMORY_WARNING_THRESHOLD_MB = 500

# 記憶體使用終止閾值 (MB)
SUBPROCESS_MEMORY_KILL_THRESHOLD_MB = 1000

# ============ 載入設定函數 ============

def load_subprocess_settings():
    """載入子進程設定"""
    settings = {}
    
    # 獲取當前模組的所有大寫變數
    import sys
    current_module = sys.modules[__name__]
    
    for name in dir(current_module):
        if name.isupper() and not name.startswith('_'):
            settings[name] = getattr(current_module, name)
    
    return settings

def apply_subprocess_settings():
    """將子進程設定應用到主設定模組"""
    try:
        import config.settings as main_settings
        subprocess_settings = load_subprocess_settings()
        
        for key, value in subprocess_settings.items():
            setattr(main_settings, key, value)
        
        print(f"[subprocess_settings] 已載入 {len(subprocess_settings)} 個子進程設定")
        return True
        
    except Exception as e:
        print(f"[subprocess_settings] 設定載入失敗: {e}")
        return False

def get_safe_mode_config():
    """取得安全模式的建議配置"""
    return {
        # 強制所有危險操作通過子進程
        'USE_SUBPROCESS': True,
        'FORCE_EXCEL_SUBPROCESS': True,
        'FORCE_BASELINE_SUBPROCESS': True,
        'USE_XML_SUBPROCESS': True,
        
        # 禁用主進程的危險操作
        'DISALLOW_MAINPROC_XML_FALLBACK': True,
        'DISALLOW_MAINPROC_META_ET': True,
        'LAST_AUTHOR_SUBPROCESS_ONLY': True,
        'FORMULA_ONLY_STRICT_SAFE': True,
        
        # 啟用保護機制
        'ENABLE_XML_GC_GUARD': True,
        'IDLE_GC_ENABLED': True,
        'IDLE_GC_SKIP_WHEN_TK': True,
        
        # 保守的超時設定
        'SUBPROCESS_TIMEOUT_SEC': 45,
        'XML_SUBPROCESS_TIMEOUT_SEC': 20,
        'BASELINE_SUBPROCESS_TIMEOUT_SEC': 60,
        
        # 啟用重試機制
        'SUBPROCESS_SAFE_RETRY': True,
        'XML_SUBPROCESS_SAFE_RETRY': True,
        'SUBPROCESS_MAX_RETRIES': 2,
        
        # 記憶體監控
        'ENABLE_SUBPROCESS_MEMORY_MONITOR': True,
        'SUBPROCESS_MEMORY_WARNING_THRESHOLD_MB': 300,
        'SUBPROCESS_MEMORY_KILL_THRESHOLD_MB': 500,
    }

def apply_safe_mode_config():
    """應用安全模式配置"""
    try:
        import config.settings as main_settings
        safe_config = get_safe_mode_config()
        
        for key, value in safe_config.items():
            setattr(main_settings, key, value)
        
        print(f"[subprocess_settings] 已應用安全模式配置")
        return True
        
    except Exception as e:
        print(f"[subprocess_settings] 安全模式配置失敗: {e}")
        return False

def validate_subprocess_environment():
    """驗證子進程執行環境"""
    validation_results = {
        'valid': True,
        'errors': [],
        'warnings': [],
        'info': []
    }
    
    try:
        # 檢查 Python 版本
        import sys
        python_version = sys.version_info
        if python_version < (3, 7):
            validation_results['errors'].append(f"Python 版本過舊: {python_version}")
            validation_results['valid'] = False
        else:
            validation_results['info'].append(f"Python 版本: {python_version}")
        
        # 檢查必要模組
        required_modules = [
            'subprocess',
            'threading', 
            'multiprocessing',
            'json',
            'gzip'
        ]
        
        for module_name in required_modules:
            try:
                __import__(module_name)
                validation_results['info'].append(f"模組 {module_name}: ✅ 可用")
            except ImportError:
                validation_results['errors'].append(f"缺少必要模組: {module_name}")
                validation_results['valid'] = False
        
        # 檢查可選但建議的模組
        optional_modules = [
            ('lz4', 'LZ4 壓縮支援'),
            ('zstandard', 'Zstandard 壓縮支援'),
            ('openpyxl', 'Excel 檔案支援'),
            ('psutil', '記憶體監控支援')
        ]
        
        for module_name, description in optional_modules:
            try:
                __import__(module_name)
                validation_results['info'].append(f"{description}: ✅ 可用")
            except ImportError:
                validation_results['warnings'].append(f"{description}: ❌ 不可用")
        
        # 檢查檔案系統權限
        import tempfile
        import os
        
        try:
            with tempfile.NamedTemporaryFile(delete=True) as tmp:
                tmp.write(b'test')
                tmp.flush()
                validation_results['info'].append("檔案系統權限: ✅ 正常")
        except Exception as e:
            validation_results['errors'].append(f"檔案系統權限問題: {e}")
            validation_results['valid'] = False
        
        # 檢查子進程支援
        try:
            import subprocess
            result = subprocess.run([sys.executable, '-c', 'print("test")'], 
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0 and result.stdout.strip() == 'test':
                validation_results['info'].append("子進程支援: ✅ 正常")
            else:
                validation_results['errors'].append("子進程執行異常")
                validation_results['valid'] = False
        except Exception as e:
            validation_results['errors'].append(f"子進程支援檢查失敗: {e}")
            validation_results['valid'] = False
    
    except Exception as e:
        validation_results['errors'].append(f"環境驗證過程失敗: {e}")
        validation_results['valid'] = False
    
    return validation_results

def print_validation_results(results):
    """列印驗證結果"""
    
    print("🔍 子進程環境驗證結果:")
    print(f"   整體狀態: {'✅ 通過' if results['valid'] else '❌ 失敗'}")
    
    if results['info']:
        print("\n   ℹ️ 資訊:")
        for info in results['info']:
            print(f"      {info}")
    
    if results['warnings']:
        print("\n   ⚠️ 警告:")
        for warning in results['warnings']:
            print(f"      {warning}")
    
    if results['errors']:
        print("\n   ❌ 錯誤:")
        for error in results['errors']:
            print(f"      {error}")
    
    print()

if __name__ == "__main__":
    print("子進程設定模組測試")
    print("=" * 40)
    
    # 驗證環境
    results = validate_subprocess_environment()
    print_validation_results(results)
    
    # 顯示設定
    settings = load_subprocess_settings()
    print(f"可用設定項目: {len(settings)}")
    for key in sorted(settings.keys()):
        print(f"  {key} = {settings[key]}")