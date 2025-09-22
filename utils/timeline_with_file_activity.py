#!/usr/bin/env python3
"""
整合檔案活動的 Timeline 生成器
將檔案開啟/關閉事件添加到 HTML timeline 中
"""
import os
import json
import sqlite3
from datetime import datetime, timedelta
from typing import List, Dict, Any
import config.settings as settings

def get_file_activities(hours: int = 24) -> List[Dict[str, Any]]:
    """取得檔案活動記錄"""
    try:
        log_folder = getattr(settings, 'LOG_FOLDER', '.')
        db_path = os.path.join(log_folder, 'file_activity.sqlite')
        
        if not os.path.exists(db_path):
            return []
        
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute('''
                SELECT * FROM file_activity 
                WHERE timestamp >= datetime('now', '-{} hours')
                ORDER BY timestamp DESC
            '''.format(hours))
            
            activities = []
            for row in cursor.fetchall():
                activity = dict(row)
                # 解析 temp_files JSON
                if activity.get('temp_files'):
                    try:
                        activity['temp_files'] = json.loads(activity['temp_files'])
                    except:
                        activity['temp_files'] = []
                activities.append(activity)
            
            return activities
    except Exception as e:
        print(f"[timeline] 取得檔案活動失敗: {e}")
        return []

def convert_activities_to_timeline_events(activities: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """將檔案活動轉換為 timeline 事件格式"""
    events = []
    
    for activity in activities:
        # 基本事件資訊
        event = {
            'id': f"file_activity_{activity['id']}",
            'type': 'file_activity',
            'timestamp': activity['timestamp'],
            'file_name': activity['file_name'],
            'file_path': activity['file_path'],
            'action': activity['action'],
            'user_name': activity['user_name'],
            'session_id': activity.get('session_id', ''),
        }
        
        # 根據動作類型添加特定資訊
        if activity['action'] == 'open':
            event.update({
                'title': f"📂 {activity['file_name']} 被開啟",
                'description': f"使用者 {activity['user_name']} 開啟了檔案",
                'temp_files': activity.get('temp_files', []),
                'color': '#4CAF50',  # 綠色
                'icon': '📂'
            })
        elif activity['action'] == 'close':
            duration = activity.get('duration_seconds', 0)
            duration_str = format_duration(duration)
            event.update({
                'title': f"📁 {activity['file_name']} 被關閉",
                'description': f"使用者 {activity['user_name']} 關閉了檔案，使用時長: {duration_str}",
                'duration_seconds': duration,
                'duration_formatted': duration_str,
                'color': '#F44336',  # 紅色
                'icon': '📁'
            })
        
        events.append(event)
    
    return events

def format_duration(seconds: float) -> str:
    """格式化時間長度"""
    if seconds < 60:
        return f"{int(seconds)}秒"
    elif seconds < 3600:
        minutes = int(seconds // 60)
        secs = int(seconds % 60)
        return f"{minutes}分{secs}秒"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        return f"{hours}時{minutes}分"

def generate_file_activity_timeline_html(activities: List[Dict[str, Any]], output_path: str):
    """生成包含檔案活動的 HTML timeline"""
    
    # 轉換活動為事件
    events = convert_activities_to_timeline_events(activities)
    
    # 統計資訊
    open_events = [e for e in events if e['action'] == 'open']
    close_events = [e for e in events if e['action'] == 'close']
    users = list(set(e['user_name'] for e in events if e.get('user_name')))
    
    # 用戶活動統計
    user_stats = {}
    for event in events:
        user = event.get('user_name', '未知')
        if user not in user_stats:
            user_stats[user] = {'open': 0, 'close': 0, 'total_duration': 0}
        
        user_stats[user][event['action']] += 1
        if event['action'] == 'close':
            user_stats[user]['total_duration'] += event.get('duration_seconds', 0)
    
    # HTML 模板
    html_content = f'''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel 檔案活動 Timeline</title>
    <style>
        body {{
            font-family: 'Microsoft JhengHei', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            text-align: center;
        }}
        .stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }}
        .stat-card {{
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            text-align: center;
        }}
        .stat-value {{
            font-size: 2em;
            font-weight: bold;
            color: #333;
        }}
        .stat-label {{
            color: #666;
            margin-top: 5px;
        }}
        .timeline {{
            background: white;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .event {{
            display: flex;
            margin-bottom: 15px;
            padding: 15px;
            border-left: 4px solid;
            border-radius: 5px;
            background: #f9f9f9;
        }}
        .event.open {{
            border-left-color: #4CAF50;
            background: #f1f8e9;
        }}
        .event.close {{
            border-left-color: #F44336;
            background: #ffebee;
        }}
        .event-icon {{
            font-size: 1.5em;
            margin-right: 15px;
            min-width: 30px;
        }}
        .event-content {{
            flex: 1;
        }}
        .event-title {{
            font-weight: bold;
            margin-bottom: 5px;
        }}
        .event-meta {{
            color: #666;
            font-size: 0.9em;
        }}
        .user-stats {{
            margin-top: 20px;
            background: white;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .user-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 0;
            border-bottom: 1px solid #eee;
        }}
        .user-row:last-child {{
            border-bottom: none;
        }}
        .filter-buttons {{
            margin-bottom: 20px;
            text-align: center;
        }}
        .filter-btn {{
            background: #667eea;
            color: white;
            border: none;
            padding: 8px 16px;
            margin: 0 5px;
            border-radius: 5px;
            cursor: pointer;
        }}
        .filter-btn:hover {{
            background: #5a6fd8;
        }}
        .filter-btn.active {{
            background: #4a5bc4;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Excel 檔案活動 Timeline</h1>
        <p>檔案開啟/關閉活動記錄</p>
    </div>
    
    <div class="stats">
        <div class="stat-card">
            <div class="stat-value">{len(events)}</div>
            <div class="stat-label">總活動數</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">{len(open_events)}</div>
            <div class="stat-label">檔案開啟</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">{len(close_events)}</div>
            <div class="stat-label">檔案關閉</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">{len(users)}</div>
            <div class="stat-label">活躍使用者</div>
        </div>
    </div>
    
    <div class="filter-buttons">
        <button class="filter-btn active" onclick="filterEvents('all')">全部</button>
        <button class="filter-btn" onclick="filterEvents('open')">開啟事件</button>
        <button class="filter-btn" onclick="filterEvents('close')">關閉事件</button>
    </div>
    
    <div class="timeline">
        <h2>📅 活動時間軸</h2>
'''
    
    # 添加事件
    for event in events:
        action_class = event['action']
        timestamp = datetime.fromisoformat(event['timestamp']).strftime('%Y-%m-%d %H:%M:%S')
        
        html_content += f'''
        <div class="event {action_class}" data-action="{event['action']}">
            <div class="event-icon">{event['icon']}</div>
            <div class="event-content">
                <div class="event-title">{event['title']}</div>
                <div class="event-meta">
                    🕒 {timestamp} | 👤 {event['user_name']}
                    {f" | ⏱️ {event.get('duration_formatted', '')}" if event['action'] == 'close' else ""}
                </div>
            </div>
        </div>
'''
    
    html_content += '''
    </div>
    
    <div class="user-stats">
        <h2>👥 使用者統計</h2>
'''
    
    # 添加用戶統計
    for user, stats in user_stats.items():
        avg_duration = stats['total_duration'] / max(stats['close'], 1)
        avg_duration_str = format_duration(avg_duration)
        total_duration_str = format_duration(stats['total_duration'])
        
        html_content += f'''
        <div class="user-row">
            <div>
                <strong>👤 {user}</strong>
            </div>
            <div>
                📂 {stats['open']} 次開啟 | 
                📁 {stats['close']} 次關閉 | 
                ⏱️ 總時長: {total_duration_str} | 
                📊 平均: {avg_duration_str}
            </div>
        </div>
'''
    
    html_content += '''
    </div>
    
    <script>
        function filterEvents(type) {
            const events = document.querySelectorAll('.event');
            const buttons = document.querySelectorAll('.filter-btn');
            
            // 更新按鈕狀態
            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');
            
            // 過濾事件
            events.forEach(event => {
                if (type === 'all' || event.dataset.action === type) {
                    event.style.display = 'flex';
                } else {
                    event.style.display = 'none';
                }
            });
        }
        
        // 自動刷新（每分鐘）
        setTimeout(() => {
            location.reload();
        }, 60000);
    </script>
</body>
</html>
'''
    
    # 寫入檔案
    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"[timeline] 檔案活動 timeline 已生成: {output_path}")
    except Exception as e:
        print(f"[timeline] 生成檔案活動 timeline 失敗: {e}")

def update_file_activity_timeline():
    """更新檔案活動 timeline"""
    try:
        # 取得最近 24 小時的活動
        activities = get_file_activities(hours=24)
        
        if not activities:
            print("[timeline] 沒有檔案活動記錄")
            return
        
        # 生成 HTML
        log_folder = getattr(settings, 'LOG_FOLDER', '.')
        timeline_dir = os.path.join(log_folder, 'timeline')
        output_path = os.path.join(timeline_dir, 'file_activity.html')
        
        generate_file_activity_timeline_html(activities, output_path)
        
    except Exception as e:
        print(f"[timeline] 更新檔案活動 timeline 失敗: {e}")

if __name__ == "__main__":
    update_file_activity_timeline()