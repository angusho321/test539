import pandas as pd
import numpy as np
from itertools import combinations
from collections import defaultdict
import datetime
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ==========================================
# 設定區
# ==========================================
FILE_539 = 'lottery_hist.xlsx'
FILE_FANTASY = 'fantasy5_hist.xlsx'

OUTPUT_539 = 'best_strategies_539.csv'
OUTPUT_FANTASY = 'best_strategies_fantasy5.csv'

# 視窗定義 (對應 Python weekday: 0=週一 ... 6=週日)
WINDOWS_MAPPING = {
    "週一~週三": [0, 1, 2],
    "週二~週四": [1, 2, 3],
    "週三~週五": [2, 3, 4],
    "週四~週六": [3, 4, 5]
}

# ==========================================
# 核心演算法
# ==========================================

def load_data(file_path, is_fantasy=False):
    """讀取資料並處理時區"""
    if not os.path.exists(file_path):
        return None
    
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except:
        try:
            df = pd.read_csv(file_path.replace('.xlsx', '.csv')) # 備援
        except:
            return None

    df['日期'] = pd.to_datetime(df['日期'])
    
    # 時區處理
    if is_fantasy:
        # 天天樂: 美國時間 + 1天 = 台灣下注時間
        df['Analysis_Date'] = df['日期'] + pd.Timedelta(days=1)
    else:
        # 539: 不需要轉換
        df['Analysis_Date'] = df['日期']

    return df

def get_data_by_week(df):
    """將資料轉換為 {(年, 週): {星期幾: {號碼集合}}}"""
    data = defaultdict(lambda: defaultdict(set))
    for _, row in df.iterrows():
        dt = row['Analysis_Date']
        year, week, _ = dt.isocalendar()
        weekday = dt.weekday()
        
        # 嘗試抓取號碼欄位 (相容不同命名)
        try:
            # 優先嘗試標準欄位
            if '號碼1' in df.columns:
                nums = {row[c] for c in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']}
            else:
                # 假設結構固定
                cols = df.columns
                nums = {row[cols[2]], row[cols[3]], row[cols[4]], row[cols[5]], row[cols[6]]}
        except:
            continue
            
        data[(year, week)][weekday] = nums
    return data

def calculate_stats(data_by_week, weeks_list, mode="win_rate"):
    """
    通用計算核心
    mode="win_rate": 計算勝率 (短期/長期)
    mode="streak": 計算連莊 (從最新週往回推)
    """
    results = []
    all_combos = list(combinations(range(1, 40), 3)) # 1-39號取3個
    
    # 預處理每週的號碼聯集 (針對不同視窗)
    # week_unions[window_name][week_key] = set(all numbers in that window)
    week_unions = defaultdict(dict)
    
    # 如果是連莊模式，必須確保週次是倒序 (最新 -> 最舊)
    target_weeks = sorted(weeks_list, reverse=True) if mode == "streak" else weeks_list
    
    for window_name, days in WINDOWS_MAPPING.items():
        for w in target_weeks:
            union_set = set()
            has_data = False
            for d in days:
                if d in data_by_week[w]:
                    union_set.update(data_by_week[w][d])
                    has_data = True
            if has_data:
                week_unions[window_name][w] = union_set

    # 開始遍歷所有組合 (9139組)
    for combo in all_combos:
        combo_set = set(combo)
        
        for window_name in WINDOWS_MAPPING.keys():
            valid_weeks = [w for w in target_weeks if w in week_unions[window_name]]
            if len(valid_weeks) < 4 and mode == "win_rate": continue
            
            if mode == "win_rate":
                wins = 0
                for w in valid_weeks:
                    # 判斷中獎: 組合 與 當週開獎號碼 有交集
                    if not combo_set.isdisjoint(week_unions[window_name][w]):
                        wins += 1
                
                rate = wins / len(valid_weeks)
                # 門檻過濾
                if rate >= 0.8: # 寬鬆門檻，後續篩選 Top 2
                    results.append({
                        "Window": window_name,
                        "Combo": combo,
                        "Score": rate, # 排序用
                        "Display": f"{rate:.1%} ({wins}/{len(valid_weeks)})"
                    })
            
            elif mode == "streak":
                streak = 0
                for w in valid_weeks:
                    if not combo_set.isdisjoint(week_unions[window_name][w]):
                        streak += 1
                    else:
                        break # 中斷
                
                if streak >= 4: # 至少連4週才紀錄
                    results.append({
                        "Window": window_name,
                        "Combo": combo,
                        "Score": streak,
                        "Display": f"{streak}週"
                    })

    return pd.DataFrame(results)

def select_best_strategies(df, threshold=0.0):
    """
    挑選邏輯:
    1. 過濾分數 < threshold
    2. 選第一名 (Score 最高)
    3. 選第二名 (Score 次高，且 Window 與第一名不同)
    """
    if df.empty:
        return "無數據", "無數據"
        
    df = df[df['Score'] >= threshold].sort_values('Score', ascending=False)
    if df.empty:
        return "無數據", "無數據"
        
    # 第一名
    top1 = df.iloc[0]
    
    # 第二名 (互斥視窗)
    top2 = None
    for _, row in df.iterrows():
        if row['Window'] != top1['Window']:
            top2 = row
            break
            
    def format_row(row):
        nums = ",".join(f"{x:02d}" for x in row['Combo'])
        return f"【{row['Window']}】{nums} [{row['Display']}]"

    res1 = format_row(top1)
    res2 = format_row(top2) if top2 is not None else "無互斥時段數據"
    
    return res1, res2

# ==========================================
# Google Drive 上傳
# ==========================================
def upload_to_drive(local_file, folder_id, creds_json):
    if not os.path.exists(local_file) or not folder_id or not creds_json:
        return

    try:
        creds_dict = json.loads(creds_json)
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=['https://www.googleapis.com/auth/drive']
        )
        service = build('drive', 'v3', credentials=creds)
        file_name = os.path.basename(local_file)

        # 搜尋是否存在
        query = f"name = '{file_name}' and '{folder_id}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])

        media = MediaFileUpload(local_file, mimetype='text/csv')

        if not files:
            # 新增
            service.files().create(
                body={'name': file_name, 'parents': [folder_id]},
                media_body=media
            ).execute()
            print(f"✅ [Drive] 新增: {file_name}")
        else:
            # 更新
            service.files().update(
                fileId=files[0]['id'],
                media_body=media
            ).execute()
            print(f"✅ [Drive] 更新: {file_name}")

    except Exception as e:
        print(f"❌ Drive Error: {e}")

# ==========================================
# 主流程
# ==========================================
def process_all():
    folder_id = os.environ.get('GOOGLE_DRIVE_FOLDER_ID')
    creds = os.environ.get('GOOGLE_CREDENTIALS')

    tasks = [
        ("539", FILE_539, OUTPUT_539, False),
        ("天天樂", FILE_FANTASY, OUTPUT_FANTASY, True)
    ]

    for name, input_file, output_file, is_fantasy in tasks:
        print(f"\n⚡ 分析 {name} (轉換時區: {is_fantasy})...")
        df = load_data(input_file, is_fantasy)
        if df is None:
            print(f"❌ 找不到 {input_file}，跳過")
            continue
            
        # 日期篩選
        max_date = df['Analysis_Date'].max()
        cutoff_8wk = max_date - pd.Timedelta(weeks=8)
        cutoff_1yr = max_date - pd.Timedelta(weeks=52)
        
        data_by_week = get_data_by_week(df)
        all_weeks = sorted(data_by_week.keys())
        
        weeks_8wk = [w for w in all_weeks if w in get_data_by_week(df[df['Analysis_Date'] >= cutoff_8wk])][1:] # 略過資料不全的當週
        weeks_1yr = [w for w in all_weeks if w in get_data_by_week(df[df['Analysis_Date'] >= cutoff_1yr])]
        
        # 1. 短期 (8週)
        print("   -> 計算短期勝率...")
        df_short = calculate_stats(data_by_week, weeks_8wk, mode="win_rate")
        
        # 2. 長期 (1年)
        print("   -> 計算長期勝率...")
        df_long = calculate_stats(data_by_week, weeks_1yr, mode="win_rate")
        
        # 3. 連莊
        print("   -> 計算連莊霸主...")
        df_streak = calculate_stats(data_by_week, all_weeks, mode="streak")
        
        # 彙整
        report = []
        
        # 短期 (門檻 85%)
        s1, s2 = select_best_strategies(df_short, threshold=0.85)
        report.append({"策略維度": "短期爆發 (近8週)", "第一組": s1, "第二組": s2})
        
        # 長期 (門檻 90%，低於顯示無數據)
        l1, l2 = select_best_strategies(df_long, threshold=0.90)
        report.append({"策略維度": "長期穩健 (近1年)", "第一組": l1, "第二組": l2})
        
        # 連莊 (至少連5週)
        st1, st2 = select_best_strategies(df_streak, threshold=5)
        report.append({"策略維度": "連莊霸主 (連勝中)", "第一組": st1, "第二組": st2})
        
        # 輸出 CSV
        res_df = pd.DataFrame(report)
        res_df.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"📄 已建立: {output_file}")
        
        # 上傳
        if folder_id:
            upload_to_drive(output_file, folder_id, creds)

if __name__ == "__main__":
    process_all()