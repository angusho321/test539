import pandas as pd
import numpy as np
from itertools import combinations
from collections import defaultdict
import datetime
import os
import json
import sys
import argparse
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
            # 嘗試讀取 CSV（可能是 fantasy5_hist.xlsx - Sheet1.csv）
            csv_path = file_path.replace('.xlsx', '.csv')
            if not os.path.exists(csv_path):
                # 嘗試其他可能的 CSV 檔名
                csv_path = file_path.replace('.xlsx', ' - Sheet1.csv')
            df = pd.read_csv(csv_path) # 備援
        except:
            return None

    # 處理日期欄位：支援多種日期格式（包含或不包含時間）
    try:
        # 先嘗試解析為 datetime，讓 pandas 自動推斷格式
        df['日期'] = pd.to_datetime(df['日期'], format='mixed', errors='coerce')
        # 如果自動推斷失敗，嘗試常見格式
        if df['日期'].isna().any():
            df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
        if df['日期'].isna().any():
            df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d', errors='coerce')
    except:
        # 最後嘗試不指定格式，讓 pandas 自動處理
        df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    
    # 移除無法解析的日期行
    df = df.dropna(subset=['日期'])
    
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
def upload_to_drive(local_file, file_id=None, folder_id=None, creds_json=None):
    """
    上傳文件到 Google Drive
    優先使用文件 ID 更新現有文件，如果沒有則使用資料夾 ID 創建新文件
    如果本地文件是 CSV，會轉換為 XLSX 格式上傳
    """
    if not os.path.exists(local_file):
        print(f"❌ 本地文件不存在: {local_file}")
        return False
    
    if not creds_json:
        print(f"⚠️ 未設置 GOOGLE_CREDENTIALS")
        return False

    try:
        # 解析認證資訊
        if isinstance(creds_json, str):
            creds_dict = json.loads(creds_json)
        else:
            creds_dict = creds_json
        
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=['https://www.googleapis.com/auth/drive']
        )
        service = build('drive', 'v3', credentials=creds)
        file_name = os.path.basename(local_file)
        
        # 獲取服務帳號郵件（用於調試）
        service_account_email = creds_dict.get('client_email', 'unknown')
        print(f"🔍 使用服務帳號: {service_account_email}")

        # 如果本地文件是 CSV，轉換為 XLSX（因為 Google Drive 上的文件是 XLSX）
        upload_file = local_file
        upload_mime_type = 'text/csv'
        if local_file.endswith('.csv'):
            # 轉換 CSV 為 XLSX
            xlsx_file = local_file.replace('.csv', '.xlsx')
            try:
                df = pd.read_csv(local_file, encoding='utf-8-sig')
                df.to_excel(xlsx_file, index=False, engine='openpyxl')
                upload_file = xlsx_file
                upload_mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                print(f"📊 已將 CSV 轉換為 XLSX: {xlsx_file}")
            except Exception as e:
                print(f"⚠️ CSV 轉換為 XLSX 失敗，使用原始 CSV: {e}")
                # 繼續使用 CSV
        
        # 如果目標文件是 XLSX，更新文件名
        if file_id:
            # 檢查目標文件類型
            try:
                file_info = service.files().get(fileId=file_id, fields='name,mimeType').execute()
                target_name = file_info.get('name', '')
                if target_name.endswith('.xlsx') and upload_file.endswith('.csv'):
                    # 目標是 XLSX，但我們有 CSV，需要轉換
                    if not upload_file.endswith('.xlsx'):
                        xlsx_file = local_file.replace('.csv', '.xlsx')
                        df = pd.read_csv(local_file, encoding='utf-8-sig')
                        df.to_excel(xlsx_file, index=False, engine='openpyxl')
                        upload_file = xlsx_file
                        upload_mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        print(f"📊 已將 CSV 轉換為 XLSX 以匹配目標文件格式")
            except:
                pass  # 如果無法獲取文件資訊，繼續使用原始文件

        media = MediaFileUpload(upload_file, mimetype=upload_mime_type)

        # 優先嘗試使用文件 ID 更新現有文件
        if file_id:
            try:
                print(f"🔍 嘗試更新現有文件 ID: {file_id}")
                # 驗證文件是否存在且有權限
                file_info = service.files().get(
                    fileId=file_id,
                    fields='id,name,parents,mimeType'
                ).execute()
                print(f"✅ 文件驗證成功: {file_info.get('name', '未知')}")
                print(f"   📁 文件 ID: {file_info.get('id')}")
                print(f"   📂 父資料夾: {file_info.get('parents', ['根目錄'])}")
                
                # 更新文件
                updated_file = service.files().update(
                    fileId=file_id,
                    media_body=media,
                    fields='id,name,webViewLink'
                ).execute()
                print(f"✅ [Drive] 更新文件: {updated_file.get('name')} (ID: {file_id})")
                print(f"   🔗 檢視連結: {updated_file.get('webViewLink', 'N/A')}")
                
                # 清理臨時創建的 XLSX 文件（如果原始是 CSV）
                if upload_file != local_file and upload_file.endswith('.xlsx'):
                    try:
                        os.remove(upload_file)
                        print(f"🧹 已清理臨時文件: {upload_file}")
                    except:
                        pass
                
                return True
            except Exception as update_error:
                error_msg = str(update_error)
                if '404' in error_msg or 'notFound' in error_msg:
                    print(f"⚠️ 文件 ID 不存在或無權限，嘗試創建新文件...")
                else:
                    print(f"⚠️ 更新文件失敗: {update_error}")
                    print(f"   嘗試其他解決方案...")

        # 如果沒有文件 ID 或更新失敗，嘗試搜索現有文件或創建新文件
        # 優先在根目錄搜索（與其他文件同路徑），如果提供了資料夾 ID 則在資料夾中搜索
        try:
            # 搜索文件名（如果本地是 CSV，搜索對應的 XLSX 文件名）
            search_name = file_name.replace('.csv', '.xlsx') if file_name.endswith('.csv') else file_name
            print(f"🔍 嘗試搜索現有文件: {search_name}")
            
            # 構建搜索查詢
            if folder_id:
                # 在指定資料夾中搜索
                print(f"   📁 在資料夾中搜索 (ID: {folder_id})")
                query = f"name = '{search_name}' and '{folder_id}' in parents and trashed = false"
            else:
                # 在根目錄搜索（不指定 parents，與其他文件同路徑）
                print(f"   📁 在根目錄搜索（與其他文件同路徑）")
                query = f"name = '{search_name}' and trashed = false"
                # 排除在資料夾中的文件（只搜索根目錄）
                # 注意：Google Drive API 無法直接搜索根目錄，我們需要先搜索所有同名文件，然後過濾
            
            # 搜索文件
            # 如果沒有指定資料夾，搜索所有同名文件（包括根目錄和資料夾中的）
            if not folder_id:
                # 搜索所有同名文件
                query = f"name = '{search_name}' and trashed = false"
            
            results = service.files().list(q=query, fields="files(id, name, parents)").execute()
            existing_files = results.get('files', [])
            
            # 如果沒有指定資料夾，優先選擇根目錄的文件（與其他文件同路徑）
            if not folder_id and existing_files:
                # 嘗試找到與其他文件（如 fantasy5_hist.xlsx）同路徑的文件
                # 先獲取一個參考文件的 parents（如果可能）
                try:
                    # 嘗試獲取 fantasy5_hist 或 prediction_log 的 parents 作為參考
                    ref_file_id = os.environ.get('FANTASY5_HIST_FILE_ID') or os.environ.get('FANTASY5_PREDICTION_LOG_FILE_ID')
                    if ref_file_id:
                        ref_info = service.files().get(fileId=ref_file_id, fields='parents').execute()
                        ref_parents = ref_info.get('parents', [])
                        # 優先選擇與參考文件相同 parents 的文件
                        matching_files = [f for f in existing_files if f.get('parents', []) == ref_parents]
                        if matching_files:
                            existing_files = matching_files
                except:
                    pass  # 如果無法獲取參考，使用所有找到的文件
            
            if existing_files:
                # 找到現有文件，更新它
                existing_file_id = existing_files[0]['id']
                existing_file_name = existing_files[0]['name']
                print(f"📄 找到現有文件: {existing_file_name} (ID: {existing_file_id})")
                updated_file = service.files().update(
                    fileId=existing_file_id,
                    media_body=media,
                    fields='id,name,webViewLink'
                ).execute()
                print(f"✅ [Drive] 更新現有文件: {updated_file.get('name')} (ID: {existing_file_id})")
                print(f"   🔗 檢視連結: {updated_file.get('webViewLink', 'N/A')}")
                print(f"   💡 建議將此文件 ID ({existing_file_id}) 新增為 GitHub Secret")
                
                # 清理臨時創建的 XLSX 文件（如果原始是 CSV）
                if upload_file != local_file and upload_file.endswith('.xlsx'):
                    try:
                        os.remove(upload_file)
                        print(f"🧹 已清理臨時文件: {upload_file}")
                    except:
                        pass
                
                return True
            else:
                # 沒有找到現有文件，創建新文件
                print(f"📝 未找到現有文件，創建新文件...")
                create_name = search_name if upload_file.endswith('.xlsx') else file_name
                file_metadata = {
                    'name': create_name
                }
                
                # 如果指定了資料夾，設定父資料夾；否則創建在根目錄
                if folder_id:
                    file_metadata['parents'] = [folder_id]
                    print(f"   📁 目標資料夾 ID: {folder_id}")
                else:
                    print(f"   📁 創建在根目錄（與其他文件同路徑）")
                
                created_file = service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id,name,webViewLink'
                ).execute()
                print(f"✅ [Drive] 新增文件: {created_file.get('name')}")
                print(f"   📁 文件 ID: {created_file.get('id')}")
                print(f"   🔗 檢視連結: {created_file.get('webViewLink', 'N/A')}")
                print(f"   💡 建議將此文件 ID ({created_file.get('id')}) 新增為 GitHub Secret")
                
                # 清理臨時創建的 XLSX 文件（如果原始是 CSV）
                if upload_file != local_file and upload_file.endswith('.xlsx'):
                    try:
                        os.remove(upload_file)
                        print(f"🧹 已清理臨時文件: {upload_file}")
                    except:
                        pass
                
                return True
                
        except Exception as create_error:
            error_msg = str(create_error)
            print(f"❌ [Drive] 搜索或創建文件失敗: {create_error}")
            if '404' in error_msg or 'notFound' in error_msg:
                if folder_id:
                    print(f"   💡 如果文件在根目錄，請不要設置 GOOGLE_DRIVE_FOLDER_ID")
                    print(f"   💡 或者直接設置 BEST_STRATEGIES_FANTASY5_FILE_ID 或 BEST_STRATEGIES_539_FILE_ID")
            return False
        
        return True

    except json.JSONDecodeError as e:
        print(f"❌ [Drive] 認證資訊格式錯誤: {e}")
        return False
    except Exception as e:
        error_msg = str(e)
        print(f"❌ [Drive] 上傳失敗: {error_msg}")
        
        # 針對常見錯誤提供解決方案
        if '404' in error_msg or 'notFound' in error_msg:
            print(f"   💡 這通常是因為:")
            print(f"      1. 資料夾 ID 不正確")
            print(f"      2. 服務帳號沒有權限訪問該資料夾")
            print(f"      3. 資料夾已被刪除")
            print(f"   💡 解決方案:")
            print(f"      1. 確認 GOOGLE_DRIVE_FOLDER_ID 是否正確")
            print(f"      2. 在 Google Drive 中分享資料夾給服務帳號")
            print(f"      3. 確保服務帳號有「編輯者」權限")
        elif '403' in error_msg or 'permission' in error_msg.lower():
            print(f"   💡 權限不足，請確認服務帳號有「編輯者」權限")
        elif '401' in error_msg or 'unauthorized' in error_msg.lower():
            print(f"   💡 認證失敗，請確認 GOOGLE_CREDENTIALS 是否正確")
        
        import traceback
        traceback.print_exc()
        return False

# ==========================================
# 主流程
# ==========================================
def process_single(name, input_file, output_file, is_fantasy, file_id=None, folder_id=None, creds=None):
    """處理單一彩球的分析"""
    print(f"\n⚡ 分析 {name} (轉換時區: {is_fantasy})...")
    df = load_data(input_file, is_fantasy)
    if df is None:
        print(f"❌ 找不到 {input_file}，跳過")
        return False
        
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
    
    # 輸出 CSV (確保一定會創建文件)
    res_df = pd.DataFrame(report)
    res_df.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"📄 已建立: {output_file}")
    
    # 驗證文件確實存在
    if not os.path.exists(output_file):
        print(f"❌ 警告: {output_file} 創建失敗")
        return False
    
    # 上傳到 Google Drive
    if creds:
        try:
            upload_to_drive(output_file, file_id=file_id, folder_id=folder_id, creds_json=creds)
            print(f"✅ {output_file} 已上傳到 Google Drive")
        except Exception as e:
            print(f"⚠️ 上傳 {output_file} 到 Google Drive 時發生錯誤: {e}")
            print(f"   本地文件已創建: {output_file}")
    else:
        print(f"⚠️ 未設置 GOOGLE_CREDENTIALS")
        print(f"   📄 本地文件已創建: {output_file}")
        print(f"   💡 提示: 在 GitHub Actions 中，環境變數會自動從 Secrets 讀取")
        print(f"   💡 本地測試時，可以手動設置環境變數或跳過上傳步驟")
    
    return True

def process_all():
    """處理所有彩球的分析（預設行為）"""
    folder_id = os.environ.get('GOOGLE_DRIVE_FOLDER_ID')
    creds = os.environ.get('GOOGLE_CREDENTIALS')
    
    file_id_539 = os.environ.get('BEST_STRATEGIES_539_FILE_ID')
    file_id_fantasy = os.environ.get('BEST_STRATEGIES_FANTASY5_FILE_ID')

    tasks = [
        ("539", FILE_539, OUTPUT_539, False, file_id_539),
        ("天天樂", FILE_FANTASY, OUTPUT_FANTASY, True, file_id_fantasy)
    ]

    for name, input_file, output_file, is_fantasy, file_id in tasks:
        process_single(name, input_file, output_file, is_fantasy, file_id=file_id, folder_id=folder_id, creds=creds)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='分析彩球策略')
    parser.add_argument('--type', type=str, choices=['539', 'fantasy5', 'all'], 
                       default='all', help='指定要分析的彩球類型: 539, fantasy5, 或 all (預設)')
    
    args = parser.parse_args()
    
    folder_id = os.environ.get('GOOGLE_DRIVE_FOLDER_ID')
    creds = os.environ.get('GOOGLE_CREDENTIALS')
    file_id_539 = os.environ.get('BEST_STRATEGIES_539_FILE_ID')
    file_id_fantasy = os.environ.get('BEST_STRATEGIES_FANTASY5_FILE_ID')
    
    if args.type == '539':
        print("🎯 僅分析 539...")
        process_single("539", FILE_539, OUTPUT_539, False, file_id=file_id_539, folder_id=folder_id, creds=creds)
    elif args.type == 'fantasy5':
        print("🎯 僅分析天天樂...")
        process_single("天天樂", FILE_FANTASY, OUTPUT_FANTASY, True, file_id=file_id_fantasy, folder_id=folder_id, creds=creds)
    else:
        print("🎯 分析所有彩球...")
        process_all()