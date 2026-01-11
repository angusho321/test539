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

OUTPUT_539 = 'best_strategies_539.xlsx'
OUTPUT_FANTASY = 'best_strategies_fantasy5.xlsx'

# 時間段定義 (對應 Python weekday: 0=週一 ... 6=週日)
TIME_WINDOWS_539 = {
    "周一至周三": [0, 1, 2],  # 週一、週二、週三
    "周二至周四": [1, 2, 3],  # 週二、週三、週四
    "周三至周五": [2, 3, 4],  # 週三、週四、週五
    "周四至周六": [3, 4, 5]   # 週四、週五、週六
}

TIME_WINDOWS_FANTASY = {
    "周一至周三": [0, 1, 2],  # 週一、週二、週三
    "周二至周四": [1, 2, 3],  # 週二、週三、週四
    "周三至周五": [2, 3, 4],  # 週三、週四、週五
    "周四至周六": [3, 4, 5],  # 週四、週五、週六
    "周五至周日": [4, 5, 6]   # 週五、週六、週日（僅天天樂）
}

def get_time_windows(is_fantasy=False):
    """根據彩球類型返回對應的時間段"""
    return TIME_WINDOWS_FANTASY if is_fantasy else TIME_WINDOWS_539

# 使用近一年的紀錄
RECENT_YEARS = 1

# ==========================================
# 核心演算法
# ==========================================

def load_data(file_path, is_fantasy=False):
    """讀取資料並處理時區，只保留近一年的紀錄"""
    df = None
    
    # 先嘗試讀取 Excel
    if os.path.exists(file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
        except Exception as e:
            print(f"   ⚠️ 讀取 Excel 失敗: {e}，嘗試 CSV...")
    
    # 如果 Excel 不存在或讀取失敗，嘗試 CSV
    if df is None:
        # 嘗試多種可能的 CSV 檔名
        csv_paths = [
            file_path.replace('.xlsx', '.csv'),
            file_path.replace('.xlsx', ' - Sheet1.csv'),
            file_path + ' - Sheet1.csv'
        ]
        
        for csv_path in csv_paths:
            if os.path.exists(csv_path):
                try:
                    df = pd.read_csv(csv_path, encoding='utf-8-sig')
                    print(f"   📄 從 CSV 讀取: {csv_path}")
                    break
                except Exception as e:
                    print(f"   ⚠️ 讀取 CSV 失敗 ({csv_path}): {e}")
                    continue
    
    if df is None:
        print(f"❌ 無法讀取文件: {file_path}")
        return None

    # 處理日期欄位：支援多種日期格式（包含或不包含時間）
    try:
        df['日期'] = pd.to_datetime(df['日期'], format='mixed', errors='coerce')
        if df['日期'].isna().any():
            df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
        if df['日期'].isna().any():
            df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d', errors='coerce')
    except:
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
    
    # 計算近一年的日期範圍
    max_date = df['Analysis_Date'].max()
    cutoff_date = max_date - pd.Timedelta(days=365 * RECENT_YEARS)
    
    # 只保留近一年的紀錄
    df = df[df['Analysis_Date'] >= cutoff_date].copy()
    
    # 按日期排序（最舊的在前）
    df = df.sort_values('Analysis_Date', ascending=True).reset_index(drop=True)
    
    print(f"   📊 已載入 {len(df)} 筆近一年紀錄")
    if len(df) > 0:
        print(f"   📅 日期範圍: {df['Analysis_Date'].min()} 至 {df['Analysis_Date'].max()}")
    
    return df

def extract_numbers(row, is_fantasy=False):
    """從資料列中提取號碼"""
    try:
        if '號碼1' in row.index:
            nums = [int(row['號碼1']), int(row['號碼2']), int(row['號碼3']), 
                   int(row['號碼4']), int(row['號碼5'])]
        else:
            # 嘗試其他可能的欄位名稱
            cols = row.index.tolist()
            nums = [int(row[cols[2]]), int(row[cols[3]]), int(row[cols[4]]), 
                   int(row[cols[5]]), int(row[cols[6]])]
        return set(nums)
    except:
        return None

def calculate_window_win_rate(df, window_name, window_days, is_fantasy=False):
    """
    計算指定時間段的勝率（優化版本）
    規則：三個號碼中任一個在該時間段（三天）的任一天出現即算中獎
    
    返回: [{'combo': tuple, 'win_rate': float, 'wins': int, 'total': int, 'missed_dates': list}, ...]
    """
    # 過濾出該時間段的資料
    window_data = df[df['Analysis_Date'].dt.weekday.isin(window_days)].copy()
    
    if len(window_data) == 0:
        return []
    
    # 預先提取所有號碼集合，避免重複計算
    window_data['Numbers'] = window_data.apply(
        lambda row: extract_numbers(row, is_fantasy), axis=1
    )
    window_data = window_data[window_data['Numbers'].notna()].copy()
    
    # 將資料按週分組，並預先計算每週的號碼聯集
    window_data['YearWeek'] = window_data['Analysis_Date'].apply(
        lambda x: (x.isocalendar()[0], x.isocalendar()[1])
    )
    
    # 預先計算每週的號碼聯集（該週時間段內所有開出的號碼）
    # 同時記錄每週的時間段第一天日期
    week_unions = {}
    week_first_dates = {}  # 記錄每週的時間段第一天日期
    
    # 時間段的第一天（weekday最小的那天）
    first_weekday = min(window_days)
    
    for (year, week), group in window_data.groupby('YearWeek'):
        union_set = set()
        for nums in group['Numbers']:
            if nums:
                union_set.update(nums)
        week_unions[(year, week)] = union_set
        
        # 找到該週中時間段的第一天（weekday為first_weekday的那天）
        first_day_records = group[group['Analysis_Date'].dt.weekday == first_weekday]
        if len(first_day_records) > 0:
            first_day = first_day_records['Analysis_Date'].min()
            week_first_dates[(year, week)] = first_day.date()
        else:
            # 如果該週沒有第一天（例如該週的第一天還沒開獎），使用該週時間段內最早的那天
            first_day = group['Analysis_Date'].min()
            week_first_dates[(year, week)] = first_day.date()
    
    total_weeks = len(week_unions)
    if total_weeks == 0:
        return []
    
    # 獲取所有可能的3碼組合
    max_num = 39 if not is_fantasy else 39  # 539和Fantasy5都是1-39
    all_combos = list(combinations(range(1, max_num + 1), 3))
    total_combos = len(all_combos)
    
    results = []
    
    # 顯示進度
    print(f"         計算中... (共 {total_combos} 組組合, {total_weeks} 週)", end='', flush=True)
    
    for idx, combo in enumerate(all_combos):
        # 每1000個組合顯示一次進度
        if idx % 1000 == 0 and idx > 0:
            progress = (idx / total_combos) * 100
            print(f"\r         進度: {progress:.1f}% ({idx}/{total_combos})", end='', flush=True)
        
        combo_set = set(combo)
        wins = 0
        missed_dates = []  # 記錄未中獎的時間段第一天日期
        
        # 使用預先計算的週聯集，快速判斷
        for (year, week), week_union in week_unions.items():
            # 如果組合與該週的號碼聯集有交集，則中獎
            if not combo_set.isdisjoint(week_union):
                wins += 1
            else:
                # 未中獎，記錄該週的時間段第一天日期
                first_date = week_first_dates.get((year, week))
                if first_date:
                    missed_dates.append(first_date)
        
        win_rate = wins / total_weeks
        results.append({
            'combo': combo,
            'win_rate': win_rate,
            'wins': wins,
            'total': total_weeks,
            'missed_dates': sorted(missed_dates)  # 按日期排序
        })
    
    # 按勝率排序
    results.sort(key=lambda x: x['win_rate'], reverse=True)
    print(f"\r         完成！找到 {len(results)} 組結果" + " " * 40)  # 清除進度顯示
    
    # 先取前30名進行去重處理（確保有足夠的候選）
    top_results = results[:30]
    
    # 移除重複兩碼組合的策略
    deduplicated_results = remove_duplicate_two_ball_combos(top_results)
    
    # 返回去重後的前10名
    return deduplicated_results[:10]

def remove_duplicate_two_ball_combos(results):
    """
    移除重複兩碼組合的策略，只保留勝率最高的
    例如：07,22,24 和 07,24,28 都有 07,24，只保留勝率較高的
    如果勝率相同，只保留第一個（按原始順序）
    """
    if not results:
        return results
    
    # 建立兩碼組合到最佳三碼組合的映射
    two_ball_to_best = {}  # {(ball1, ball2): best_result}
    
    # 第一遍：找出每個兩碼組合對應的最佳三碼組合
    for result in results:
        combo = result['combo']
        # 提取所有可能的兩碼子組合（C(3,2) = 3個）
        two_ball_combos = list(combinations(combo, 2))
        
        for two_ball in two_ball_combos:
            # 排序兩碼組合，確保一致性
            two_ball_sorted = tuple(sorted(two_ball))
            
            # 如果這個兩碼組合還沒記錄，或當前組合勝率更高，則更新
            if two_ball_sorted not in two_ball_to_best:
                two_ball_to_best[two_ball_sorted] = result
            else:
                # 比較勝率，保留勝率更高的
                current_best = two_ball_to_best[two_ball_sorted]
                if result['win_rate'] > current_best['win_rate']:
                    two_ball_to_best[two_ball_sorted] = result
                # 如果勝率相同，比較中獎次數
                elif result['win_rate'] == current_best['win_rate']:
                    if result['wins'] > current_best['wins']:
                        two_ball_to_best[two_ball_sorted] = result
                    # 如果勝率和中獎次數都相同，保留第一個（不更新）
    
    # 第二遍：找出所有應該被保留的三碼組合
    # 一個三碼組合要被保留，當且僅當它至少有一個兩碼子組合是該兩碼組合的最佳選擇
    kept_combos = set()
    for two_ball, best_result in two_ball_to_best.items():
        kept_combos.add(best_result['combo'])
    
    # 收集所有被保留的組合，保持原始順序
    final_results = []
    seen_combos = set()
    for result in results:
        if result['combo'] in kept_combos and result['combo'] not in seen_combos:
            final_results.append(result)
            seen_combos.add(result['combo'])
    
    # 按勝率重新排序（確保順序正確）
    final_results.sort(key=lambda x: (x['win_rate'], x['wins']), reverse=True)
    
    return final_results

def format_combo_result(result):
    """格式化組合結果"""
    combo_str = ",".join(f"{x:02d}" for x in result['combo'])
    win_rate_pct = result['win_rate'] * 100
    return f"{combo_str} [{win_rate_pct:.1f}% ({result['wins']}/{result['total']})]"

def format_missed_dates(missed_dates):
    """格式化未中獎日期（顯示所有日期）"""
    if not missed_dates:
        return ""
    
    # 格式化為 YYYY/MM/DD，顯示所有日期
    formatted_dates = [date.strftime('%Y/%m/%d') if hasattr(date, 'strftime') else str(date) for date in missed_dates]
    
    return ", ".join(formatted_dates)

def generate_predictions(df, is_fantasy=False):
    """
    生成所有時間段的預測
    返回: DataFrame，包含各時間段的勝率前十名和槓龜日期
    """
    print(f"   🔍 開始計算各時間段勝率...")
    
    # 根據彩球類型獲取對應的時間段
    time_windows = get_time_windows(is_fantasy)
    
    # 計算每個時間段的勝率前十名
    window_results = {}
    for window_name, window_days in time_windows.items():
        print(f"      -> 計算 {window_name}...")
        results = calculate_window_win_rate(df, window_name, window_days, is_fantasy)
        window_results[window_name] = results
    
    # 構建輸出 DataFrame
    # 找出最長的結果列表（最多10個）
    max_len = max(len(results) for results in window_results.values()) if window_results else 0
    
    # 動態創建輸出數據（根據時間段，每個時間段後面加一個槓龜欄位）
    # 構建欄位順序和數據
    columns = []
    data_dict = {}
    window_to_missed = {}  # 建立時間段名稱到槓龜欄位名稱的對應關係
    
    # 為每個時間段創建兩個欄位：時間段名稱和「槓龜1」、「槓龜2」等
    missed_counter = 1
    for window_name in time_windows.keys():
        columns.append(window_name)
        missed_col_name = f"槓龜{missed_counter}"
        columns.append(missed_col_name)
        # 初始化數據列表
        data_dict[window_name] = []
        data_dict[missed_col_name] = []
        # 記錄對應關係
        window_to_missed[window_name] = missed_col_name
        missed_counter += 1
    
    for i in range(max_len):
        for window_name in time_windows.keys():
            missed_col_name = window_to_missed[window_name]
            if i < len(window_results[window_name]):
                result = window_results[window_name][i]
                data_dict[window_name].append(format_combo_result(result))
                # 格式化槓龜日期（顯示所有日期）
                missed_dates = result.get('missed_dates', [])
                data_dict[missed_col_name].append(format_missed_dates(missed_dates))
            else:
                data_dict[window_name].append("")
                data_dict[missed_col_name].append("")
    
    # 創建 DataFrame
    result_df = pd.DataFrame({col: data_dict[col] for col in columns})
    result_df = result_df[columns]  # 確保欄位順序正確
    
    return result_df

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
        # 根據文件擴展名設置正確的 MIME type
        if local_file.endswith('.xlsx'):
            upload_mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        elif local_file.endswith('.csv'):
            upload_mime_type = 'text/csv'
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
        else:
            # 其他文件類型，嘗試推斷 MIME type
            upload_mime_type = 'application/octet-stream'
        
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
    
    # 載入資料（只保留最新365筆）
    df = load_data(input_file, is_fantasy)
    if df is None or len(df) == 0:
        print(f"❌ 找不到或無法讀取 {input_file}，跳過")
        return False
    
    # 生成預測
    result_df = generate_predictions(df, is_fantasy)
    
    # 輸出 XLSX
    try:
        result_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"📄 已建立: {output_file}")
    except Exception as e:
        print(f"❌ 建立文件失敗: {e}")
        return False
    
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