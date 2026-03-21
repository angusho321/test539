# --------------------------------------------------------
#  verify_fantasy5_predictions.py - 加州Fantasy 5預測驗證功能
# --------------------------------------------------------

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta
import ast

def is_lottery_draw_day(check_date=None):
    """
    檢查指定日期是否為開獎日（加州Fantasy 5每日開獎，包括週日）
    Args:
        check_date: 要檢查的日期，None表示今天
    Returns:
        bool: True表示是開獎日，False表示不是
    """
    if check_date is None:
        check_date = datetime.now()
    elif isinstance(check_date, str):
        try:
            check_date = datetime.strptime(check_date, "%Y-%m-%d")
        except:
            print(f"日期格式錯誤: {check_date}")
            return True  # 預設為開獎日
    
    # 加州Fantasy 5每日開獎，包括週日
    is_draw_day = True
    
    weekday_names = ['週一', '週二', '週三', '週四', '週五', '週六', '週日']
    print(f"檢查日期: {check_date.strftime('%Y-%m-%d')} ({weekday_names[check_date.weekday()]})")
    
    if is_draw_day:
        print(f"{weekday_names[check_date.weekday()]} 是加州Fantasy 5開獎日（每日開獎）")
    else:
        print(f"⏸️ {weekday_names[check_date.weekday()]} 不開獎，跳過驗證")
    
    return is_draw_day

def load_latest_fantasy5_results(excel_path: str):
    """讀取最新的加州Fantasy 5開獎結果"""
    try:
        df = pd.read_excel(excel_path, engine='openpyxl')
        if len(df) == 0:
            print("加州Fantasy 5開獎資料檔案為空")
            return None
        
        # 取得最新一期的開獎結果
        latest_row = df.iloc[-1]
        latest_numbers = [latest_row['號碼1'], latest_row['號碼2'], latest_row['號碼3'], 
                         latest_row['號碼4'], latest_row['號碼5']]
        latest_date = latest_row.get('日期', '未知日期')
        
        print(f"最新加州Fantasy 5開獎結果 ({latest_date}):")
        print(f"   開獎號碼: {sorted(latest_numbers)}")
        return {
            'date': latest_date,
            'numbers': sorted(latest_numbers)
        }
    except Exception as e:
        print(f"讀取加州Fantasy 5開獎資料時發生錯誤: {e}")
        return None

def parse_prediction_numbers(prediction_str):
    """解析預測號碼字串為數字列表"""
    try:
        # 移除多餘的空白和換行
        prediction_str = prediction_str.strip()
        
        # 嘗試使用 ast.literal_eval 解析
        if prediction_str.startswith('[') and prediction_str.endswith(']'):
            return ast.literal_eval(prediction_str)
        
        # 如果不是標準格式，嘗試其他方式解析
        # 移除方括號和空白，分割數字
        numbers_str = prediction_str.replace('[', '').replace(']', '').replace(' ', '')
        numbers = [int(x) for x in numbers_str.split(',') if x.strip().isdigit()]
        return numbers
        
    except Exception as e:
        print(f"解析預測號碼時發生錯誤: {prediction_str} -> {e}")
        return []

def count_matching_numbers(prediction_numbers, actual_numbers):
    """計算預測號碼與實際開獎號碼的符合數量"""
    if not prediction_numbers or not actual_numbers:
        return 0
    
    prediction_set = set(prediction_numbers)
    actual_set = set(actual_numbers)
    matches = len(prediction_set.intersection(actual_set))
    return matches

def verify_fantasy5_predictions(prediction_log_file="fantasy5_prediction_log.xlsx", 
                      lottery_results_file="fantasy5_hist.xlsx",
                      days_to_verify=7):
    """
    驗證加州Fantasy 5預測結果
    Args:
        prediction_log_file: 預測記錄檔案
        lottery_results_file: 開獎結果檔案
        days_to_verify: 驗證過去幾天的預測記錄
    """
    print("開始驗證加州Fantasy 5預測結果...")
    
    # 檢查檔案是否存在
    if not Path(prediction_log_file).exists():
        print("找不到加州Fantasy 5預測記錄檔案")
        return
    
    if not Path(lottery_results_file).exists():
        print("找不到加州Fantasy 5開獎結果檔案")
        return
    
    # 讀取預測記錄
    try:
        predictions_df = pd.read_excel(prediction_log_file, engine='openpyxl')
        print(f"找到 {len(predictions_df)} 筆加州Fantasy 5預測記錄")
    except Exception as e:
        print(f"讀取預測記錄時發生錯誤: {e}")
        return
    
    # 讀取開獎結果資料
    try:
        lottery_df = pd.read_excel(lottery_results_file, engine='openpyxl')
        if len(lottery_df) == 0:
            print("加州Fantasy 5開獎結果檔案為空")
            return
        print(f"加州Fantasy 5開獎資料包含 {len(lottery_df)} 期結果")
    except Exception as e:
        print(f"讀取開獎結果時發生錯誤: {e}")
        return
    
    # 尋找需要驗證的記錄
    current_date = datetime.now()
    verification_count = 0
    updates_made = False
    
    for index, row in predictions_df.iterrows():
        # 檢查是否已經驗證過
        if '驗證結果' in row and pd.notna(row.get('驗證結果', '')) and row.get('驗證結果', '') != '':
            continue
        
        # 檢查預測日期是否在驗證範圍內
        try:
            prediction_date = pd.to_datetime(row['日期'])
            prediction_date_str = prediction_date.strftime('%Y-%m-%d')
            days_diff = (current_date - prediction_date).days
            
            if days_diff > days_to_verify:
                continue  # 超過驗證期限
                
        except Exception as e:
            print(f"解析預測日期時發生錯誤: {row.get('日期', 'N/A')} -> {e}")
            continue
        
        # 尋找對應日期的開獎結果
        matching_lottery = None
        for _, lottery_row in lottery_df.iterrows():
            try:
                lottery_date = pd.to_datetime(lottery_row['日期'])
                lottery_date_str = lottery_date.strftime('%Y-%m-%d')
                
                # 檢查日期是否匹配或預測日期之後有開獎
                if lottery_date_str == prediction_date_str or lottery_date > prediction_date:
                    matching_lottery = lottery_row
                    break
            except:
                continue
        
        if matching_lottery is None:
            print(f"{prediction_date_str} 的加州Fantasy 5預測尚無對應開獎結果，跳過驗證")
            continue
        
        # 提取開獎號碼
        try:
            actual_numbers = [
                matching_lottery['號碼1'], matching_lottery['號碼2'], 
                matching_lottery['號碼3'], matching_lottery['號碼4'], 
                matching_lottery['號碼5']
            ]
            actual_numbers = sorted([int(x) for x in actual_numbers if pd.notna(x)])
            actual_date = matching_lottery.get('日期', '未知日期')
        except Exception as e:
            print(f"解析開獎號碼時發生錯誤: {e}")
            continue
        
        print(f"\n驗證 {prediction_date_str} 的加州Fantasy 5預測...")
        print(f"   對應開獎: {actual_date} -> {actual_numbers}")
        
        # 驗證各策略的預測結果 - 支援簡化後的欄位格式
        # 檢查所有可能的策略欄位
        strategy_columns = [col for col in row.index if any(keyword in col for keyword in ['EV策略', '智能選號'])]
        verification_results = []
        max_matches = 0
        best_strategy = ""
        
        for strategy in strategy_columns:
            if pd.notna(row[strategy]) and str(row[strategy]).strip() != '':
                prediction_numbers = parse_prediction_numbers(str(row[strategy]))
                if prediction_numbers:
                    # 九顆策略：用全部9個號碼與開獎的5個號碼比對
                    # 七顆策略：用全部7個號碼與開獎的5個號碼比對
                    if '九顆' in strategy:
                        prediction_to_compare = sorted(prediction_numbers)  # 九顆用全部9個號碼
                    elif '七顆' in strategy:
                        prediction_to_compare = sorted(prediction_numbers)  # 七顆用全部7個號碼
                    else:
                        prediction_to_compare = sorted(prediction_numbers)  # 其他策略用全部號碼
                    
                    matches = count_matching_numbers(prediction_to_compare, actual_numbers)
                    verification_results.append(f"{strategy}:{matches}中")
                    
                    if matches > max_matches:
                        max_matches = matches
                        best_strategy = strategy
                    
                    print(f"   {strategy}: {prediction_to_compare} -> {matches}個號碼符合")
        
        # 更新驗證結果
        if verification_results:
            verification_text = f"開獎:{actual_numbers} | " + " | ".join(verification_results)
            if best_strategy:
                verification_text += f" | 最佳:{best_strategy}"
            
            predictions_df.at[index, '驗證結果'] = verification_text
            predictions_df.at[index, '中獎號碼數'] = max_matches
            verification_count += 1
            updates_made = True
    
    # 儲存更新後的記錄
    if updates_made:
        try:
            predictions_df.to_excel(prediction_log_file, index=False, engine='openpyxl')
            print(f"\n已更新 {verification_count} 筆加州Fantasy 5預測記錄的驗證結果")
        except Exception as e:
            print(f"儲存驗證結果時發生錯誤: {e}")
    else:
        print("\n沒有找到需要驗證的加州Fantasy 5記錄")
    
    # 顯示驗證統計
    show_verification_statistics(predictions_df)

def show_verification_statistics(predictions_df):
    """顯示驗證統計結果"""
    print(f"\n加州Fantasy 5驗證統計結果:")
    
    # 檢查是否有資料
    if predictions_df.empty:
        print("   尚無加州Fantasy 5預測記錄")
        return
    
    # 檢查是否有中獎號碼數欄位
    if '中獎號碼數' not in predictions_df.columns:
        print("   尚無已驗證的加州Fantasy 5記錄")
        return
    
    # 篩選已驗證的記錄
    verified_df = predictions_df[pd.notna(predictions_df['中獎號碼數']) & 
                                (predictions_df['中獎號碼數'] != '')]
    
    if len(verified_df) == 0:
        print("   尚無已驗證的加州Fantasy 5記錄")
        return
    
    print(f"   已驗證期數: {len(verified_df)}")
    
    # 統計各中獎號碼數的次數
    match_counts = verified_df['中獎號碼數'].value_counts().sort_index()
    for matches, count in match_counts.items():
        percentage = count / len(verified_df) * 100
        print(f"   {matches}個號碼符合: {count}次 ({percentage:.1f}%)")
    
    # 顯示最佳表現
    if len(verified_df) > 0:
        best_performance = verified_df['中獎號碼數'].max()
        best_records = verified_df[verified_df['中獎號碼數'] == best_performance]
        print(f"   最佳表現: {best_performance}個號碼符合 (共{len(best_records)}次)")

def auto_verify_on_startup(prediction_log_file="fantasy5_prediction_log.xlsx", 
                          lottery_results_file="fantasy5_hist.xlsx"):
    """程式啟動時自動驗證加州Fantasy 5預測結果"""
    print("\n🔄 自動檢查是否有需要驗證的加州Fantasy 5預測記錄...")
    
    if not Path(prediction_log_file).exists():
        print("尚無加州Fantasy 5預測記錄檔案")
        return
    
    if not Path(lottery_results_file).exists():
        print("尚無加州Fantasy 5開獎結果檔案")
        return
    
    # 檢查是否有可驗證的記錄
    try:
        predictions_df = pd.read_excel(prediction_log_file, engine='openpyxl')
        lottery_df = pd.read_excel(lottery_results_file, engine='openpyxl')
        
        if len(predictions_df) == 0:
            print("加州Fantasy 5預測記錄檔案為空")
            return
            
        if len(lottery_df) == 0:
            print("加州Fantasy 5開獎結果檔案為空")
            return
        
        # 檢查是否有未驗證的記錄
        unverified_count = predictions_df[
            (predictions_df['驗證結果'].isna()) | 
            (predictions_df['驗證結果'] == '')
        ].shape[0]
        
        if unverified_count == 0:
            print("所有加州Fantasy 5預測記錄都已驗證")
            return
        
        print(f"找到 {unverified_count} 筆未驗證的加州Fantasy 5預測記錄")
        
        # 檢查最新開獎日期
        latest_lottery_date = pd.to_datetime(lottery_df['日期']).max()
        today = pd.Timestamp.now().normalize()
        
        print(f"最新加州Fantasy 5開獎日期: {latest_lottery_date.strftime('%Y-%m-%d')}")
        print(f"今日日期: {today.strftime('%Y-%m-%d')}")
        
        # 加州Fantasy 5每日開獎，包括週日
        if latest_lottery_date < today:
            verify_fantasy5_predictions(prediction_log_file, lottery_results_file, days_to_verify=7)
        else:
            print("⏰ 今日尚無加州Fantasy 5開獎結果，跳過自動驗證")
            
    except Exception as e:
        print(f"自動驗證檢查時發生錯誤: {e}")
        return

if __name__ == "__main__":
    print("加州Fantasy 5預測驗證系統")
    print("="*50)
    
    # 執行驗證
    verify_fantasy5_predictions()
    
    print("\n" + "="*50)
    print("使用說明:")
    print("1. 每次有新的加州Fantasy 5開獎結果時，執行此程式進行驗證")
    print("2. 程式會自動比對預測記錄與最新開獎結果")
    print("3. 驗證結果會自動更新到 fantasy5_prediction_log.xlsx")
    print("4. 可以調整 days_to_verify 參數來控制驗證範圍")
