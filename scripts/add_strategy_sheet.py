#!/usr/bin/env python3
"""
週一冠軍策略分析腳本
將分析結果寫入 Excel 的 Monday_Strategy 分頁
"""

import pandas as pd
import numpy as np
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import load_workbook
import sys

def load_lottery_data(file_path):
    """讀取彩票歷史資料"""
    if not Path(file_path).exists():
        print(f"❌ 檔案不存在: {file_path}")
        return None
    
    try:
        df = pd.read_excel(file_path, engine='openpyxl', sheet_name='Sheet1')
    except:
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
        except Exception as e:
            print(f"❌ 讀取 Excel 失敗: {e}")
            return None
    
    # 處理日期欄位
    if '日期' in df.columns:
        try:
            df['日期'] = pd.to_datetime(df['日期'], format='mixed', errors='coerce')
            if df['日期'].isna().any():
                df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
            if df['日期'].isna().any():
                df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d', errors='coerce')
        except:
            df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
        
        df = df.dropna(subset=['日期'])
        df = df.sort_values('日期').reset_index(drop=True)
    
    return df

def get_monday_records(df):
    """取得所有週一的開獎記錄"""
    if df is None or df.empty:
        return pd.DataFrame()
    
    # 確保日期是 datetime
    df['weekday'] = df['日期'].dt.weekday  # 0=週一, 6=週日
    monday_records = df[df['weekday'] == 0].copy()
    return monday_records.sort_values('日期').reset_index(drop=True)

def calculate_number_with_offset(base_number, offset):
    """
    計算新號碼：基準號碼 + 加數
    特殊處理：若相加結果 > 39，則結果 = 結果 - 39
    """
    result = base_number + offset
    if result > 39:
        result = result - 39
    return result

def calculate_strategy_numbers(monday_record, lottery_type, offset_a, offset_b):
    """
    根據週一開獎記錄計算策略號碼
    必須提供 offset_a 和 offset_b 參數
    """
    if offset_a is None or offset_b is None:
        raise ValueError("offset_a 和 offset_b 必須提供，不能為 None")
    
    if lottery_type == '539':
        num1 = int(monday_record['號碼1'])
        num2 = int(monday_record['號碼2'])
        
        A = calculate_number_with_offset(num1, offset_a)
        B = calculate_number_with_offset(num2, offset_b)
        
        return A, B
    else:  # fantasy5
        num1 = int(monday_record['號碼1'])
        num4 = int(monday_record['號碼4'])
        
        A = calculate_number_with_offset(num1, offset_a)
        # Fantasy5 的 B = 週一第4支 + Offset_B
        # 當 offset_b = 0 時，相當於直接沿用第4支
        B = calculate_number_with_offset(num4, offset_b)
        
        return A, B

def get_target_weekdays(lottery_type):
    """取得目標追號期（週二至週六）"""
    return [1, 2, 3, 4, 5]  # 週二至週六

def backtest_strategy(df, monday_records, lottery_type, offset_a=None, offset_b=None, weeks=None):
    """
    回測策略近一年的勝率
    如果提供了 offset_a 和 offset_b，則使用這些偏移量進行回測
    如果 weeks 為 None，則使用所有傳入的 monday_records
    """
    if monday_records.empty or len(monday_records) < 2:
        return 0.0, 0, 0
    
    # 如果指定了 weeks，則只取最近 N 週的週一記錄；否則使用全部
    if weeks is not None:
        recent_mondays = monday_records.tail(weeks).copy()
    else:
        recent_mondays = monday_records.copy()
    
    wins = 0
    total = 0
    target_weekdays = get_target_weekdays(lottery_type)
    
    for idx, monday_row in recent_mondays.iterrows():
        monday_date = monday_row['日期']
        A, B = calculate_strategy_numbers(monday_row, lottery_type, offset_a, offset_b)
        
        # 找出這個週一之後的週二至週六開獎記錄
        week_start = monday_date
        week_end = monday_date + timedelta(days=6)
        
        week_records = df[
            (df['日期'] > week_start) & 
            (df['日期'] <= week_end) &
            (df['日期'].dt.weekday.isin(target_weekdays))
        ].copy()
        
        if week_records.empty:
            continue
        
        total += 1
        
        # 檢查是否中獎（A 或 B 出現在任何一天的開獎號碼中）
        for _, record in week_records.iterrows():
            drawn_numbers = [
                int(record['號碼1']),
                int(record['號碼2']),
                int(record['號碼3']),
                int(record['號碼4']),
                int(record['號碼5'])
            ]
            
            if A in drawn_numbers or B in drawn_numbers:
                wins += 1
                break  # 只要有一期中獎就算這週中獎
    
    win_rate = (wins / total * 100) if total > 0 else 0.0
    return win_rate, wins, total

def find_best_strategies(df, monday_records, lottery_type, weeks=52, min_win_rate=90.0):
    """
    動態分析過去一年的歷史數據，找出勝率超過指定閾值的最佳策略組合
    返回前兩名最佳策略
    """
    if monday_records.empty or len(monday_records) < 2:
        return []
    
    print(f"🔍 開始動態分析所有可能的策略組合（Offset 範圍: 0-38）...")
    
    # 只取最近52週的週一記錄
    recent_mondays = monday_records.tail(weeks).copy()
    
    if recent_mondays.empty:
        return []
    
    # 嘗試所有可能的 Offset 組合（0-38）
    all_strategies = []
    total_combinations = 39 * 39  # 39 * 39 = 1521 種組合
    processed = 0
    
    for offset_a in range(0, 39):
        for offset_b in range(0, 39):
            processed += 1
            if processed % 100 == 0:
                progress = (processed / total_combinations) * 100
                print(f"   進度: {progress:.1f}% ({processed}/{total_combinations})", end='\r', flush=True)
            
            # 回測這個策略組合（使用所有 recent_mondays，因為已經過濾為最近52週）
            win_rate, wins, total = backtest_strategy(
                df, recent_mondays, lottery_type, 
                offset_a=offset_a, offset_b=offset_b, 
                weeks=None  # 使用所有傳入的 monday_records（已經過濾為最近52週）
            )
            
            # 只保留勝率超過閾值的策略
            if win_rate >= min_win_rate and total > 0:
                all_strategies.append({
                    'offset_a': offset_a,
                    'offset_b': offset_b,
                    'win_rate': win_rate,
                    'wins': wins,
                    'total': total
                })
    
    print(f"\n   完成！找到 {len(all_strategies)} 組勝率 >= {min_win_rate}% 的策略")
    
    # 排序：先按勝率降序，再按中獎次數降序
    all_strategies.sort(key=lambda x: (-x['win_rate'], -x['wins']))
    
    # 返回前兩名
    return all_strategies[:2]

def check_current_week_status(df, latest_monday, lottery_type, offset_a, offset_b):
    """檢查本週狀態（使用指定的 offset）"""
    if latest_monday is None:
        return "無資料", None, None
    
    monday_date = latest_monday['日期']
    A, B = calculate_strategy_numbers(latest_monday, lottery_type, offset_a, offset_b)
    
    # 找出本週的週二至週六開獎記錄
    week_start = monday_date
    week_end = monday_date + timedelta(days=6)
    target_weekdays = get_target_weekdays(lottery_type)
    
    week_records = df[
        (df['日期'] > week_start) & 
        (df['日期'] <= week_end) &
        (df['日期'].dt.weekday.isin(target_weekdays))
    ].copy()
    
    if week_records.empty:
        return "等待開獎", None, None
    
    # 檢查是否已開獎
    today = datetime.now().date()
    latest_record_date = week_records['日期'].max().date()
    
    # 檢查是否中獎
    for _, record in week_records.iterrows():
        drawn_numbers = [
            int(record['號碼1']),
            int(record['號碼2']),
            int(record['號碼3']),
            int(record['號碼4']),
            int(record['號碼5'])
        ]
        
        if A in drawn_numbers or B in drawn_numbers:
            win_date = record['日期'].date()
            return "已中獎", win_date, record
    
    # 如果本週已過完但沒中獎
    if latest_record_date >= (monday_date + timedelta(days=5)).date():
        return "未中獎", None, None
    
    return "等待開獎", None, None

def add_strategy_sheet(file_path, lottery_type):
    """將策略分析結果寫入 Excel 的新分頁"""
    print(f"📊 開始分析 {lottery_type} 的週一冠軍策略...")
    
    # 讀取資料
    df = load_lottery_data(file_path)
    if df is None:
        return False
    
    print(f"✅ 成功讀取 {len(df)} 筆歷史記錄")
    
    # 取得週一記錄
    monday_records = get_monday_records(df)
    if monday_records.empty:
        print("❌ 沒有找到週一的開獎記錄")
        return False
    
    print(f"📅 找到 {len(monday_records)} 筆週一開獎記錄")
    
    # 取得最新的週一記錄
    latest_monday = monday_records.iloc[-1]
    latest_monday_date = latest_monday['日期']
    
    # 動態分析找出最佳策略（勝率 > 90%）
    best_strategies = find_best_strategies(df, monday_records, lottery_type, weeks=52, min_win_rate=90.0)
    
    # 準備最佳策略字串和號碼
    first_strategy_str = "無符合策略"
    second_strategy_str = "無符合策略"
    A, B = None, None
    C, D = None, None
    first_win_rate = 0.0
    second_win_rate = 0.0
    
    if len(best_strategies) >= 1:
        s1 = best_strategies[0]
        # 計算第一組的實際號碼
        A, B = calculate_strategy_numbers(latest_monday, lottery_type, s1['offset_a'], s1['offset_b'])
        first_win_rate = s1['win_rate']
        first_strategy_str = f"{A} {B} {first_win_rate:.1f}%"
        print(f"🏆 第一組最佳策略: 號碼A={A}, 號碼B={B}, 勝率={first_win_rate:.1f}% (中獎: {s1['wins']}/{s1['total']})")
    
    if len(best_strategies) >= 2:
        s2 = best_strategies[1]
        # 計算第二組的實際號碼
        C, D = calculate_strategy_numbers(latest_monday, lottery_type, s2['offset_a'], s2['offset_b'])
        second_win_rate = s2['win_rate']
        second_strategy_str = f"{C} {D} {second_win_rate:.1f}%"
        print(f"🥈 第二組最佳策略: 號碼C={C}, 號碼D={D}, 勝率={second_win_rate:.1f}% (中獎: {s2['wins']}/{s2['total']})")
    
    # 使用第一組最佳策略計算本週預測號碼（如果有）
    if len(best_strategies) >= 1:
        best_offset_a = best_strategies[0]['offset_a']
        best_offset_b = best_strategies[0]['offset_b']
        print(f"🎯 本週預測號碼（使用第一組策略）: A={A}, B={B}")
        
        # 檢查本週狀態（使用第一組最佳策略）
        status, win_date, win_record = check_current_week_status(
            df, latest_monday, lottery_type, best_offset_a, best_offset_b
        )
        win_rate, wins, total = first_win_rate, best_strategies[0]['wins'], best_strategies[0]['total']
    else:
        # 如果沒有找到最佳策略，無法計算預測號碼
        print("⚠️ 未找到勝率 >= 90% 的策略，無法計算本週預測號碼")
        win_rate, wins, total = 0.0, 0, 0
        status, win_date, win_record = "無符合策略", None, None
    
    print(f"📋 本週狀態: {status}")
    
    # 準備寫入 Excel 的資料
    strategy_data = {
        '項目': [
            '第一組',
            '第二組'
        ],
        '內容': [
            first_strategy_str,
            second_strategy_str
        ]
    }
    
    strategy_df = pd.DataFrame(strategy_data)
    
    # 使用 openpyxl 來處理 Excel（保留原有分頁）
    try:
        book = load_workbook(file_path)
        
        # 如果 Monday_Strategy 分頁已存在，刪除它
        if 'Monday_Strategy' in book.sheetnames:
            print("🔄 刪除舊的 Monday_Strategy 分頁...")
            del book['Monday_Strategy']
        
        # 創建新的分頁
        from openpyxl.utils.dataframe import dataframe_to_rows
        ws = book.create_sheet('Monday_Strategy')
        
        # 寫入標題（加粗）
        from openpyxl.styles import Font
        ws['A1'] = '項目'
        ws['B1'] = '內容'
        ws['A1'].font = Font(bold=True)
        ws['B1'].font = Font(bold=True)
        
        # 寫入資料
        for r_idx, row in enumerate(dataframe_to_rows(strategy_df, index=False, header=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)
        
        # 調整欄寬
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 30
        
        # 保存
        book.save(file_path)
        print(f"✅ 成功將策略分析結果寫入 {file_path} 的 Monday_Strategy 分頁")
        return True
        
    except Exception as e:
        print(f"⚠️ 使用 openpyxl 寫入失敗: {e}")
        print("   嘗試使用 pandas 方法...")
        # 如果 openpyxl 方法失敗，嘗試使用 pandas
        try:
            # 先讀取原有資料
            existing_sheets = {}
            try:
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                for sheet_name in excel_file.sheet_names:
                    if sheet_name != 'Monday_Strategy':
                        existing_sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            except:
                pass
            
            # 寫入所有分頁（包括新的 Monday_Strategy）
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
                # 寫入原有分頁
                for sheet_name, sheet_df in existing_sheets.items():
                    sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                # 寫入新的策略分頁
                strategy_df.to_excel(writer, sheet_name='Monday_Strategy', index=False)
            
            print(f"✅ 成功將策略分析結果寫入 {file_path} 的 Monday_Strategy 分頁（使用 pandas）")
            return True
        except Exception as e2:
            print(f"❌ 使用 pandas 寫入也失敗: {e2}")
            import traceback
            traceback.print_exc()
            return False

def main():
    parser = argparse.ArgumentParser(description='週一冠軍策略分析')
    parser.add_argument('--type', choices=['539', 'fantasy5'], required=True, help='彩球類型')
    parser.add_argument('--file', required=True, help='Excel 檔案路徑')
    
    args = parser.parse_args()
    
    lottery_type = args.type
    file_path = args.file
    
    success = add_strategy_sheet(file_path, lottery_type)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()

