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

def calculate_strategy_numbers(monday_record, lottery_type):
    """根據週一開獎記錄計算策略號碼"""
    if lottery_type == '539':
        # 539: A = 週一第1支 + 06, B = 週一第2支 + 12
        num1 = int(monday_record['號碼1'])
        num2 = int(monday_record['號碼2'])
        
        A = (num1 + 6) % 39
        if A == 0:
            A = 39
        B = (num2 + 12) % 39
        if B == 0:
            B = 39
        
        return A, B
    else:  # fantasy5
        # Fantasy5: A = 週一第1支 + 13, B = 週一第4支
        num1 = int(monday_record['號碼1'])
        num4 = int(monday_record['號碼4'])
        
        A = (num1 + 13) % 39
        if A == 0:
            A = 39
        B = num4  # 直接沿用
        
        return A, B

def get_target_weekdays(lottery_type):
    """取得目標追號期（週二至週六）"""
    return [1, 2, 3, 4, 5]  # 週二至週六

def backtest_strategy(df, monday_records, lottery_type, weeks=52):
    """回測策略近一年的勝率"""
    if monday_records.empty or len(monday_records) < 2:
        return 0.0, 0, 0
    
    # 只取最近52週的週一記錄
    recent_mondays = monday_records.tail(weeks).copy()
    
    wins = 0
    total = 0
    target_weekdays = get_target_weekdays(lottery_type)
    
    for idx, monday_row in recent_mondays.iterrows():
        monday_date = monday_row['日期']
        A, B = calculate_strategy_numbers(monday_row, lottery_type)
        
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

def check_current_week_status(df, latest_monday, lottery_type):
    """檢查本週狀態"""
    if latest_monday is None:
        return "無資料", None, None
    
    monday_date = latest_monday['日期']
    A, B = calculate_strategy_numbers(latest_monday, lottery_type)
    
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
    
    # 計算策略號碼
    A, B = calculate_strategy_numbers(latest_monday, lottery_type)
    print(f"🎯 本週策略號碼: A={A}, B={B}")
    
    # 回測近一年勝率
    win_rate, wins, total = backtest_strategy(df, monday_records, lottery_type, weeks=52)
    print(f"📈 近一年勝率: {win_rate:.1f}% ({wins}/{total})")
    
    # 檢查本週狀態
    status, win_date, win_record = check_current_week_status(df, latest_monday, lottery_type)
    print(f"📋 本週狀態: {status}")
    
    # 準備寫入 Excel 的資料
    strategy_data = {
        '項目': [
            '策略名稱',
            '本週一日期',
            '週一第1支',
            '週一第2支',
            '週一第4支',
            '策略號碼A',
            '策略號碼B',
            '追號期間',
            '近一年勝率',
            '近一年中獎次數',
            '近一年總週數',
            '本週狀態',
            '中獎日期',
            '更新時間'
        ],
        '內容': [
            '週一冠軍策略',
            latest_monday_date.strftime('%Y-%m-%d'),
            int(latest_monday['號碼1']),
            int(latest_monday['號碼2']),
            int(latest_monday['號碼4']) if lottery_type == 'fantasy5' else 'N/A',
            A,
            B,
            '週二至週六',
            f'{win_rate:.1f}%',
            wins,
            total,
            status,
            win_date.strftime('%Y-%m-%d') if win_date else 'N/A',
            datetime.now().strftime('%Y-%m-%d %H:%M:%S')
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

