#!/usr/bin/env python3
"""
使用 TaiwanLotteryCrawler 套件更新完整的 539 開獎歷史記錄
"""

import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import sys
import os

# 添加父目錄到路徑
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def get_complete_539_history():
    """使用 TaiwanLotteryCrawler 獲取完整的 539 歷史記錄"""
    
    try:
        from TaiwanLottery import TaiwanLotteryCrawler
        print("✅ 成功載入 TaiwanLotteryCrawler 套件")
    except ImportError:
        print("❌ 請先安裝 TaiwanLotteryCrawler 套件:")
        print("   pip install taiwanlottery")
        return None
    
    lottery = TaiwanLotteryCrawler()
    all_data = []
    
    # 獲取從 2019 年到現在的所有記錄（539 彩券開始時間）
    current_date = datetime.now()
    start_year = 2019  # 539 彩券大約從 2019 年開始
    
    print(f"🔍 開始爬取 {start_year} 年至 {current_date.year} 年的 539 開獎記錄...")
    
    for year in range(start_year, current_date.year + 1):
        for month in range(1, 13):
            # 如果是當前年份，只爬取到當前月份
            if year == current_date.year and month > current_date.month:
                break
            # 如果是未來的年份，跳過
            if year > current_date.year:
                break
                
            try:
                print(f"📅 爬取 {year}-{month:02d} 的記錄...")
                
                # 使用 daily_cash 方法獲取 539 記錄
                month_data = lottery.daily_cash([str(year), f"{month:02d}"])
                
                if month_data and len(month_data) > 0:
                    # 直接將 list 加入 all_data
                    all_data.extend(month_data)
                    print(f"   ✅ 獲取 {len(month_data)} 筆記錄")
                else:
                    print(f"   ⚠️ 該月份無記錄")
                    
            except Exception as e:
                print(f"   ❌ 爬取 {year}-{month:02d} 時發生錯誤: {e}")
                continue
    
    if not all_data:
        print("❌ 沒有獲取到任何記錄")
        return None
    
    print(f"\n📊 總共獲取 {len(all_data)} 筆 539 開獎記錄")
    
    # 顯示第一筆記錄的結構
    if all_data:
        print("\n📋 記錄格式範例:")
        print(f"   {all_data[0]}")
    
    return all_data

def standardize_lottery_data(raw_data):
    """標準化彩票資料格式，使其與現有系統相容"""
    
    if raw_data is None or len(raw_data) == 0:
        return None
    
    print("\n🔧 標準化資料格式...")
    
    # 顯示原始格式
    print("原始格式範例:", raw_data[0] if raw_data else "無資料")
    
    # 根據 TaiwanLotteryCrawler 的輸出格式進行標準化
    standardized_data = []
    
    for index, record in enumerate(raw_data):
        try:
            # 解析開獎日期
            draw_date = pd.to_datetime(record['開獎日期'])
            
            # 解析獎號
            winning_numbers = record['獎號']
            
            # 確保有 5 個號碼
            if len(winning_numbers) >= 5:
                # 計算星期
                weekday_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
                weekday = weekday_map[draw_date.weekday()]
                
                standardized_record = {
                    '日期': draw_date,
                    '星期': weekday,
                    '號碼1': winning_numbers[0],
                    '號碼2': winning_numbers[1],
                    '號碼3': winning_numbers[2],
                    '號碼4': winning_numbers[3],
                    '號碼5': winning_numbers[4],
                    '期別': record.get('期別', '')
                }
                
                standardized_data.append(standardized_record)
                
        except Exception as e:
            print(f"⚠️ 處理第 {index} 筆記錄時發生錯誤: {e}")
            print(f"   記錄內容: {record}")
            continue
    
    if not standardized_data:
        print("❌ 沒有成功標準化的記錄")
        return None
    
    standardized_df = pd.DataFrame(standardized_data)
    
    # 排序資料
    standardized_df = standardized_df.sort_values('日期')
    
    print(f"✅ 標準化完成，共 {len(standardized_df)} 筆記錄")
    print(f"📅 日期範圍: {standardized_df['日期'].min()} ~ {standardized_df['日期'].max()}")
    
    return standardized_df

def merge_with_existing_data(new_df, existing_file="lottery_hist.xlsx"):
    """將新資料與現有資料合併"""
    
    existing_path = Path(existing_file)
    
    if existing_path.exists():
        try:
            existing_df = pd.read_excel(existing_file, engine='openpyxl')
            print(f"📁 現有記錄: {len(existing_df)} 筆")
            
            # 合併資料
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            
            # 過濾掉未來日期的記錄
            current_date = datetime.now().date()
            if '日期' in combined_df.columns:
                before_filter = len(combined_df)
                combined_df = combined_df[pd.to_datetime(combined_df['日期']).dt.date <= current_date]
                after_filter = len(combined_df)
                
                if before_filter > after_filter:
                    print(f"🗓️ 移除了 {before_filter - after_filter} 筆未來日期記錄")
            
            # 去重（基於日期）
            if '日期' in combined_df.columns:
                before_count = len(combined_df)
                combined_df = combined_df.drop_duplicates(subset=['日期'], keep='last')
                after_count = len(combined_df)
                
                if before_count > after_count:
                    print(f"🔄 移除了 {before_count - after_count} 筆重複記錄")
            
            # 重新排序
            if '日期' in combined_df.columns:
                combined_df = combined_df.sort_values('日期').reset_index(drop=True)
            
            final_df = combined_df
            
        except Exception as e:
            print(f"❌ 讀取現有檔案時發生錯誤: {e}")
            final_df = new_df
    else:
        print("📄 創建新的歷史記錄檔案")
        final_df = new_df
    
    return final_df

def main():
    """主要執行函數"""
    print("🎲 開始更新 539 開獎歷史記錄...")
    print("📦 使用 TaiwanLotteryCrawler 套件獲取完整記錄")
    
    # 獲取完整歷史記錄
    complete_data = get_complete_539_history()
    
    if complete_data is None:
        print("❌ 無法獲取歷史記錄")
        return False
    
    # 標準化資料格式
    standardized_data = standardize_lottery_data(complete_data)
    
    if standardized_data is None:
        print("❌ 資料標準化失敗")
        return False
    
    # 與現有資料合併
    final_data = merge_with_existing_data(standardized_data)
    
    # 儲存更新後的資料
    try:
        final_data.to_excel('lottery_hist.xlsx', index=False, engine='openpyxl')
        print(f"\n✅ 成功更新 lottery_hist.xlsx")
        print(f"📊 總記錄筆數: {len(final_data)}")
        
        if '日期' in final_data.columns:
            print(f"📅 日期範圍: {final_data['日期'].min()} ~ {final_data['日期'].max()}")
        
        # 顯示最新 5 筆記錄
        print("\n📋 最新 5 筆記錄:")
        print(final_data.tail())
        
        return True
        
    except Exception as e:
        print(f"❌ 儲存檔案時發生錯誤: {e}")
        return False

if __name__ == "__main__":
    success = main()
    if success:
        print("\n🎉 歷史記錄更新完成！")
    else:
        print("\n❌ 歷史記錄更新失敗")
        sys.exit(1)
