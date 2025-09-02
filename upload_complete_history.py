#!/usr/bin/env python3
"""
一次性上傳完整的歷史記錄到 Google Drive
這個腳本只需要執行一次，將我們的 1882 筆完整記錄上傳到雲端
"""

import os
import pandas as pd
from pathlib import Path

def check_local_history():
    """檢查本地歷史記錄"""
    if not Path('lottery_hist.xlsx').exists():
        print("❌ 本地沒有 lottery_hist.xlsx 檔案")
        return False
    
    df = pd.read_excel('lottery_hist.xlsx', engine='openpyxl')
    print(f"📊 本地歷史記錄: {len(df)} 筆")
    print(f"📅 日期範圍: {df['日期'].min()} ~ {df['日期'].max()}")
    
    if len(df) < 1800:
        print("⚠️ 記錄數量可能不完整，建議先執行 update_lottery_history_with_taiwanlottery.py")
        return False
    
    return True

def main():
    """主要執行函數"""
    print("🚀 準備上傳完整歷史記錄到 Google Drive")
    print("="*50)
    
    # 檢查本地檔案
    if not check_local_history():
        print("\n❌ 本地歷史記錄檢查失敗")
        print("💡 請先執行: python scripts/update_lottery_history_with_taiwanlottery.py")
        return False
    
    # 上傳到 Google Drive
    print("\n⬆️ 開始上傳到 Google Drive...")
    
    try:
        # 使用現有的上傳腳本
        from scripts.upload_lottery_hist import upload_lottery_hist_to_drive
        
        # 設定環境變數（如果未設定）
        if not os.environ.get('GOOGLE_CREDENTIALS'):
            print("❌ 請設定 GOOGLE_CREDENTIALS 環境變數")
            print("💡 在 GitHub Actions 中這會自動設定")
            return False
        
        if not os.environ.get('LOTTERY_HIST_FILE_ID'):
            print("❌ 請設定 LOTTERY_HIST_FILE_ID 環境變數")
            print("💡 這是 Google Drive 中目標檔案的 ID")
            return False
        
        success = upload_lottery_hist_to_drive()
        
        if success:
            print("\n🎉 完整歷史記錄上傳成功！")
            print("💡 現在 GitHub Actions 會下載完整的歷史記錄")
            return True
        else:
            print("\n❌ 上傳失敗")
            return False
            
    except Exception as e:
        print(f"❌ 上傳過程發生錯誤: {e}")
        return False

if __name__ == "__main__":
    success = main()
    if not success:
        print("\n💡 手動上傳步驟:")
        print("1. 將本地的 lottery_hist.xlsx 手動上傳到 Google Drive")
        print("2. 取代原有的歷史檔案")
        print("3. 確保 GitHub Secret LOTTERY_HIST_FILE_ID 指向正確的檔案")
