#!/usr/bin/env python3
"""
上傳開獎歷史資料到 Google Drive
晚上執行，更新爬取的最新開獎結果
"""

import os
import json
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaFileUpload
from datetime import datetime

def upload_lottery_hist_to_drive():
    """上傳 lottery_hist.xlsx 到 Google Drive"""
    
    # 設定 Google Drive API
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
    # 從環境變數讀取認證
    creds = Credentials.from_service_account_file(
        'credentials.json', scopes=SCOPES)
    
    service = build('drive', 'v3', credentials=creds)
    
    # 開獎歷史檔案 ID
    hist_file_id = os.environ.get('LOTTERY_HIST_FILE_ID')
    
    try:
        # 檢查檔案是否存在
        if not os.path.exists('lottery_hist.xlsx'):
            print("❌ lottery_hist.xlsx 不存在")
            print("📁 當前目錄檔案:")
            import glob
            files = glob.glob("*")
            for f in files:
                print(f"   - {f}")
            return False
        
        # 顯示檔案資訊
        file_size = os.path.getsize('lottery_hist.xlsx')
        print(f"📊 準備上傳開獎歷史檔案: lottery_hist.xlsx ({file_size} bytes)")

        # 安全防護：拒絕用殘缺小檔覆蓋雲端完整歷史
        MIN_ROWS = 100
        try:
            import pandas as pd
            row_count = len(pd.read_excel('lottery_hist.xlsx', engine='openpyxl'))
        except Exception as read_err:
            print(f"❌ 無法讀取 lottery_hist.xlsx 進行驗證，中止上傳: {read_err}")
            return False
        if row_count < MIN_ROWS:
            print(f"🛑 安全中止：lottery_hist.xlsx 僅 {row_count} 列（< {MIN_ROWS}），拒絕上傳以免覆蓋雲端完整歷史")
            return False
        print(f"🔍 上傳前驗證通過：{row_count} 列")
        
        if hist_file_id:
            try:
                # 更新現有檔案
                print(f"🔄 更新開獎歷史檔案 ID: {hist_file_id}")
                media = MediaFileUpload('lottery_hist.xlsx',
                                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                
                updated_file = service.files().update(fileId=hist_file_id,
                                                     media_body=media,
                                                     fields='id,name,webViewLink').execute()
                print(f"✅ 成功更新開獎歷史: {updated_file.get('name')}")
                print(f"📁 檔案 ID: {updated_file.get('id')}")
                print(f"🔗 檢視連結: {updated_file.get('webViewLink')}")
                return True
                
            except Exception as update_error:
                print(f"⚠️ 更新開獎歷史失敗: {str(update_error)}")
                return False
        else:
            print("❌ 未設定 LOTTERY_HIST_FILE_ID 環境變數")
            return False
        
    except Exception as e:
        print(f"❌ 上傳開獎歷史失敗: {str(e)}")
        return False

if __name__ == "__main__":
    import sys
    ok = upload_lottery_hist_to_drive()
    sys.exit(0 if ok else 1)
