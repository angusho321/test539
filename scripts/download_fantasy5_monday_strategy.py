#!/usr/bin/env python3
"""
從 Google Drive 下載天天樂週一策略檔案
"""

import os
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaIoBaseDownload
import io

def download_fantasy5_monday_strategy():
    """從 Google Drive 下載 fantasy5_monday_strategy.xlsx"""
    
    # 設定 Google Drive API
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
    # 從環境變數讀取認證
    creds = Credentials.from_service_account_file(
        'credentials.json', scopes=SCOPES)
    
    service = build('drive', 'v3', credentials=creds)
    
    # 檔案 ID
    FILE_ID = os.environ.get('FANTASY5_MONDAY_STRATEGY_FILE_ID')
    
    if not FILE_ID:
        print("⚠️ 未設定 FANTASY5_MONDAY_STRATEGY_FILE_ID 環境變數，跳過下載（將創建新檔案）")
        return True  # 如果沒有檔案 ID，允許繼續執行（會創建新檔案）
    
    try:
        # 下載檔案
        request = service.files().get_media(fileId=FILE_ID)
        
        with open('fantasy5_monday_strategy.xlsx', 'wb') as file:
            downloader = MediaIoBaseDownload(file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                print(f"⬇️ 下載進度: {int(status.progress() * 100)}%")
        
        print("✅ 成功下載 fantasy5_monday_strategy.xlsx")
        return True
        
    except Exception as e:
        print(f"⚠️ 下載失敗（將創建新檔案）: {str(e)}")
        return True  # 如果下載失敗，允許繼續執行（會創建新檔案）

if __name__ == "__main__":
    download_fantasy5_monday_strategy()
