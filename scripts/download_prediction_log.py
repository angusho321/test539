#!/usr/bin/env python3
"""
從 Google Drive 下載預測記錄
晚上執行，準備進行驗證
"""

import os
import json
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaIoBaseDownload
import io

def download_prediction_log():
    """從 Google Drive 下載 prediction_log.xlsx"""
    
    # 設定 Google Drive API
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
    # 從環境變數讀取認證
    creds = Credentials.from_service_account_file(
        'credentials.json', scopes=SCOPES)
    
    service = build('drive', 'v3', credentials=creds)
    
    # 檔案 ID
    FILE_ID = os.environ.get('PREDICTION_LOG_FILE_ID')
    
    if not FILE_ID:
        print("❌ 請設定 PREDICTION_LOG_FILE_ID 環境變數")
        return False
    
    try:
        # 下載檔案
        request = service.files().get_media(fileId=FILE_ID)
        
        with open('prediction_log.xlsx', 'wb') as file:
            downloader = MediaIoBaseDownload(file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                print(f"⬇️ 下載預測記錄進度: {int(status.progress() * 100)}%")
        
        print("✅ 成功下載 prediction_log.xlsx")
        return True
        
    except Exception as e:
        print(f"❌ 下載預測記錄失敗: {str(e)}")
        return False

if __name__ == "__main__":
    download_prediction_log()
