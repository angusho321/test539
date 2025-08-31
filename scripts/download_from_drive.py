#!/usr/bin/env python3
"""
從 Google Drive 下載彩票歷史資料
"""

import os
import json
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaIoBaseDownload
import io

def download_from_google_drive():
    """從 Google Drive 下載 lottery_hist.xlsx"""
    
    # 設定 Google Drive API
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
    # 從環境變數讀取認證
    creds = Credentials.from_service_account_file(
        'credentials.json', scopes=SCOPES)
    
    service = build('drive', 'v3', credentials=creds)
    
    # 檔案 ID (需要從 Google Drive 分享連結取得)
    # 格式: https://drive.google.com/file/d/FILE_ID/view
    FILE_ID = os.environ.get('LOTTERY_HIST_FILE_ID')
    
    if not FILE_ID:
        print("❌ 請設定 LOTTERY_HIST_FILE_ID 環境變數")
        return False
    
    try:
        # 下載檔案
        request = service.files().get_media(fileId=FILE_ID)
        
        with open('lottery_hist.xlsx', 'wb') as file:
            downloader = MediaIoBaseDownload(file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                print(f"⬇️ 下載進度: {int(status.progress() * 100)}%")
        
        print("✅ 成功下載 lottery_hist.xlsx")
        return True
        
    except Exception as e:
        print(f"❌ 下載失敗: {str(e)}")
        return False

if __name__ == "__main__":
    download_from_google_drive()
