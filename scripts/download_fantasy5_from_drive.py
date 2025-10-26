#!/usr/bin/env python3
"""
從 Google Drive 下載加州Fantasy 5歷史資料
"""

import os
import json
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaIoBaseDownload
from io import BytesIO

def download_fantasy5_from_drive():
    """從 Google Drive 下載 fantasy5_hist.xlsx"""
    
    # 設定 Google Drive API
    SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
    
    # 從環境變數讀取認證
    creds = Credentials.from_service_account_file(
        'credentials.json', scopes=SCOPES)
    
    service = build('drive', 'v3', credentials=creds)
    
    # 從環境變數取得檔案 ID
    file_id = os.environ.get('FANTASY5_HIST_FILE_ID')
    
    if not file_id:
        print("❌ 未設定 FANTASY5_HIST_FILE_ID 環境變數")
        return False
    
    try:
        # 下載檔案
        request = service.files().get_media(fileId=file_id)
        file_content = BytesIO()
        downloader = MediaIoBaseDownload(file_content, request)
        
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"📥 下載進度: {int(status.progress() * 100)}%")
        
        # 保存到本地
        with open('fantasy5_hist.xlsx', 'wb') as f:
            f.write(file_content.getvalue())
        
        print("✅ 成功下載 fantasy5_hist.xlsx")
        return True
        
    except Exception as e:
        print(f"❌ 下載失敗: {str(e)}")
        return False

if __name__ == "__main__":
    download_fantasy5_from_drive()
