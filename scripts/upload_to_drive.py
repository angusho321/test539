#!/usr/bin/env python3
"""
上傳分析結果到 Google Drive
"""

import os
import json
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaFileUpload
from datetime import datetime

def upload_to_google_drive():
    """上傳 prediction_log.xlsx 到 Google Drive"""
    
    # 設定 Google Drive API
    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    
    # 從環境變數讀取認證
    creds = Credentials.from_service_account_file(
        'credentials.json', scopes=SCOPES)
    
    service = build('drive', 'v3', credentials=creds)
    
    # 目標資料夾 ID (選填，如果要上傳到特定資料夾)
    FOLDER_ID = os.environ.get('GOOGLE_DRIVE_FOLDER_ID')
    
    try:
        # 檢查檔案是否存在
        if not os.path.exists('prediction_log.xlsx'):
            print("❌ prediction_log.xlsx 不存在")
            print("📁 當前目錄檔案:")
            import glob
            files = glob.glob("*")
            for f in files:
                print(f"   - {f}")
            return False
        
        # 顯示檔案資訊
        file_size = os.path.getsize('prediction_log.xlsx')
        print(f"📊 準備上傳檔案: prediction_log.xlsx ({file_size} bytes)")
        
        # 設定檔案名稱 (包含日期)
        today = datetime.now().strftime("%Y-%m-%d")
        filename = f"prediction_log_{today}.xlsx"
        
        # 檔案 metadata
        file_metadata = {
            'name': filename
        }
        
        # 如果指定了資料夾，設定父資料夾
        if FOLDER_ID:
            file_metadata['parents'] = [FOLDER_ID]
        
        # 上傳檔案
        media = MediaFileUpload('prediction_log.xlsx',
                              mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        file = service.files().create(body=file_metadata,
                                    media_body=media,
                                    fields='id,name,webViewLink').execute()
        
        print(f"✅ 成功上傳: {file.get('name')}")
        print(f"📁 檔案 ID: {file.get('id')}")
        print(f"🔗 檢視連結: {file.get('webViewLink')}")
        
        # 同時更新原始檔案 (覆蓋模式)
        original_file_id = os.environ.get('PREDICTION_LOG_FILE_ID')
        if original_file_id:
            media = MediaFileUpload('prediction_log.xlsx',
                                  mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            
            service.files().update(fileId=original_file_id,
                                 media_body=media).execute()
            print("✅ 同時更新了原始預測記錄檔案")
        
        return True
        
    except Exception as e:
        print(f"❌ 上傳失敗: {str(e)}")
        return False

if __name__ == "__main__":
    upload_to_google_drive()
