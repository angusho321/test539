#!/usr/bin/env python3
"""
上傳 539 週一策略檔案到 Google Drive
"""

import os
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaFileUpload

def upload_539_monday_strategy():
    """上傳 539_monday_strategy.xlsx 到 Google Drive"""
    
    # 設定 Google Drive API
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
    # 從環境變數讀取認證
    creds = Credentials.from_service_account_file(
        'credentials.json', scopes=SCOPES)
    
    service = build('drive', 'v3', credentials=creds)
    
    # 檔案 ID
    file_id = os.environ.get('MONDAY_STRATEGY_539_FILE_ID')
    
    try:
        # 檢查檔案是否存在
        if not os.path.exists('539_monday_strategy.xlsx'):
            print("❌ 539_monday_strategy.xlsx 不存在")
            return False
        
        # 顯示檔案資訊
        file_size = os.path.getsize('539_monday_strategy.xlsx')
        print(f"📊 準備上傳 539 週一策略檔案: 539_monday_strategy.xlsx ({file_size} bytes)")
        
        if file_id:
            try:
                # 更新現有檔案
                print(f"🔄 更新 539 週一策略檔案 ID: {file_id}")
                media = MediaFileUpload('539_monday_strategy.xlsx',
                                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                
                updated_file = service.files().update(fileId=file_id,
                                                     media_body=media,
                                                     fields='id,name,webViewLink').execute()
                print(f"✅ 成功更新 539 週一策略: {updated_file.get('name')}")
                print(f"📁 檔案 ID: {updated_file.get('id')}")
                print(f"🔗 檢視連結: {updated_file.get('webViewLink')}")
                return True
                
            except Exception as update_error:
                print(f"⚠️ 更新 539 週一策略失敗: {str(update_error)}")
                print("💡 如果檔案不存在，請在 Google Drive 手動建立並分享給服務帳號")
                return False
        else:
            print("❌ 未設定 MONDAY_STRATEGY_539_FILE_ID 環境變數")
            print("💡 請在 Google Drive 手動建立 539_monday_strategy.xlsx 並分享給服務帳號")
            print("   然後將檔案 ID 新增為 GitHub Secret: MONDAY_STRATEGY_539_FILE_ID")
            return False
        
    except Exception as e:
        print(f"❌ 上傳 539 週一策略失敗: {str(e)}")
        return False

if __name__ == "__main__":
    upload_539_monday_strategy()
