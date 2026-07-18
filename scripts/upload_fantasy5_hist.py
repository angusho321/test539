#!/usr/bin/env python3
"""
上傳加州Fantasy 5歷史資料到 Google Drive
"""

import os
import json
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaFileUpload

def upload_fantasy5_hist_to_drive():
    """上傳 fantasy5_hist.xlsx 到 Google Drive"""
    
    # 設定 Google Drive API
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
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
        # 檢查檔案是否存在
        if not os.path.exists('fantasy5_hist.xlsx'):
            print("❌ fantasy5_hist.xlsx 不存在")
            return False

        # 安全防護：拒絕用殘缺小檔覆蓋雲端完整歷史
        MIN_ROWS = 100
        try:
            import pandas as pd
            row_count = len(pd.read_excel('fantasy5_hist.xlsx', engine='openpyxl'))
        except Exception as read_err:
            print(f"❌ 無法讀取 fantasy5_hist.xlsx 進行驗證，中止上傳: {read_err}")
            return False
        if row_count < MIN_ROWS:
            print(f"🛑 安全中止：fantasy5_hist.xlsx 僅 {row_count} 列（< {MIN_ROWS}），拒絕上傳以免覆蓋雲端完整歷史")
            return False
        print(f"🔍 上傳前驗證通過：{row_count} 列")

        # 上傳檔案
        media = MediaFileUpload('fantasy5_hist.xlsx',
                              mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        updated_file = service.files().update(fileId=file_id,
                                            media_body=media,
                                            fields='id,name,webViewLink').execute()
        
        print(f"✅ 成功更新加州Fantasy 5歷史檔案: {updated_file.get('name')}")
        print(f"🔗 檢視連結: {updated_file.get('webViewLink')}")
        return True
        
    except Exception as e:
        print(f"❌ 上傳失敗: {str(e)}")
        return False

if __name__ == "__main__":
    import sys
    ok = upload_fantasy5_hist_to_drive()
    sys.exit(0 if ok else 1)
