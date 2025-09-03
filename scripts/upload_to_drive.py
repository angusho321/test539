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
    
    # 設定 Google Drive API - 使用更廣泛的權限範圍
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
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
        
        # 優先嘗試更新現有檔案
        original_file_id = os.environ.get('PREDICTION_LOG_FILE_ID')
        
        if original_file_id:
            try:
                # 首先驗證服務帳號是否能存取檔案
                print(f"🔍 驗證檔案存取權限 ID: {original_file_id}")
                
                # 嘗試取得檔案資訊
                file_info = service.files().get(fileId=original_file_id, 
                                               fields='id,name,parents,permissions,mimeType').execute()
                print(f"📄 檔案資訊: {file_info.get('name')}")
                print(f"📁 檔案 ID: {file_info.get('id')}")
                print(f"📂 父資料夾: {file_info.get('parents', ['根目錄'])}")
                print(f"📄 檔案類型: {file_info.get('mimeType')}")
                
                # 檢查權限
                try:
                    permissions = service.permissions().list(fileId=original_file_id).execute()
                    print(f"🔐 檔案權限數量: {len(permissions.get('permissions', []))}")
                    for perm in permissions.get('permissions', []):
                        if 'serviceAccount' in perm.get('type', ''):
                            print(f"   服務帳號權限: {perm.get('role', 'unknown')}")
                except Exception as perm_error:
                    print(f"⚠️ 無法讀取權限: {str(perm_error)}")
                
                # 根據檔案類型選擇上傳方式
                target_mime_type = file_info.get('mimeType', '')
                
                if 'spreadsheet' in target_mime_type or 'google-apps' in target_mime_type:
                    # 目標是 Google Sheets，上傳為 Google Sheets
                    print(f"📊 更新 Google Sheets...")
                    media = MediaFileUpload('prediction_log.xlsx',
                                          mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                else:
                    # 目標是一般 Excel 檔案
                    print(f"📁 更新 Excel 檔案...")
                    media = MediaFileUpload('prediction_log.xlsx',
                                          mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                
                updated_file = service.files().update(fileId=original_file_id,
                                                     media_body=media,
                                                     fields='id,name,webViewLink').execute()
                print(f"✅ 成功更新檔案: {updated_file.get('name')}")
                print(f"📁 檔案 ID: {updated_file.get('id')}")
                print(f"🔗 檢視連結: {updated_file.get('webViewLink')}")
                return True
                
            except Exception as update_error:
                print(f"⚠️ 檔案存取或更新失敗: {str(update_error)}")
                print("🔄 嘗試其他解決方案...")
        
        # 如果沒有指定檔案ID或更新失敗，嘗試建立新檔案
        try:
            print("📝 嘗試建立新檔案...")
            
            # 檔案 metadata
            file_metadata = {
                'name': filename
            }
            
            # 如果指定了資料夾，設定父資料夾
            if FOLDER_ID:
                file_metadata['parents'] = [FOLDER_ID]
                print(f"📁 目標資料夾 ID: {FOLDER_ID}")
            
            # 上傳檔案
            media = MediaFileUpload('prediction_log.xlsx',
                                  mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            
            file = service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id,name,webViewLink').execute()
            
            print(f"✅ 成功建立新檔案: {file.get('name')}")
            print(f"📁 檔案 ID: {file.get('id')}")
            print(f"🔗 檢視連結: {file.get('webViewLink')}")
            print(f"💡 建議將此檔案 ID 新增為 GitHub Secret: PREDICTION_LOG_FILE_ID")
            
        except Exception as create_error:
            print(f"❌ 建立新檔案也失敗: {str(create_error)}")
            print("💡 解決方案：")
            print("   1. 在 Google Drive 手動建立 prediction_log.xlsx")
            print("   2. 分享給服務帳號（編輯者權限）")
            print("   3. 將檔案 ID 新增為 GitHub Secret: PREDICTION_LOG_FILE_ID")
            print(f"   4. 服務帳號郵件: lottery-analysis-bot@test-539-470702.iam.gserviceaccount.com")
            raise create_error
        
        return True
        
    except Exception as e:
        print(f"❌ 上傳失敗: {str(e)}")
        return False

if __name__ == "__main__":
    upload_to_google_drive()
