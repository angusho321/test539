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
    """從 Google Drive 下載 prediction_log.xlsx，並與本地檔案合併"""
    import pandas as pd
    from pathlib import Path
    
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
        # 檢查本地是否已有檔案
        local_file_exists = Path('prediction_log.xlsx').exists()
        
        if local_file_exists:
            print("📁 發現本地預測記錄檔案，準備合併...")
            # 備份本地檔案
            local_df = pd.read_excel('prediction_log.xlsx', engine='openpyxl')
            print(f"📊 本地記錄: {len(local_df)} 筆")
        
        # 檢查檔案類型並下載
        temp_filename = 'prediction_log_temp.xlsx'
        
        # 獲取檔案資訊
        file_metadata = service.files().get(fileId=FILE_ID).execute()
        mime_type = file_metadata.get('mimeType', '')
        
        print(f"📄 檔案類型: {mime_type}")
        
        if 'spreadsheet' in mime_type or 'google-apps' in mime_type:
            # Google Sheets 檔案，需要使用 export
            print("📊 檢測到 Google Sheets，使用 export 下載...")
            request = service.files().export_media(
                fileId=FILE_ID,
                mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
            # Google Sheets export 不支援 MediaIoBaseDownload，直接下載
            with open(temp_filename, 'wb') as file:
                file.write(request.execute())
            print("✅ Google Sheets 下載完成")
        else:
            # 一般檔案，使用 get_media
            print("📁 檢測到一般檔案，使用 get_media 下載...")
            request = service.files().get_media(fileId=FILE_ID)
            
            with open(temp_filename, 'wb') as file:
                downloader = MediaIoBaseDownload(file, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                    print(f"⬇️ 下載預測記錄進度: {int(status.progress() * 100)}%")
        
        # 讀取下載的檔案
        remote_df = pd.read_excel(temp_filename, engine='openpyxl')
        print(f"☁️ 雲端記錄: {len(remote_df)} 筆")
        
        # 如果本地有檔案，進行合併
        if local_file_exists:
            # 合併本地和雲端資料
            combined_df = pd.concat([remote_df, local_df], ignore_index=True)
            
            # 去重（只基於日期，保留同一天的不同時間記錄）
            if '日期' in combined_df.columns:
                before_count = len(combined_df)
                # 只有在日期和時間都完全相同時才去重
                if '時間' in combined_df.columns:
                    combined_df = combined_df.drop_duplicates(subset=['日期', '時間'], keep='last')
                else:
                    # 如果沒有時間欄位，只基於日期去重
                    combined_df = combined_df.drop_duplicates(subset=['日期'], keep='last')
                after_count = len(combined_df)
                
                if before_count > after_count:
                    print(f"🔄 移除了 {before_count - after_count} 筆完全重複的記錄")
                
                print(f"📊 合併後保留: {len(combined_df)} 筆記錄")
            
            # 按日期時間排序
            if '日期' in combined_df.columns:
                try:
                    combined_df['日期'] = pd.to_datetime(combined_df['日期'])
                    combined_df = combined_df.sort_values(['日期', '時間'])
                except:
                    print("⚠️ 無法按日期排序")
            
            final_df = combined_df
            print(f"🔗 合併後記錄: {len(final_df)} 筆")
        else:
            final_df = remote_df
            print("📥 使用雲端記錄")
        
        # 儲存最終檔案
        final_df.to_excel('prediction_log.xlsx', index=False, engine='openpyxl')
        
        # 清理暫存檔案
        Path(temp_filename).unlink()
        
        print("✅ 成功下載並合併預測記錄")
        return True
        
    except Exception as e:
        print(f"❌ 下載預測記錄失敗: {str(e)}")
        return False

if __name__ == "__main__":
    download_prediction_log()
