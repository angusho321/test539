# 🎲 539彩票分析系統 - GitHub Actions 自動化部署

## 📋 概述

這個專案使用 GitHub Actions 實現彩票分析的自動化執行，每日定時分析並將結果上傳至 Google Drive。

## 🚀 設定步驟

### 1. 準備 Google Drive API 認證

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立新專案或選擇現有專案
3. 啟用 Google Drive API
4. 建立服務帳號 (Service Account)
5. 下載服務帳號的 JSON 金鑰檔案

### 2. 設定 GitHub Secrets

在您的 GitHub 倉庫中設定以下 Secrets：

- `GOOGLE_CREDENTIALS`: 貼上整個 JSON 金鑰檔案內容
- `LOTTERY_HIST_FILE_ID`: Google Drive 中 lottery_hist.xlsx 的檔案 ID
- `PREDICTION_LOG_FILE_ID`: (選填) 預測記錄檔案的 ID (用於更新)
- `GOOGLE_DRIVE_FOLDER_ID`: (選填) 目標資料夾 ID

#### 如何取得檔案 ID：
從 Google Drive 分享連結中取得：
```
https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74mHxYjY/view
                            ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                            這部分就是檔案 ID
```

### 3. 上傳檔案

1. 將您的 `lottery_hist.xlsx` 上傳到 Google Drive
2. 確保服務帳號有讀取權限
3. 將本專案的所有檔案上傳到您的 GitHub 倉庫

### 4. 檔案結構

```
你的倉庫/
├── .github/
│   └── workflows/
│       └── lottery-analysis.yml    # GitHub Actions 工作流程
├── scripts/
│   ├── download_from_drive.py      # 從 Google Drive 下載檔案
│   └── upload_to_drive.py          # 上傳結果到 Google Drive
├── lottery_analysis_cloud.py       # 主要分析程式 (雲端版)
├── requirements.txt                # Python 依賴套件
└── README_GitHub_Actions.md        # 本說明檔案
```

## ⏰ 執行時間

- **自動執行**: 每天 UTC 01:00 (台灣時間 09:00)
- **手動執行**: 可在 GitHub Actions 頁面手動觸發

## 📊 輸出結果

- 分析結果會上傳到 Google Drive
- 同時在 GitHub Actions 中保存為 artifact (保留 30 天)
- 日誌可在 GitHub Actions 執行歷史中查看

## 🔧 自訂設定

### 修改執行時間

編輯 `.github/workflows/lottery-analysis.yml` 中的 cron 表達式：

```yaml
schedule:
  - cron: '0 1 * * *'  # UTC 01:00 (台灣時間 09:00)
```

### 修改分析邏輯

編輯 `lottery_analysis_cloud.py` 檔案，加入您的自訂分析邏輯。

## 🛠️ 故障排除

### 常見問題

1. **權限錯誤**: 確保服務帳號有 Google Drive 檔案的讀寫權限
2. **檔案 ID 錯誤**: 檢查 Secrets 中的檔案 ID 是否正確
3. **執行失敗**: 查看 GitHub Actions 日誌了解詳細錯誤訊息

### 檢查執行狀態

1. 前往您的 GitHub 倉庫
2. 點選 "Actions" 頁籤
3. 查看最近的執行記錄和日誌

## 💰 成本說明

- **GitHub Actions**: 公開倉庫完全免費
- **Google Drive API**: 免費額度內使用
- **總成本**: $0 💰

## 📞 支援

如有問題，請檢查：
1. GitHub Actions 執行日誌
2. Google Cloud Console 中的 API 配額
3. 檔案權限設定
