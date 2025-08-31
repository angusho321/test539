# 🚀 GitHub Actions 自動化彩票分析設定指南

## 📋 前置準備檢查清單

### ✅ 已完成
- [x] GitHub 倉庫建立
- [x] Google Cloud 服務帳號建立
- [x] GitHub Secrets 設定
- [x] 程式碼已推送

### ⚠️ 需要手動完成的步驟

#### 1. 建立 Google Drive 預測記錄檔案

**重要：** 由於服務帳號沒有儲存配額，需要手動建立檔案讓服務帳號更新：

1. **在您的 Google Drive 建立空白 Excel 檔案**
   - 檔案名稱：`prediction_log.xlsx`
   - 內容：空白即可

2. **分享檔案給服務帳號**
   - 右鍵點選檔案 → 「共用」
   - 新增：`lottery-analysis-bot@test-539-470702.iam.gserviceaccount.com`
   - 權限：**編輯者**

3. **取得檔案 ID**
   - 點選檔案的「取得連結」
   - 從 URL 複製檔案 ID：
   ```
   https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74mHxYjY/view
                               ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                               這就是檔案 ID
   ```

4. **新增 GitHub Secret**
   - 前往：https://github.com/angusho321/test539/settings/secrets/actions
   - 新增 Secret：
     ```
     Name: PREDICTION_LOG_FILE_ID
     Value: [您的檔案 ID]
     ```

#### 2. 確認現有設定

確認以下 Secrets 已正確設定：
- ✅ `GOOGLE_CREDENTIALS` (JSON 金鑰)
- ✅ `LOTTERY_HIST_FILE_ID` (歷史資料檔案 ID)
- ⚠️ `PREDICTION_LOG_FILE_ID` (預測記錄檔案 ID) - **需要新增**

## 🧪 測試步驟

1. **手動觸發測試**
   - 前往：https://github.com/angusho321/test539/actions
   - 選擇 "📊 Daily Lottery Analysis"
   - 點選 "Run workflow"

2. **檢查執行結果**
   - 查看每個步驟的日誌
   - 確認看到：`✅ 成功更新檔案`
   - 檢查 Google Drive 中的檔案是否有更新

## ⏰ 排程設定

- **自動執行時間**：每天台灣時間上午 9:00
- **Cron 表達式**：`0 1 * * *` (UTC 01:00)

## 🔧 故障排除

### 如果上傳失敗
- 確認 `PREDICTION_LOG_FILE_ID` 設定正確
- 確認檔案已分享給服務帳號
- 確認服務帳號有編輯權限

### 如果分析失敗
- 確認 `LOTTERY_HIST_FILE_ID` 正確
- 確認歷史資料檔案可存取

## 📱 監控建議

建議設定 GitHub Actions 執行結果通知：
1. GitHub 帳號設定 → Notifications
2. 啟用 Actions 失敗通知

## 🎯 完成後

一旦設定完成，系統將：
- ✅ 每天自動下載最新彩票資料
- ✅ 執行多種策略分析
- ✅ 更新 Google Drive 中的預測記錄
- ✅ 保留 30 天的執行記錄

**出國期間完全自動化運行！** 🌍✈️
