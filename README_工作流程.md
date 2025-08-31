# 539彩票分析系統 - 每日工作流程

## 📋 建議的每日操作流程

### 步驟1: 每日預測 (開獎前)
```bash
# 方法1: 使用批次腳本
雙擊執行 "1_daily_prediction.bat"

# 方法2: 直接執行
python lottery_analysis.py
```
- **何時執行**: 每天一次，建議早上執行
- **功能**: 根據歷史資料生成當日預測號碼
- **輸出**: 更新 `prediction_log.xlsx`
- **注意**: ⚠️ 每天只執行一次，避免覆蓋當日預測

### 步驟2: 更新開獎資料 (開獎後)
```bash
# 方法1: 使用批次腳本
雙擊執行 "2_update_results.bat"

# 方法2: 直接執行
python auto_lottery_system.py --update-only
```
- **何時執行**: 每日開獎後 (晚上約8:30後)
- **功能**: 從 pilio.idv.tw 獲取最新開獎資料
- **輸出**: 更新 `lottery_hist.xlsx`
- **安全**: 不會影響當日已生成的預測記錄

### 步驟3: 驗證預測結果
```bash
# 方法1: 使用批次腳本
雙擊執行 "3_verify_predictions.bat"

# 方法2: 直接執行
python verify_predictions.py
```
- **何時執行**: 更新開獎資料後立即執行
- **功能**: 比對預測號碼與實際開獎結果
- **輸出**: 更新 `prediction_log.xlsx` 的驗證欄位

## 🔧 其他操作模式

### 查看使用說明
```bash
python auto_lottery_system.py --help
```

### 僅執行預測分析 (如果誤刪當日預測)
```bash
python auto_lottery_system.py --predict-only
```

### 傳統完整模式 (不建議日常使用)
```bash
python auto_lottery_system.py
```

## 📊 檔案說明

| 檔案 | 用途 | 更新時機 |
|------|------|----------|
| `lottery_hist.xlsx` | 歷史開獎資料 | 每次執行步驟2 |
| `prediction_log.xlsx` | 預測記錄與驗證結果 | 步驟1和步驟3 |
| `auto_system_log.txt` | 系統執行記錄 | 自動記錄 |

## ⚠️ 重要注意事項

1. **不要重複執行步驟1** - 會覆蓋當日預測記錄
2. **按順序執行** - 步驟2需要步驟1的預測記錄存在
3. **檢查網路連線** - 步驟2需要網路存取 pilio.idv.tw
4. **定期備份** - 建議定期備份Excel檔案

## 🎯 預期結果

正常執行後，您應該看到：
- 步驟1: 生成當日各策略的預測號碼
- 步驟2: 獲取最新開獎資料
- 步驟3: 顯示各策略的預測準確度統計

這樣的流程確保每日只生成一次預測，避免重複覆蓋，並能準確追蹤預測表現！

