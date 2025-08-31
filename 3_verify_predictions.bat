@echo off
REM ========================================
REM 驗證預測結果腳本 - 步驟3
REM ========================================
echo 🎯 驗證539預測結果
echo =====================================
echo 執行時間: %date% %time%
echo.

echo 🔍 正在驗證預測準確度...
python verify_predictions.py

echo.
echo ✅ 驗證完成！
echo 💡 可以查看 prediction_log.xlsx 查看詳細結果
echo.
pause

