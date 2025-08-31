@echo off
REM ========================================
REM 更新開獎結果腳本 - 步驟2
REM ========================================
echo 📥 更新539開獎結果
echo =====================================
echo 執行時間: %date% %time%
echo.

echo 🕷️ 正在從pilio網站獲取最新開獎資料...
python auto_lottery_system.py --update-only

echo.
echo 💡 下一步:
echo    執行 "3_verify_predictions.bat" 來驗證預測結果
echo.
pause

