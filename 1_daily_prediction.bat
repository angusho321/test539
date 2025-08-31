@echo off
REM ========================================
REM 每日預測腳本 - 步驟1
REM ========================================
echo 🔮 每日539預測分析
echo =====================================
echo 執行時間: %date% %time%
echo.

echo 📊 正在執行預測分析...
python lottery_analysis.py

echo.
echo ✅ 每日預測完成！
echo.
echo 💡 下一步:
echo    等開獎後執行 "2_update_results.bat"
echo.
pause

