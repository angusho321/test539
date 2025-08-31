#!/usr/bin/env python3
"""
彩票分析程式 - 雲端版本
適用於 GitHub Actions 等雲端環境
"""

import pandas as pd
import numpy as np
import random
from pathlib import Path
from datetime import datetime
import os
import logging

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 從原本的 lottery_analysis.py 複製核心函數
# (這裡只包含主要函數，完整版請從原檔案複製)

def load_lottery_excel(excel_path: str):
    """讀入 .xlsx 開獎紀錄"""
    return pd.read_excel(excel_path, engine='openpyxl')

def compute_num_frequency(df: pd.DataFrame):
    """回傳每個號碼的頻次"""
    nums = df[['號碼1','號碼2','號碼3','號碼4','號碼5']].values.ravel()
    freq = pd.Series(nums).value_counts().sort_index()
    freq.index.name = '號碼'
    freq.name = '頻次'
    return freq

def get_hot_cold_numbers(freq: pd.Series, top_n=6, bottom_n=6):
    """取得熱號冷號"""
    hot_numbers = freq.nlargest(top_n).index.tolist()
    cold_numbers = freq.nsmallest(bottom_n).index.tolist()
    return hot_numbers, cold_numbers

def suggest_numbers(strategy='smart', n=9, historical_stats=None):
    """產生建議號碼 (簡化版本)"""
    numbers = list(range(1, 40))
    global hot_numbers, cold_numbers

    if strategy == 'random':
        return sorted(random.sample(numbers, n))
    elif strategy == 'hot':
        sel = hot_numbers.copy()
        if len(sel) >= n:
            return sorted(random.sample(sel, n))
        else:
            remain = [x for x in numbers if x not in sel]
            sel += random.sample(remain, n - len(sel))
            return sorted(sel)
    elif strategy == 'cold':
        sel = cold_numbers.copy()
        if len(sel) >= n:
            return sorted(random.sample(sel, n))
        else:
            remain = [x for x in numbers if x not in sel]
            sel += random.sample(remain, n - len(sel))
            return sorted(sel)
    else:  # smart or balanced
        return sorted(random.sample(numbers, n))

def check_daily_analysis_exists(log_file="prediction_log.xlsx"):
    """檢查今日是否已有分析記錄 (雲端版本)"""
    if not Path(log_file).exists():
        return False
    
    try:
        current_time = datetime.now()
        date_str = current_time.strftime("%Y-%m-%d")
        
        existing_df = pd.read_excel(log_file, engine='openpyxl')
        today_records = existing_df[existing_df['日期'] == date_str]
        
        if len(today_records) > 0:
            logger.info(f"📅 檢測到今日({date_str})已有分析記錄")
            return True
        return False
    except Exception as e:
        logger.warning(f"⚠️ 檢查今日記錄時發生錯誤: {e}")
        return False

def log_predictions_to_excel(predictions, log_file="prediction_log.xlsx"):
    """記錄預測結果 (雲端版本)"""
    current_time = datetime.now()
    date_str = current_time.strftime("%Y-%m-%d")
    time_str = current_time.strftime("%H:%M:%S")
    
    # 準備記錄數據
    log_data = {
        '日期': date_str,
        '時間': time_str,
        '智能選號': str(predictions.get('smart', [])),
        '平衡策略': str(predictions.get('balanced', [])),
        '隨機選號': str(predictions.get('random', [])),
        '熱號優先': str(predictions.get('hot', [])),
        '冷號優先': str(predictions.get('cold', [])),
        '智能選號_前5號': str(predictions.get('smart', [])[:5]) if predictions.get('smart') else '',
        '平衡策略_前5號': str(predictions.get('balanced', [])[:5]) if predictions.get('balanced') else '',
        '驗證結果': '',
        '中獎號碼數': '',
        '備註': f"雲端自動分析 - {os.environ.get('GITHUB_WORKFLOW', 'Unknown')}"
    }
    
    # 檢查檔案是否存在
    log_path = Path(log_file)
    
    if log_path.exists():
        try:
            existing_df = pd.read_excel(log_file, engine='openpyxl')
            new_df = pd.DataFrame([log_data])
            
            # 檢查是否有新欄位需要添加到現有數據
            for col in new_df.columns:
                if col not in existing_df.columns:
                    existing_df[col] = ''
            
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        except Exception as e:
            logger.error(f"讀取現有記錄檔案時發生錯誤: {e}")
            combined_df = pd.DataFrame([log_data])
    else:
        combined_df = pd.DataFrame([log_data])
    
    # 寫入Excel
    try:
        combined_df.to_excel(log_file, index=False, engine='openpyxl')
        logger.info(f"✅ 預測記錄已保存到: {log_file}")
        logger.info(f"   記錄時間: {date_str} {time_str}")
        return True
    except Exception as e:
        logger.error(f"❌ 保存預測記錄時發生錯誤: {e}")
        return False

def main():
    """主程式 - 雲端版本"""
    logger.info("🎲 539彩票分析系統 - 雲端版本")
    logger.info("="*60)
    
    # 檢查輸入檔案
    excel_file = "lottery_hist.xlsx"
    if not Path(excel_file).exists():
        logger.error(f"❌ 找不到檔案: {excel_file}")
        return False
    
    # 檢查今日是否已有分析記錄
    if check_daily_analysis_exists("prediction_log.xlsx"):
        logger.info("✅ 今日分析已完成，跳過重複分析")
        # 但仍然嘗試上傳現有的檔案
        if os.path.exists("prediction_log.xlsx"):
            logger.info("📤 嘗試上傳現有的預測記錄檔案")
            return True
        else:
            logger.warning("⚠️ 預測記錄檔案不存在，將重新生成")
            # 繼續執行分析
    
    try:
        # 載入資料
        df = load_lottery_excel(excel_file)
        logger.info(f"📊 成功載入 {len(df)} 筆歷史資料")
        
        # 基本統計
        freq_series = compute_num_frequency(df)
        global hot_numbers, cold_numbers
        hot_numbers, cold_numbers = get_hot_cold_numbers(freq_series, top_n=6, bottom_n=6)
        
        logger.info(f"🔥 熱號: {hot_numbers}")
        logger.info(f"❄️ 冷號: {cold_numbers}")
        
        # 生成各策略的建議號碼
        strategies = ['smart', 'balanced', 'random', 'hot', 'cold']
        predictions = {}
        
        for strategy in strategies:
            numbers = suggest_numbers(strategy, n=9)
            predictions[strategy] = numbers
            logger.info(f"📋 {strategy.upper()}: {numbers}")
        
        # 記錄預測結果
        success = log_predictions_to_excel(predictions, "prediction_log.xlsx")
        
        if success:
            logger.info("🎉 分析完成並成功記錄!")
        else:
            logger.error("❌ 記錄預測結果時發生錯誤")
            return False
            
        return True
        
    except Exception as e:
        logger.error(f"❌ 分析過程發生錯誤: {str(e)}")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
