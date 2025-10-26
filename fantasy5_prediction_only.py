#!/usr/bin/env python3
"""
加州Fantasy 5預測程式 - 僅預測版本
專門用於加州Fantasy 5，只生成預測不驗證
"""

import pandas as pd
import numpy as np
import random
from pathlib import Path
from datetime import datetime
import os
import logging
from collections import defaultdict
from itertools import combinations

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 從原本的 lottery_analysis.py 複製核心函數
def load_lottery_excel(excel_path: str):
    """讀入 .xlsx 開獎紀錄"""
    return pd.read_excel(excel_path, engine='openpyxl')



def compute_weighted_frequency(df, decay_factor=0.95, recent_days=365):
    """
    計算時間加權的號碼頻率
    越近期的記錄權重越高，避免資料鈍化問題
    """
    try:
        # 只取最近的記錄
        cutoff_date = datetime.now() - pd.Timedelta(days=recent_days)
        recent_df = df[df['日期'] >= cutoff_date].copy()
        
        if len(recent_df) == 0:
            logger.warning("⚠️ 沒有足夠的近期記錄，使用全部資料")
            recent_df = df.copy()
        
        logger.info(f"⚖️ 使用最近 {len(recent_df)} 筆記錄進行加權分析")
        
        # 計算每筆記錄距今的天數
        today = datetime.now()
        recent_df['days_ago'] = (today - recent_df['日期']).dt.days
        
        # 計算加權頻率
        weighted_freq = defaultdict(float)
        total_weight = 0
        
        for _, row in recent_df.iterrows():
            # 計算權重：越近期權重越高
            weight = decay_factor ** row['days_ago']
            total_weight += weight
            
            # 累加每個號碼的加權頻率
            for col in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']:
                if pd.notna(row[col]):
                    number = int(row[col])
                    weighted_freq[number] += weight
        
        # 正規化頻率
        if total_weight > 0:
            for num in weighted_freq:
                weighted_freq[num] = weighted_freq[num] / total_weight
        
        logger.info(f"✅ 完成時間加權分析，衰減係數: {decay_factor}")
        
        return dict(weighted_freq)
        
    except Exception as e:
        logger.error(f"❌ 時間加權計算失敗: {e}")
        return {}

def compute_weighted_frequency_150(df, decay_factor=0.95, recent_periods=150):
    """
    計算150筆的時間加權號碼頻率 (A/B測試版本)
    """
    try:
        # 只取最近150筆記錄
        recent_df = df.tail(recent_periods).copy()
        
        if len(recent_df) == 0:
            logger.warning("⚠️ 沒有足夠的記錄，使用全部資料")
            recent_df = df.copy()
        
        logger.info(f"⚖️ 使用最近 {len(recent_df)} 筆記錄進行150筆加權分析")
        
        # 計算每筆記錄距今的天數
        today = datetime.now()
        recent_df['days_ago'] = (today - recent_df['日期']).dt.days
        
        # 計算加權頻率
        weighted_freq = defaultdict(float)
        total_weight = 0
        
        for _, row in recent_df.iterrows():
            # 計算權重：越近期權重越高
            weight = decay_factor ** row['days_ago']
            total_weight += weight
            
            # 累加每個號碼的加權頻率
            for col in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']:
                if pd.notna(row[col]):
                    number = int(row[col])
                    weighted_freq[number] += weight
        
        # 正規化頻率
        if total_weight > 0:
            for num in weighted_freq:
                weighted_freq[num] = weighted_freq[num] / total_weight
        
        logger.info(f"✅ 完成150筆時間加權分析，衰減係數: {decay_factor}")
        
        return dict(weighted_freq)
        
    except Exception as e:
        logger.error(f"❌ 150筆時間加權計算失敗: {e}")
        return {}

def suggest_numbers(strategy='smart', n=9, historical_stats=None, df=None, randomness_factor=0.3):
    """產生建議號碼 (加州Fantasy 5版本)"""
    numbers = list(range(1, 40))  # Fantasy 5也是1-39

    if strategy == 'smart':
        # 智能選號：時間加權智能選號
        if df is not None:
            try:
                weighted_freq = compute_weighted_frequency(df)
                if weighted_freq:
                    # 根據加權頻率選號
                    weights = [weighted_freq.get(num, 0.001) for num in numbers]
                    
                    # 加入隨機性避免過度依賴歷史（可調整的隨機因子）
                    for i in range(len(weights)):
                        weights[i] = weights[i] * (1 - randomness_factor) + random.random() * randomness_factor
                    
                    # 正規化權重
                    weights = np.array(weights)
                    weights = weights / weights.sum()
                    
                    # 根據權重選號
                    selected = np.random.choice(numbers, size=n, replace=False, p=weights)
                    result = sorted([int(x) for x in selected.tolist()])  # 確保都是 int
                    logger.info(f"🧠 時間加權智能選號 (隨機因子:{randomness_factor}): {result}")
                    return result
            except Exception as e:
                logger.warning(f"⚠️ 時間加權選號失敗，使用隨機選號: {e}")
        
        return sorted(random.sample(numbers, n))
    elif strategy == 'balanced':
        # 平衡策略：純隨機選號（受益於智能選號的隨機狀態污染）
        result = sorted(random.sample(numbers, n))
        logger.info(f"⚖️ 平衡策略 (隨機因子:{randomness_factor}): {result}")
        return result
    else:
        # 其他策略暫不使用
        return sorted(random.sample(numbers, n))

def suggest_numbers_150(strategy='smart', n=9, historical_stats=None, df=None, randomness_factor=0.3):
    """產生建議號碼 (150筆A/B測試版本)"""
    numbers = list(range(1, 40))  # Fantasy 5也是1-39

    if strategy == 'smart':
        # 智能選號：150筆時間加權智能選號
        if df is not None:
            try:
                weighted_freq = compute_weighted_frequency_150(df)
                if weighted_freq:
                    # 根據加權頻率選號
                    weights = [weighted_freq.get(num, 0.001) for num in numbers]
                    
                    # 加入隨機性避免過度依賴歷史（可調整的隨機因子）
                    for i in range(len(weights)):
                        weights[i] = weights[i] * (1 - randomness_factor) + random.random() * randomness_factor
                    
                    # 正規化權重
                    weights = np.array(weights)
                    weights = weights / weights.sum()
                    
                    # 根據權重選號
                    selected = np.random.choice(numbers, size=n, replace=False, p=weights)
                    result = sorted([int(x) for x in selected.tolist()])  # 確保都是 int
                    logger.info(f"🧠 150筆時間加權智能選號 (隨機因子:{randomness_factor}): {result}")
                    return result
            except Exception as e:
                logger.warning(f"⚠️ 150筆時間加權選號失敗，使用隨機選號: {e}")
        
        return sorted(random.sample(numbers, n))
    elif strategy == 'balanced':
        # 平衡策略：純隨機選號
        result = sorted(random.sample(numbers, n))
        logger.info(f"⚖️ 150筆平衡策略 (隨機因子:{randomness_factor}): {result}")
        return result
    else:
        # 其他策略暫不使用
        return sorted(random.sample(numbers, n))

def select_top_weighted_numbers(nine_numbers, df, n=7):
    """
    從九顆號碼中選取加權最高的七顆
    使用智能選號的加權邏輯來排序九顆號碼
    """
    try:
        if df is not None:
            # 使用與智能選號相同的加權計算
            weighted_freq = compute_weighted_frequency(df)
            if weighted_freq:
                # 計算九顆號碼的加權分數
                number_scores = []
                for num in nine_numbers:
                    score = weighted_freq.get(num, 0.001)
                    number_scores.append((num, score))
                
                # 按加權分數排序（高分在前）
                number_scores.sort(key=lambda x: x[1], reverse=True)
                
                # 選取前七顆
                top_seven = [num for num, _ in number_scores[:n]]
                result = sorted(top_seven)
                logger.info(f"🎯 從九顆中選取加權最高的七顆: {result}")
                return result
    except Exception as e:
        logger.warning(f"⚠️ 七顆選號失敗，使用前七顆: {e}")
    
    # 備用方案：直接取前七顆
    return sorted(nine_numbers[:n])

def select_top_weighted_numbers_150(nine_numbers, df, n=7):
    """
    從九顆號碼中選取加權最高的七顆 (150筆版本)
    使用150筆智能選號的加權邏輯來排序九顆號碼
    """
    try:
        if df is not None:
            # 使用150筆加權計算
            weighted_freq = compute_weighted_frequency_150(df)
            if weighted_freq:
                # 計算九顆號碼的加權分數
                number_scores = []
                for num in nine_numbers:
                    score = weighted_freq.get(num, 0.001)
                    number_scores.append((num, score))
                
                # 按加權分數排序（高分在前）
                number_scores.sort(key=lambda x: x[1], reverse=True)
                
                # 選取前七顆
                top_seven = [num for num, _ in number_scores[:n]]
                result = sorted(top_seven)
                logger.info(f"🎯 從九顆中選取150筆加權最高的七顆: {result}")
                return result
    except Exception as e:
        logger.warning(f"⚠️ 150筆七顆選號失敗，使用前七顆: {e}")
    
    # 備用方案：直接取前七顆
    return sorted(nine_numbers[:n])

def log_predictions_to_excel(predictions, log_file="fantasy5_prediction_log.xlsx"):
    """記錄預測結果 (加州Fantasy 5版本)"""
    current_time = datetime.now()
    date_str = current_time.strftime("%Y-%m-%d")
    time_str = current_time.strftime("%H:%M:%S")
    
    # 準備記錄數據 - 加州Fantasy 5版本 (包含A/B測試)
    log_data = {
        '日期': date_str,
        '時間': time_str,
        '智能選號_九顆': str(predictions.get('smart_9', [])),
        '智能選號_七顆': str(predictions.get('smart_7', [])),
        '平衡策略_九顆': str(predictions.get('balanced_9', [])),
        '平衡策略_七顆': str(predictions.get('balanced_7', [])),
        '智能選號_九顆-150': str(predictions.get('smart_9_150', [])),
        '智能選號_七顆-150': str(predictions.get('smart_7_150', [])),
        '平衡策略_九顆-150': str(predictions.get('balanced_9_150', [])),
        '平衡策略_七顆-150': str(predictions.get('balanced_7_150', [])),
        '中獎號碼數': '',  # 留空，等待驗證
        '備註': f"加州Fantasy 5最佳策略(隨機因子0.3) - A/B測試365筆vs150筆 - {os.environ.get('GITHUB_WORKFLOW', 'Unknown')}",
        '驗證結果': ''  # 留空，等待驗證
    }
    
    # 檢查檔案是否存在
    log_path = Path(log_file)
    
    if log_path.exists():
        try:
            existing_df = pd.read_excel(log_file, engine='openpyxl')
            
            # 檢查是否有相同日期和時間的記錄（避免重複執行產生的記錄）
            same_datetime_records = existing_df[
                (existing_df['日期'] == date_str) & 
                (existing_df['時間'] == time_str)
            ]
            
            if len(same_datetime_records) > 0:
                # 有相同日期時間的記錄，更新最新的一筆（避免重複執行覆蓋）
                latest_same_datetime_index = same_datetime_records.index[-1]
                
                # 保留已驗證的結果（如果有的話）
                old_record = existing_df.loc[latest_same_datetime_index]
                if pd.notna(old_record.get('驗證結果', '')) and old_record.get('驗證結果', '') != '':
                    log_data['驗證結果'] = old_record['驗證結果']
                    log_data['中獎號碼數'] = old_record['中獎號碼數']
                    log_data['備註'] = f"相同時間更新（保留驗證結果） - {os.environ.get('GITHUB_WORKFLOW', 'Unknown')}"
                    logger.info("🔄 更新相同日期時間記錄，保留已驗證結果")
                else:
                    logger.info("🔄 更新相同日期時間記錄")
                
                # 更新該記錄
                for key, value in log_data.items():
                    if key in existing_df.columns:
                        existing_df.loc[latest_same_datetime_index, key] = value
                    else:
                        existing_df[key] = ''
                        existing_df.loc[latest_same_datetime_index, key] = value
                
                combined_df = existing_df
            else:
                # 沒有相同日期時間的記錄，新增一筆
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
        logger.info(f"✅ 加州Fantasy 5預測記錄已保存到: {log_file}")
        logger.info(f"   記錄時間: {date_str} {time_str}")
        return True
    except Exception as e:
        logger.error(f"❌ 保存預測記錄時發生錯誤: {e}")
        return False

def is_prediction_day(check_date=None):
    """
    檢查指定日期是否需要預測（加州Fantasy 5每日開獎，包括週日）
    Args:
        check_date: 要檢查的日期，None表示今天
    Returns:
        bool: True表示需要預測，False表示不需要
    """
    if check_date is None:
        check_date = datetime.now()
    elif isinstance(check_date, str):
        try:
            check_date = datetime.strptime(check_date, "%Y-%m-%d")
        except:
            logger.warning(f"⚠️ 日期格式錯誤: {check_date}")
            return True  # 預設為預測日
    
    # 加州Fantasy 5每日開獎，包括週日
    is_predict_day = True
    
    weekday_names = ['週一', '週二', '週三', '週四', '週五', '週六', '週日']
    logger.info(f"📅 檢查日期: {check_date.strftime('%Y-%m-%d')} ({weekday_names[check_date.weekday()]})")
    
    if is_predict_day:
        logger.info(f"✅ {weekday_names[check_date.weekday()]} 執行加州Fantasy 5預測（每日開獎）")
    else:
        logger.info(f"⏸️ {weekday_names[check_date.weekday()]} 跳過預測")
    
    return is_predict_day

def main():
    """主程式 - 加州Fantasy 5預測版本"""
    logger.info("🌅 加州Fantasy 5預測系統 - 上午預測版本")
    logger.info("="*60)
    
    # 檢查今天是否需要預測
    if not is_prediction_day():
        logger.info("⏸️ 今日跳過預測")
        return True  # 返回成功，因為這是預期的行為
    
    # 檢查輸入檔案
    excel_file = "fantasy5_hist.xlsx"
    if not Path(excel_file).exists():
        logger.error(f"❌ 找不到檔案: {excel_file}")
        return False
    
    try:
        # 載入資料
        df = load_lottery_excel(excel_file)
        logger.info(f"📊 成功載入 {len(df)} 筆加州Fantasy 5歷史資料")
        
        # 使用最佳隨機因子 0.3
        randomness_factor = 0.3
        predictions = {}
        
        logger.info("🎯 使用最佳隨機因子: 0.3")
        logger.info("="*60)
        
        # 生成九顆策略 (365筆)
        smart_9 = suggest_numbers('smart', n=9, df=df, randomness_factor=randomness_factor)
        balanced_9 = suggest_numbers('balanced', n=9, df=df, randomness_factor=randomness_factor)
        
        # 生成七顆策略（基於九顆選號）(365筆)
        smart_7 = select_top_weighted_numbers(smart_9, df, n=7)
        balanced_7 = select_top_weighted_numbers(balanced_9, df, n=7)
        
        # 生成九顆策略 (150筆A/B測試)
        smart_9_150 = suggest_numbers_150('smart', n=9, df=df, randomness_factor=randomness_factor)
        balanced_9_150 = suggest_numbers_150('balanced', n=9, df=df, randomness_factor=randomness_factor)
        
        # 生成七顆策略（基於九顆選號）(150筆A/B測試)
        smart_7_150 = select_top_weighted_numbers_150(smart_9_150, df, n=7)
        balanced_7_150 = select_top_weighted_numbers_150(balanced_9_150, df, n=7)
        
        # 儲存結果
        predictions['smart_9'] = smart_9
        predictions['smart_7'] = smart_7
        predictions['balanced_9'] = balanced_9
        predictions['balanced_7'] = balanced_7
        predictions['smart_9_150'] = smart_9_150
        predictions['smart_7_150'] = smart_7_150
        predictions['balanced_9_150'] = balanced_9_150
        predictions['balanced_7_150'] = balanced_7_150
        
        logger.info(f"📋 智能選號_九顆 (365筆): {smart_9}")
        logger.info(f"📋 智能選號_七顆 (365筆): {smart_7}")
        logger.info(f"📋 平衡策略_九顆 (365筆): {balanced_9}")
        logger.info(f"📋 平衡策略_七顆 (365筆): {balanced_7}")
        logger.info(f"📋 智能選號_九顆-150: {smart_9_150}")
        logger.info(f"📋 智能選號_七顆-150: {smart_7_150}")
        logger.info(f"📋 平衡策略_九顆-150: {balanced_9_150}")
        logger.info(f"📋 平衡策略_七顆-150: {balanced_7_150}")
        
        # 記錄預測結果
        success = log_predictions_to_excel(predictions, "fantasy5_prediction_log.xlsx")
        
        if success:
            logger.info("🎉 加州Fantasy 5預測完成並成功記錄!")
            
            # 檢查檔案是否確實建立
            if os.path.exists("fantasy5_prediction_log.xlsx"):
                file_size = os.path.getsize("fantasy5_prediction_log.xlsx")
                logger.info(f"📊 預測檔案已建立: fantasy5_prediction_log.xlsx ({file_size} bytes)")
            else:
                logger.warning("⚠️ 預測檔案可能建立失敗")
        else:
            logger.error("❌ 記錄預測結果時發生錯誤")
            return False
            
        return True
        
    except Exception as e:
        logger.error(f"❌ 預測過程發生錯誤: {str(e)}")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
