#!/usr/bin/env python3
"""
彩票預測程式 - 僅預測版本
專門用於上午時段，只生成預測不驗證
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

def compute_num_frequency(df: pd.DataFrame, recent_periods=None):
    """
    回傳每個號碼的頻次
    Args:
        df: 彩票歷史資料
        recent_periods: 只計算最近N期的資料，None表示使用全部資料
    """
    # 如果指定了近期期數，只取最近的記錄
    if recent_periods is not None and recent_periods > 0:
        analysis_df = df.tail(recent_periods).copy()
        logger.info(f"📊 使用最近 {len(analysis_df)} 期資料計算熱號冷號")
    else:
        analysis_df = df.copy()
        logger.info(f"📊 使用全部 {len(analysis_df)} 期資料計算熱號冷號")
    
    nums = analysis_df[['號碼1','號碼2','號碼3','號碼4','號碼5']].values.ravel()
    freq = pd.Series(nums).value_counts().sort_index()
    freq.index.name = '號碼'
    freq.name = '頻次'
    return freq

def get_hot_cold_numbers(freq: pd.Series, top_n=9, bottom_n=9):
    """取得熱號冷號，確保不重複"""
    # 排序所有號碼按頻率
    sorted_freq = freq.sort_values(ascending=False)
    
    # 取前 top_n 個作為熱號
    hot_numbers = sorted_freq.head(top_n).index.tolist()
    
    # 取後 bottom_n 個作為冷號，但排除已在熱號中的
    all_numbers = sorted_freq.index.tolist()
    cold_candidates = []
    
    # 從最低頻率開始選取，避免與熱號重複
    for num in reversed(all_numbers):
        if num not in hot_numbers and len(cold_candidates) < bottom_n:
            cold_candidates.append(num)
    
    cold_numbers = sorted(cold_candidates)
    
    # 驗證無重複
    overlap = set(hot_numbers) & set(cold_numbers)
    if overlap:
        logger.warning(f"⚠️ 熱號冷號重複: {overlap}")
    
    return hot_numbers, cold_numbers

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

def suggest_numbers(strategy='smart', n=9, historical_stats=None, df=None):
    """產生建議號碼 (升級版本，支援時間加權)"""
    numbers = list(range(1, 40))
    global hot_numbers, cold_numbers

    if strategy == 'random':
        return sorted(random.sample(numbers, n))
    elif strategy == 'hot':
        # 熱號策略：直接使用熱號（已確保有9個）
        sel = hot_numbers.copy()
        if len(sel) >= n:
            return sorted(random.sample(sel, n))
        else:
            # 如果熱號不足，從中性號碼池補足（排除冷號）
            neutral_numbers = [x for x in numbers if x not in hot_numbers and x not in cold_numbers]
            remain = neutral_numbers.copy()
            if len(remain) >= (n - len(sel)):
                sel += random.sample(remain, n - len(sel))
            else:
                sel += remain  # 如果中性號碼也不足，全部加入
            return sorted(sel)
    elif strategy == 'cold':
        # 冷號策略：直接使用冷號（已確保有9個）
        sel = cold_numbers.copy()
        if len(sel) >= n:
            return sorted(random.sample(sel, n))
        else:
            # 如果冷號不足，從中性號碼池補足（排除熱號）
            neutral_numbers = [x for x in numbers if x not in hot_numbers and x not in cold_numbers]
            remain = neutral_numbers.copy()
            if len(remain) >= (n - len(sel)):
                sel += random.sample(remain, n - len(sel))
            else:
                sel += remain  # 如果中性號碼也不足，全部加入
            return sorted(sel)
    elif strategy == 'smart':
        # 新的時間加權智能選號
        if df is not None:
            try:
                weighted_freq = compute_weighted_frequency(df)
                if weighted_freq:
                    # 根據加權頻率選號
                    weights = [weighted_freq.get(num, 0.001) for num in numbers]
                    
                    # 加入隨機性避免過度依賴歷史
                    randomness_factor = 0.3
                    for i in range(len(weights)):
                        weights[i] = weights[i] * (1 - randomness_factor) + random.random() * randomness_factor
                    
                    # 正規化權重
                    weights = np.array(weights)
                    weights = weights / weights.sum()
                    
                    # 根據權重選號
                    selected = np.random.choice(numbers, size=n, replace=False, p=weights)
                    result = sorted([int(x) for x in selected.tolist()])  # 確保都是 int
                    logger.info(f"🧠 時間加權智能選號: {result}")
                    return result
            except Exception as e:
                logger.warning(f"⚠️ 時間加權選號失敗，使用隨機選號: {e}")
        
        return sorted(random.sample(numbers, n))
    elif strategy == 'never_drawn':
        # 從未開出組合策略
        if df is not None:
            try:
                never_drawn = find_never_drawn_combinations(df, sample_size=100)
                if never_drawn:
                    # 隨機選擇一個從未開出的組合
                    selected_combo = random.choice(never_drawn)
                    result = [int(x) for x in selected_combo]  # 確保所有元素都是 int
                    # 如果號碼不足，隨機補充
                    if len(result) < n:
                        remaining = [num for num in numbers if num not in result]
                        result.extend(random.sample(remaining, n - len(result)))
                    final_result = sorted([int(x) for x in result[:n]])  # 再次確保都是 int
                    logger.info(f"💎 從未開出組合選號: {final_result}")
                    return final_result
            except Exception as e:
                logger.warning(f"⚠️ 從未開出組合選號失敗: {e}")
        
        return sorted(random.sample(numbers, n))
    elif strategy == 'fusion':
        # 融合策略：綜合多個策略的優勢
        if df is not None:
            return create_fusion_strategy(df, n)
        else:
            # 備用方案：隨機選擇
            return sorted(random.sample(numbers, n))
    else:  # balanced
        return sorted(random.sample(numbers, n))

def create_fusion_strategy(df, n=9):
    """創建融合策略：綜合多個策略的優勢"""
    try:
        from collections import Counter, defaultdict
        
        # 計算熱號冷號
        freq_series = compute_num_frequency(df, recent_periods=50)
        hot_numbers, cold_numbers = get_hot_cold_numbers(freq_series, top_n=9, bottom_n=9)
        
        # 1. 分析各策略的號碼偏好
        strategy_preferences = {}
        strategies = ['智能選號', '平衡策略', '隨機選號', '熱號優先', '冷號優先', '未開組合']
        
        # 模擬各策略的選號偏好
        strategy_preferences['智能選號'] = Counter()
        strategy_preferences['平衡策略'] = Counter()
        strategy_preferences['隨機選號'] = Counter()
        strategy_preferences['熱號優先'] = Counter(hot_numbers)
        strategy_preferences['冷號優先'] = Counter(cold_numbers)
        strategy_preferences['未開組合'] = Counter()
        
        # 2. 基於歷史數據分析號碼表現
        number_scores = defaultdict(float)
        
        # 維度1: 策略一致性 (多個策略都選中的號碼得分更高)
        for strategy, preferences in strategy_preferences.items():
            for number in preferences:
                number_scores[number] += 1
        
        # 維度2: 歷史中獎表現
        if '驗證結果' in df.columns:
            verified_df = df[df['驗證結果'].notna() & (df['驗證結果'] != '')]
            winning_numbers = []
            for _, row in verified_df.iterrows():
                # 解析驗證結果中的中獎號碼
                result_str = str(row['驗證結果'])
                if '[' in result_str and ']' in result_str:
                    try:
                        numbers_str = result_str.split('[')[1].split(']')[0]
                        numbers = [int(x.strip()) for x in numbers_str.split(',') if x.strip().isdigit()]
                        winning_numbers.extend(numbers)
                    except:
                        pass
            
            winning_freq = Counter(winning_numbers)
            for number, count in winning_freq.items():
                number_scores[number] += count * 0.5  # 中獎號碼額外加分
        
        # 維度3: 時間權重 (最近的表現權重更高)
        df['日期'] = pd.to_datetime(df['日期'])
        recent_df = df[df['日期'] >= df['日期'].max() - pd.Timedelta(days=7)]
        
        for _, row in recent_df.iterrows():
            for col in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']:
                if pd.notna(row[col]):
                    number_scores[int(row[col])] += 0.3  # 最近出現的號碼加分
        
        # 3. 生成融合策略預測
        if number_scores:
            # 選擇得分最高的號碼
            top_numbers = sorted(number_scores.items(), key=lambda x: x[1], reverse=True)
            selected_numbers = [num for num, _ in top_numbers[:n]]
            
            # 如果號碼不足，隨機補充
            if len(selected_numbers) < n:
                remaining = [num for num in range(1, 40) if num not in selected_numbers]
                selected_numbers.extend(random.sample(remaining, n - len(selected_numbers)))
            
            result = sorted(selected_numbers[:n])
            logger.info(f"🔗 融合策略選號: {result}")
            return result
        else:
            # 備用方案：智能選號 + 隨機選號
            smart_count = n // 2
            random_count = n - smart_count
            
            smart_numbers = hot_numbers[:smart_count] if len(hot_numbers) >= smart_count else hot_numbers
            random_numbers = random.sample([x for x in range(1, 40) if x not in smart_numbers], random_count)
            
            result = sorted(smart_numbers + random_numbers)
            logger.info(f"🔗 融合策略選號(備用): {result}")
            return result
            
    except Exception as e:
        logger.warning(f"⚠️ 融合策略選號失敗: {e}")
        # 備用方案：平衡策略
        hot_count = n // 2
        cold_count = n - hot_count
        return sorted(hot_numbers[:hot_count] + cold_numbers[:cold_count])

def find_never_drawn_combinations(df, combo_size=5, sample_size=100):
    """找出從未開出的號碼組合"""
    try:
        # 提取所有歷史開獎組合
        historical_combinations = set()
        
        for _, row in df.iterrows():
            numbers = []
            for col in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']:
                if pd.notna(row[col]):
                    numbers.append(int(row[col]))
            
            if len(numbers) == 5:
                combo = tuple(sorted(numbers))
                historical_combinations.add(combo)
        
        # 隨機生成並檢查從未開出的組合
        never_drawn = []
        attempts = 0
        max_attempts = sample_size * 20
        
        while len(never_drawn) < sample_size and attempts < max_attempts:
            random_combo = tuple(sorted([int(x) for x in np.random.choice(range(1, 40), combo_size, replace=False)]))
            attempts += 1
            
            if random_combo not in historical_combinations:
                never_drawn.append(random_combo)
        
        logger.info(f"💎 找到 {len(never_drawn)} 個從未開出的組合")
        return never_drawn
        
    except Exception as e:
        logger.error(f"❌ 分析從未開出組合失敗: {e}")
        return []

def log_predictions_to_excel(predictions, log_file="prediction_log.xlsx"):
    """記錄預測結果 (僅預測版本)"""
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
        '未開組合': str(predictions.get('never_drawn', [])),
        '融合策略': str(predictions.get('fusion', [])),
        '驗證結果': '',  # 留空，等待晚上驗證
        '中獎號碼數': '',  # 留空，等待驗證
        '備註': f"上午預測(含時間加權+未開組合) - {os.environ.get('GITHUB_WORKFLOW', 'Unknown')}"
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
        logger.info(f"✅ 預測記錄已保存到: {log_file}")
        logger.info(f"   記錄時間: {date_str} {time_str}")
        return True
    except Exception as e:
        logger.error(f"❌ 保存預測記錄時發生錯誤: {e}")
        return False

def is_prediction_day(check_date=None):
    """
    檢查指定日期是否需要預測（週一到週六）
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
    
    # 取得星期幾 (0=週一, 6=週日)
    weekday = check_date.weekday()
    
    # 週一到週六預測 (0-5)，週日不預測 (6)
    is_predict_day = weekday < 6
    
    weekday_names = ['週一', '週二', '週三', '週四', '週五', '週六', '週日']
    logger.info(f"📅 檢查日期: {check_date.strftime('%Y-%m-%d')} ({weekday_names[weekday]})")
    
    if is_predict_day:
        logger.info(f"✅ {weekday_names[weekday]} 執行預測")
    else:
        logger.info(f"⏸️ {weekday_names[weekday]} 跳過預測，留給週一自己預測")
    
    return is_predict_day

def main():
    """主程式 - 僅預測版本"""
    logger.info("🌅 539彩票預測系統 - 上午預測版本")
    logger.info("="*60)
    
    # 檢查今天是否需要預測
    if not is_prediction_day():
        logger.info("⏸️ 今日（週日）跳過預測")
        logger.info("💡 週一會重新開始預測週期")
        return True  # 返回成功，因為這是預期的行為
    
    # 檢查輸入檔案
    excel_file = "lottery_hist.xlsx"
    if not Path(excel_file).exists():
        logger.error(f"❌ 找不到檔案: {excel_file}")
        return False
    
    try:
        # 載入資料
        df = load_lottery_excel(excel_file)
        logger.info(f"📊 成功載入 {len(df)} 筆歷史資料")
        
        # 基本統計 - 使用最近期數計算熱號冷號
        recent_periods = 50  # 可調整：建議50-200期之間
        logger.info(f"🎯 分析範圍：最近 {recent_periods} 期開獎記錄")
        freq_series = compute_num_frequency(df, recent_periods=recent_periods)
        global hot_numbers, cold_numbers
        hot_numbers, cold_numbers = get_hot_cold_numbers(freq_series, top_n=9, bottom_n=9)
        
        logger.info(f"🔥 熱號: {hot_numbers}")
        logger.info(f"❄️ 冷號: {cold_numbers}")
        
        # 生成各策略的建議號碼
        strategies = ['smart', 'balanced', 'random', 'hot', 'cold', 'never_drawn', 'fusion']
        strategy_names = {
            'smart': '智能選號',
            'balanced': '平衡策略', 
            'random': '隨機選號',
            'hot': '熱號優先',
            'cold': '冷號優先',
            'never_drawn': '未開組合',
            'fusion': '融合策略'
        }
        predictions = {}
        
        for strategy in strategies:
            numbers = suggest_numbers(strategy, n=9, df=df)  # 傳遞 DataFrame
            predictions[strategy] = numbers
            display_name = strategy_names.get(strategy, strategy.upper())
            logger.info(f"📋 {display_name}: {numbers}")
        
        # 記錄預測結果
        success = log_predictions_to_excel(predictions, "prediction_log.xlsx")
        
        if success:
            logger.info("🎉 上午預測完成並成功記錄!")
            
            # 檢查檔案是否確實建立
            if os.path.exists("prediction_log.xlsx"):
                file_size = os.path.getsize("prediction_log.xlsx")
                logger.info(f"📊 預測檔案已建立: prediction_log.xlsx ({file_size} bytes)")
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
