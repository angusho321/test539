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

# Fantasy5 高機率特徵常數 (基於 Gemini 分析)
HOT_NUMBERS = [33, 10, 32, 39, 11, 14, 6, 20, 17, 25]  # Top 10 熱門號

WEEKDAY_STRONG_NUMBERS = {
    0: [17, 2, 36, 7, 32],    # 週一
    1: [21, 14, 29, 20, 3],   # 週二
    2: [37, 33, 39, 24, 18],  # 週三
    3: [6, 27, 20, 33, 11],   # 週四
    4: [37, 17, 11, 15, 6],   # 週五
    5: [38, 25, 1, 17, 3],    # 週六
    6: [10, 35, 9, 20, 32],   # 週日
}

SPECIAL_TAIL_NUMBERS = [9, 3, 7]  # 強勢尾數

# 從原本的 lottery_analysis.py 複製核心函數
def load_lottery_excel(excel_path: str):
    """讀入 .xlsx 開獎紀錄"""
    df = pd.read_excel(excel_path, engine='openpyxl')
    
    # 確保日期欄位是 datetime 類型
    if '日期' in df.columns:
        # 嘗試多種日期格式解析
        try:
            df['日期'] = pd.to_datetime(df['日期'], format='mixed', errors='coerce')
            # 如果還有無法解析的，嘗試其他格式
            if df['日期'].isna().any():
                df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
            if df['日期'].isna().any():
                df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d', errors='coerce')
            # 最後嘗試不指定格式
            if df['日期'].isna().any():
                df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
        except Exception as e:
            logger.warning(f"⚠️ 日期轉換警告: {e}")
            # 如果轉換失敗，嘗試不指定格式
            df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    
    return df



def compute_weighted_frequency(df, decay_factor=0.95, recent_days=365):
    """
    計算時間加權的號碼頻率
    越近期的記錄權重越高，避免資料鈍化問題
    """
    try:
        # 確保日期欄位是 datetime 類型
        if '日期' in df.columns:
            if not pd.api.types.is_datetime64_any_dtype(df['日期']):
                # 如果日期欄位不是 datetime 類型，嘗試轉換
                df = df.copy()
                df['日期'] = pd.to_datetime(df['日期'], format='mixed', errors='coerce')
                if df['日期'].isna().any():
                    df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
                if df['日期'].isna().any():
                    df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d', errors='coerce')
                if df['日期'].isna().any():
                    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
        
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


def check_odd_even_ratio(numbers):
    """檢查奇偶比例是否為 3:2 或 2:3 (Fantasy5偏好3:2)"""
    first_five = sorted(numbers)[:5]
    odd_count = sum(1 for n in first_five if n % 2 == 1)
    return odd_count in [2, 3]


def check_sum_range(numbers, min_sum=86, max_sum=118):
    """檢查和值是否在範圍內 (Fantasy5: 86-118)"""
    first_five = sorted(numbers)[:5]
    total = sum(first_five)
    return min_sum <= total <= max_sum


def has_consecutive(numbers):
    """檢查是否有連號"""
    sorted_nums = sorted(numbers)
    for i in range(len(sorted_nums) - 1):
        if sorted_nums[i+1] - sorted_nums[i] == 1:
            return True
    return False


def count_hot_numbers(numbers, hot_list=HOT_NUMBERS):
    """計算熱門號數量"""
    return sum(1 for n in numbers if n in hot_list)


def count_special_tails(numbers, special_tails=SPECIAL_TAIL_NUMBERS):
    """計算特殊尾數數量"""
    return sum(1 for n in numbers if (n % 10) in special_tails)


def calculate_score(numbers, hot_numbers, weekday_strong_numbers, special_tail_numbers, 
                   current_weekday, require_consecutive=False):
    """
    計算號碼組合的綜合評分
    
    必要條件 (60分):
      - 和值 86-118: +30
      - 奇偶比例 3:2或2:3: +30
    
    加分項目 (最高40分):
      - 熱門號 1-2個: +10~20
      - 星期強勢號: +5~15
      - 連號: +8
      - 特殊尾數 2+個: +3~10
    """
    score = 0
    
    # 1. 檢查和值 (必要，+30)
    if not check_sum_range(numbers):
        return False, 0
    score += 30
    
    # 2. 檢查奇偶比例 (必要，+30)
    if not check_odd_even_ratio(numbers):
        return False, 0
    score += 30
    
    # 3. 檢查熱門號 (加分項)
    hot_count = count_hot_numbers(numbers, hot_numbers)
    if hot_count >= 1:
        score += min(hot_count * 10, 20)
    
    # 4. 檢查星期強勢號 (加分項)
    if current_weekday is not None:
        strong_nums = weekday_strong_numbers.get(current_weekday, [])
        strong_count = sum(1 for n in numbers if n in strong_nums)
        if strong_count >= 1:
            score += min(strong_count * 5, 15)
    
    # 5. 檢查連號 (條件性，歷史39%)
    has_consec = has_consecutive(numbers)
    if require_consecutive and not has_consec:
        score -= 8  # 要求連號但沒有，扣分
    elif has_consec:
        score += 8
    
    # 6. 檢查特殊尾數 (加分項)
    tail_count = count_special_tails(numbers, special_tail_numbers)
    if tail_count >= 2:
        score += min(tail_count * 3, 10)
    
    return True, score


def suggest_numbers(strategy='smart', n=9, historical_stats=None, df=None, randomness_factor=0.3, 
                   target_weekday=None):
    """產生建議號碼 (加州Fantasy 5版本 - 高機率特徵策略)"""
    numbers = list(range(1, 40))  # Fantasy 5也是1-39

    if strategy == 'smart':
        # 高機率特徵策略：時間加權 + 6大規則過濾
        if df is not None:
            try:
                weighted_freq = compute_weighted_frequency(df)
                if weighted_freq:
                    # 決定是否要求連號（35%機率，基於歷史39%）
                    require_consecutive = random.random() < 0.35
                    
                    best_numbers = None
                    best_score = -1
                    
                    # 嘗試多次直到找到符合所有過濾的組合
                    for attempt in range(500):
                        # 根據加權頻率選號
                        weights = [weighted_freq.get(num, 0.001) for num in numbers]
                        
                        # 加入隨機性
                        adjusted_weights = []
                        for w in weights:
                            adjusted_weights.append(w * (1 - randomness_factor) + random.random() * randomness_factor)
                        
                        # 對熱門號加權 ×1.3
                        for i, num in enumerate(numbers):
                            if num in HOT_NUMBERS:
                                adjusted_weights[i] *= 1.3
                        
                        # 對星期強勢號加權 ×1.2
                        if target_weekday is not None:
                            strong_nums = WEEKDAY_STRONG_NUMBERS.get(target_weekday, [])
                            for i, num in enumerate(numbers):
                                if num in strong_nums:
                                    adjusted_weights[i] *= 1.2
                        
                        # 對特殊尾數加權 ×1.1
                        for i, num in enumerate(numbers):
                            if (num % 10) in SPECIAL_TAIL_NUMBERS:
                                adjusted_weights[i] *= 1.1
                        
                        # 正規化權重
                        total = sum(adjusted_weights)
                        if total > 0:
                            adjusted_weights = [w / total for w in adjusted_weights]
                        
                        # 根據權重選號
                        try:
                            selected = np.random.choice(numbers, size=n, replace=False, p=adjusted_weights)
                            candidate = sorted([int(x) for x in selected])
                        except:
                            candidate = sorted(random.sample(numbers, n))
                        
                        # 計算評分
                        valid, score = calculate_score(
                            candidate, HOT_NUMBERS, WEEKDAY_STRONG_NUMBERS, 
                            SPECIAL_TAIL_NUMBERS, target_weekday, require_consecutive
                        )
                        
                        if not valid:
                            continue
                        
                        # 更新最佳組合
                        if score > best_score:
                            best_score = score
                            best_numbers = candidate
                            
                            # 如果分數夠高，提前返回
                            if score >= 70:
                                sum_val = sum(sorted(best_numbers)[:5])
                                logger.info(f"🎯 高機率特徵選號 (分數:{best_score}, 和值:{sum_val}): {best_numbers}")
                                return best_numbers
                    
                    # 如果找到有效組合，返回最佳組合
                    if best_numbers is not None:
                        sum_val = sum(sorted(best_numbers)[:5])
                        logger.info(f"🎯 高機率特徵選號 (分數:{best_score}, 和值:{sum_val}): {best_numbers}")
                        return best_numbers
                    
                    # 如果所有嘗試都失敗，使用原始方法
                    logger.warning(f"⚠️ 500次嘗試後仍無有效組合，使用時間加權選號")
            except Exception as e:
                logger.warning(f"⚠️ 高機率特徵選號失敗，使用時間加權選號: {e}")
        
        # 備用方案：原始時間加權選號
        if df is not None:
            try:
                weighted_freq = compute_weighted_frequency(df)
                if weighted_freq:
                    weights = [weighted_freq.get(num, 0.001) for num in numbers]
                    for i in range(len(weights)):
                        weights[i] = weights[i] * (1 - randomness_factor) + random.random() * randomness_factor
                    weights = np.array(weights)
                    weights = weights / weights.sum()
                    selected = np.random.choice(numbers, size=n, replace=False, p=weights)
                    result = sorted([int(x) for x in selected.tolist()])
                    logger.info(f"🧠 時間加權智能選號 (隨機因子:{randomness_factor}): {result}")
                    return result
            except Exception as e:
                logger.warning(f"⚠️ 時間加權選號失敗，使用隨機選號: {e}")
        
        return sorted(random.sample(numbers, n))
    elif strategy == 'balanced':
        # 平衡策略：純隨機選號
        result = sorted(random.sample(numbers, n))
        logger.info(f"⚖️ 平衡策略 (隨機因子:{randomness_factor}): {result}")
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


def log_predictions_to_excel(predictions, log_file="fantasy5_prediction_log.xlsx"):
    """記錄預測結果 (加州Fantasy 5版本)"""
    current_time = datetime.now()
    date_str = current_time.strftime("%Y-%m-%d")
    time_str = current_time.strftime("%H:%M:%S")
    
    # 準備記錄數據 - 加州Fantasy 5版本
    log_data = {
        '日期': date_str,
        '時間': time_str,
        '智能選號_九顆': str(predictions.get('smart_9', [])),
        '智能選號_七顆': str(predictions.get('smart_7', [])),
        '平衡策略_九顆': str(predictions.get('balanced_9', [])),
        '平衡策略_七顆': str(predictions.get('balanced_7', [])),
        '中獎號碼數': '',  # 留空，等待驗證
        '備註': f"加州Fantasy 5高機率特徵策略(6大規則) - {os.environ.get('GITHUB_WORKFLOW', 'Unknown')}",
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
                    log_data['備註'] = f"加州Fantasy 5高機率特徵策略(6大規則) - 相同時間更新（保留驗證結果） - {os.environ.get('GITHUB_WORKFLOW', 'Unknown')}"
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
        
        logger.info("🎯 使用高機率特徵策略 (6大規則)")
        logger.info("   1. 熱門號: 33, 10, 32, 39, 11, 14, 6, 20, 17, 25")
        logger.info("   2. 和值範圍: 86-118")
        logger.info("   3. 奇偶比例: 3:2 或 2:3")
        logger.info("   4. 星期效應: 已啟用")
        logger.info("   5. 連號機率: 35%")
        logger.info("   6. 特殊尾數: 9, 3, 7")
        logger.info("="*60)
        
        # 獲取今天星期
        today_weekday = datetime.now().weekday()
        
        # 生成九顆策略
        smart_9 = suggest_numbers('smart', n=9, df=df, randomness_factor=randomness_factor, target_weekday=today_weekday)
        balanced_9 = suggest_numbers('balanced', n=9, df=df, randomness_factor=randomness_factor)
        
        # 生成七顆策略（基於九顆選號）
        smart_7 = select_top_weighted_numbers(smart_9, df, n=7)
        balanced_7 = select_top_weighted_numbers(balanced_9, df, n=7)
        
        # 儲存結果
        predictions['smart_9'] = smart_9
        predictions['smart_7'] = smart_7
        predictions['balanced_9'] = balanced_9
        predictions['balanced_7'] = balanced_7
        
        logger.info(f"📋 智能選號_九顆: {smart_9}")
        logger.info(f"📋 智能選號_七顆: {smart_7}")
        logger.info(f"📋 平衡策略_九顆: {balanced_9}")
        logger.info(f"📋 平衡策略_七顆: {balanced_7}")
        
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
