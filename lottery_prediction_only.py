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
EV_LOOKBACK_DAYS = 240
EV_DECAY = 0.99
EV_W_MOMENTUM = 1.5
EV_MOMENTUM_K = 7
EV_W_OVERDUE = 0.0
EV_W_WEEKDAY = 0.2
EV_W_PAIR = 0.3
EV_W_REPEAT = 0.25
EV_W_REGIME = 0.3

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
    """檢查奇偶比例是否為 2:3 或 3:2"""
    first_five = sorted(numbers)[:5]
    odd_count = sum(1 for n in first_five if n % 2 == 1)
    return odd_count in [2, 3]


def check_sum_range(numbers, min_sum=83, max_sum=116):
    """檢查和值是否在範圍內（歷史覆蓋率51.39%）"""
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


def count_hot_numbers(numbers, hot_list=[1, 27, 11, 17, 23]):
    """計算包含多少個熱門號"""
    return sum(1 for n in numbers if n in hot_list)


def count_special_tails(numbers, special_tails=[1, 4, 7]):
    """計算有多少個號碼的尾數是 1, 4, 7"""
    return sum(1 for n in numbers if (n % 10) in special_tails)


def apply_high_prob_filters(numbers, target_weekday=None, require_consecutive=False):
    """
    應用高機率特徵過濾
    
    Args:
        numbers: 號碼列表
        target_weekday: 目標星期（0-6）
        require_consecutive: 是否要求有連號
    
    Returns:
        (通過過濾, 分數)
    """
    # 星期強勢號碼（根據歷史統計）
    weekday_strong_numbers = {
        0: [6, 24, 17, 16, 1],      # 週一
        1: [38, 25, 23, 11, 19],    # 週二
        2: [38, 6, 1, 34, 23],      # 週三
        3: [27, 37, 35, 14, 17],    # 週四
        4: [8, 39, 17, 31, 9],      # 週五
        5: [8, 35, 24, 4, 20],      # 週六
    }
    
    score = 0
    reasons = []
    
    # 1. 檢查和值 (必須) - 83-116範圍覆蓋51.39%歷史
    if not check_sum_range(numbers, 83, 116):
        return False, 0
    score += 30
    reasons.append("和值✓")
    
    # 2. 檢查奇偶比例 (必須) - 2:3或3:2覆蓋63.59%歷史
    if not check_odd_even_ratio(numbers):
        return False, 0
    score += 30
    reasons.append("奇偶✓")
    
    # 3. 檢查熱門號 (1-2個，加分項)
    hot_count = count_hot_numbers(numbers)
    if hot_count >= 1:
        score += min(hot_count * 10, 20)
        reasons.append(f"熱號x{hot_count}✓")
    
    # 4. 檢查星期強勢號 (加分項)
    if target_weekday is not None:
        strong_nums = weekday_strong_numbers.get(target_weekday, [])
        strong_count = sum(1 for n in numbers if n in strong_nums)
        if strong_count >= 1:
            score += min(strong_count * 5, 15)
            reasons.append(f"星期x{strong_count}✓")
    
    # 5. 檢查連號 (條件性，歷史42.79%)
    has_consec = has_consecutive(numbers)
    if require_consecutive and not has_consec:
        score -= 10
    elif has_consec:
        score += 10
        reasons.append("連號✓")
    
    # 6. 檢查特殊尾數 (加分項，覆蓋31.75%歷史)
    tail_count = count_special_tails(numbers)
    if tail_count >= 2:
        score += min(tail_count * 3, 10)
        reasons.append(f"尾數x{tail_count}✓")
    
    return True, score


def suggest_numbers(strategy='smart', n=9, historical_stats=None, df=None, randomness_factor=0.3, 
                   use_high_prob=True, target_weekday=None, max_attempts=500):
    """
    產生建議號碼 (支援高機率特徵過濾)
    Args:
        strategy: 策略名稱
        n: 選號數量
        df: 歷史資料
        randomness_factor: 隨機因子
        use_high_prob: 是否使用高機率特徵過濾
        target_weekday: 目標星期（0-6，0=週一）
        max_attempts: 最大嘗試次數
    """
    numbers = list(range(1, 40))
    
    # 熱門號碼
    hot_numbers = [1, 27, 11, 17, 23]
    
    # 特殊尾數
    special_tails = [1, 4, 7]
    
    # 星期強勢號碼
    weekday_strong_numbers = {
        0: [6, 24, 17, 16, 1],
        1: [38, 25, 23, 11, 19],
        2: [38, 6, 1, 34, 23],
        3: [27, 37, 35, 14, 17],
        4: [8, 39, 17, 31, 9],
        5: [8, 35, 24, 4, 20],
    }

    if strategy == 'smart':
        # 智能選號：時間加權 + 高機率特徵過濾
        if df is not None:
            try:
                weighted_freq = compute_weighted_frequency(df)
                if weighted_freq:
                    # 決定是否要求連號（40%機率）
                    require_consecutive = random.random() < 0.4
                    
                    best_numbers = None
                    best_score = -1
                    
                    # 嘗試多次直到找到符合所有過濾的組合
                    for attempt in range(max_attempts):
                        # 根據加權頻率選號
                        weights = [weighted_freq.get(num, 0.001) for num in numbers]
                        
                        # 加入隨機性
                        adjusted_weights = []
                        for w in weights:
                            adjusted_weights.append(w * (1 - randomness_factor) + random.random() * randomness_factor)
                        
                        # 如果啟用高機率特徵，對特定號碼加權
                        if use_high_prob:
                            # 對熱門號加權
                            for i, num in enumerate(numbers):
                                if num in hot_numbers:
                                    adjusted_weights[i] *= 1.3
                            
                            # 對星期強勢號加權
                            if target_weekday is not None:
                                strong_nums = weekday_strong_numbers.get(target_weekday, [])
                                for i, num in enumerate(numbers):
                                    if num in strong_nums:
                                        adjusted_weights[i] *= 1.2
                            
                            # 對特殊尾數加權
                            for i, num in enumerate(numbers):
                                if (num % 10) in special_tails:
                                    adjusted_weights[i] *= 1.1
                        
                        # 正規化權重
                        adjusted_weights = np.array(adjusted_weights)
                        adjusted_weights = adjusted_weights / adjusted_weights.sum()
                        
                        # 根據權重選號
                        selected = np.random.choice(numbers, size=n, replace=False, p=adjusted_weights)
                        result = sorted([int(x) for x in selected.tolist()])
                        
                        # 檢查是否通過高機率特徵過濾
                        if use_high_prob:
                            passed, score = apply_high_prob_filters(result, target_weekday, require_consecutive)
                            
                            if passed and score > best_score:
                                best_numbers = result
                                best_score = score
                                
                                # 如果分數夠高，提前返回
                                if score >= 70:
                                    break
                        else:
                            # 不使用過濾，直接返回
                            logger.info(f"🧠 時間加權選號 (和值:{sum(result[:5])}): {result}")
                            return result
                    
                    # 如果找到符合規則的組合，返回最佳的
                    if best_numbers is not None:
                        logger.info(f"🧠 高機率特徵選號 (分數:{best_score}, 和值:{sum(best_numbers[:5])}): {best_numbers}")
                        return best_numbers
                    
                    # 如果超過最大嘗試次數，返回最後一次結果
                    logger.warning(f"⚠️ 超過 {max_attempts} 次嘗試，使用備用選號")
                    return result
            except Exception as e:
                logger.warning(f"⚠️ 智能選號失敗: {e}")
        
        return sorted(random.sample(numbers, n))
    elif strategy == 'balanced':
        # 平衡策略：純隨機選號
        result = sorted(random.sample(numbers, n))
        logger.info(f"⚖️ 平衡策略: {result}")
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


def _build_ev_scores(df, target_date):
    train_df = df[(df['日期'] < target_date) & (df['日期'] >= target_date - pd.Timedelta(days=EV_LOOKBACK_DAYS))]
    if len(train_df) == 0:
        train_df = df[df['日期'] < target_date].copy()
    if len(train_df) == 0:
        return np.zeros(40, dtype=float)
    scores = np.zeros(40, dtype=float)
    days_ago = (target_date - train_df['日期']).dt.days.clip(lower=0).to_numpy()
    row_weights = np.power(EV_DECAY, days_ago)
    draw_matrix = train_df[['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']].to_numpy(dtype=int)
    for i in range(draw_matrix.shape[0]):
        w = row_weights[i]
        for num in draw_matrix[i]:
            scores[num] += w
    if EV_W_MOMENTUM > 0:
        recent = train_df.tail(EV_MOMENTUM_K)[['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']].to_numpy(dtype=int)
        if len(recent) > 0:
            momentum = np.zeros(40, dtype=float)
            for row in recent:
                for num in row:
                    momentum[num] += 1.0
            scores += EV_W_MOMENTUM * (momentum / len(recent))
    if EV_W_OVERDUE > 0:
        last_seen = np.full(40, -1, dtype=int)
        for i in range(draw_matrix.shape[0]):
            for num in draw_matrix[i]:
                last_seen[num] = i
        overdue = np.zeros(40, dtype=float)
        end = len(draw_matrix) - 1
        for num in range(1, 40):
            overdue[num] = min(60, end - last_seen[num]) if last_seen[num] >= 0 else 60
        overdue = overdue / max(1.0, overdue[1:].max())
        scores += EV_W_OVERDUE * overdue
    if EV_W_WEEKDAY > 0:
        wd = int(target_date.weekday())
        wd_df = train_df[train_df['日期'].dt.weekday == wd]
        if len(wd_df) > 0:
            wd_scores = np.zeros(40, dtype=float)
            wd_draws = wd_df[['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']].to_numpy(dtype=int)
            for row in wd_draws:
                for num in row:
                    wd_scores[num] += 1.0
            wd_scores = wd_scores / len(wd_df)
            scores += EV_W_WEEKDAY * wd_scores
    scores[0] = -1e9
    return scores


def _normalize_scores(arr):
    v = np.array(arr, dtype=float)
    maxv = v.max() if v.size > 0 else 0.0
    if maxv <= 0:
        return np.zeros_like(v, dtype=float)
    return v / maxv


def _pair_boost(train_df):
    scores = np.zeros(40, dtype=float)
    if len(train_df) < 10:
        return scores
    draws = train_df[['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']].to_numpy(dtype=int)
    pair_counts = np.zeros((40, 40), dtype=float)
    for row in draws:
        r = sorted(int(x) for x in row)
        for i in range(len(r)):
            for j in range(i + 1, len(r)):
                a, b = r[i], r[j]
                pair_counts[a, b] += 1.0
                pair_counts[b, a] += 1.0
    for n in range(1, 40):
        scores[n] = pair_counts[n].sum()
    return _normalize_scores(scores)


def _repeat_boost(train_df):
    scores = np.zeros(40, dtype=float)
    if len(train_df) < 2:
        return scores
    draws = train_df[['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']].to_numpy(dtype=int)
    overlap_count = np.zeros(40, dtype=float)
    for i in range(1, len(draws)):
        prev_set = set(int(x) for x in draws[i - 1])
        curr = [int(x) for x in draws[i]]
        for num in curr:
            if num in prev_set:
                overlap_count[num] += 1.0
    return _normalize_scores(overlap_count)


def _regime_boost(train_df, short_days=30, long_days=180):
    scores = np.zeros(40, dtype=float)
    if len(train_df) == 0:
        return scores
    end_date = train_df['日期'].max()
    short_df = train_df[train_df['日期'] >= end_date - pd.Timedelta(days=short_days)]
    long_df = train_df[train_df['日期'] >= end_date - pd.Timedelta(days=long_days)]
    if len(short_df) == 0 or len(long_df) == 0:
        return scores
    short_freq = np.zeros(40, dtype=float)
    long_freq = np.zeros(40, dtype=float)
    for row in short_df[['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']].to_numpy(dtype=int):
        for n in row:
            short_freq[int(n)] += 1.0
    for row in long_df[['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']].to_numpy(dtype=int):
        for n in row:
            long_freq[int(n)] += 1.0
    short_freq = short_freq / max(1.0, len(short_df))
    long_freq = long_freq / max(1.0, len(long_df))
    delta = short_freq - long_freq
    delta = delta - delta.min()
    scores[1:] = delta[1:]
    return _normalize_scores(scores)


def _build_ev_scores_enhanced(df, target_date):
    base = _build_ev_scores(df, target_date)
    train_df = df[(df['日期'] < target_date) & (df['日期'] >= target_date - pd.Timedelta(days=EV_LOOKBACK_DAYS))]
    if len(train_df) == 0:
        train_df = df[df['日期'] < target_date].copy()
    scores = base.copy()
    scores += EV_W_PAIR * _pair_boost(train_df)
    scores += EV_W_REPEAT * _repeat_boost(train_df)
    scores += EV_W_REGIME * _regime_boost(train_df)
    scores[0] = -1e9
    return scores


def suggest_ev_numbers(df, n, target_date):
    scores = _build_ev_scores_enhanced(df, target_date)
    selected = np.argsort(scores)[::-1][:n]
    return sorted(int(x) for x in selected.tolist())


def log_predictions_to_excel(predictions, log_file="prediction_log.xlsx"):
    """記錄預測結果 (僅預測版本)"""
    current_time = datetime.now()
    date_str = current_time.strftime("%Y-%m-%d")
    time_str = current_time.strftime("%H:%M:%S")
    
    # 準備記錄數據
    log_data = {
        '日期': date_str,
        '時間': time_str,
        '智能選號_九顆': str(predictions.get('smart_9', [])),
        '智能選號_七顆': str(predictions.get('smart_7', [])),
        'EV策略_九顆': str(predictions.get('ev_9', [])),
        'EV策略_七顆': str(predictions.get('ev_7', [])),
        '中獎號碼數': '',  # 留空，等待驗證
        '備註': f"539智能+EV對照策略 - {os.environ.get('GITHUB_WORKFLOW', 'Unknown')}",
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
                    log_data['備註'] = f"539智能+EV對照策略（保留驗證結果） - {os.environ.get('GITHUB_WORKFLOW', 'Unknown')}"
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
        
        # 使用高機率特徵策略
        randomness_factor = 0.3
        use_high_prob = True
        
        # 取得今天星期幾
        today_weekday = datetime.now().weekday()
        
        predictions = {}
        
        logger.info("🎯 使用高機率特徵策略（6大規則）")
        logger.info(f"   1. 熱門號: 1, 27, 11, 17, 23")
        logger.info(f"   2. 和值範圍: 83-116")
        logger.info(f"   3. 奇偶比例: 2:3 或 3:2")
        logger.info(f"   4. 星期效應: 已啟用")
        logger.info(f"   5. 連號機率: 40%")
        logger.info(f"   6. 特殊尾數: 1, 4, 7")
        logger.info("="*60)
        
        # 生成智能九顆策略（帶高機率特徵過濾）
        smart_9 = suggest_numbers('smart', n=9, df=df, randomness_factor=randomness_factor,
                                 use_high_prob=use_high_prob, target_weekday=today_weekday)
        # 生成EV九顆策略（近一年回測最佳參數）
        ev_9 = suggest_ev_numbers(df, n=9, target_date=datetime.now())
        
        # 生成七顆策略（智能由智能九顆衍生，EV獨立選號）
        smart_7 = select_top_weighted_numbers(smart_9, df, n=7)
        ev_7 = suggest_ev_numbers(df, n=7, target_date=datetime.now())
        
        # 儲存結果
        predictions['smart_9'] = smart_9
        predictions['smart_7'] = smart_7
        predictions['ev_9'] = ev_9
        predictions['ev_7'] = ev_7
        
        logger.info(f"📋 智能選號_九顆: {smart_9}")
        logger.info(f"📋 智能選號_七顆: {smart_7}")
        logger.info(f"📋 EV策略_九顆: {ev_9}")
        logger.info(f"📋 EV策略_七顆: {ev_7}")
        
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
