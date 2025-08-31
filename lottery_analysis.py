# --------------------------------------------------------
#  lottery_analysis.py (xlsx 版)
# --------------------------------------------------------

import pandas as pd
import numpy as np
import random
from pathlib import Path
from datetime import datetime
# 導入驗證功能（如果檔案存在的話）
try:
    from verify_predictions import auto_verify_on_startup
    VERIFICATION_AVAILABLE = True
except ImportError:
    VERIFICATION_AVAILABLE = False

# --------------------------------------------------------
# 1. 讀取資料 --------------------------------------------------
def load_lottery_excel(excel_path: str):
    """
    讀入 .xlsx 開獎紀錄。
    欄位假設為：
        日期, 星期, 號碼1, 號碼2, 號碼3, 號碼4, 號碼5
    """
    df = pd.read_excel(excel_path, engine='openpyxl')
    # 若你使用的是舊版 pandas，engine 參數可去掉
    return df

# --------------------------------------------------------
# 2. 計算頻率與熱冷號 --------------------------------------------
def compute_num_frequency(df: pd.DataFrame):
    """回傳每個號碼的頻次（Series）"""
    nums = df[['號碼1','號碼2','號碼3','號碼4','號碼5']].values.ravel()
    freq = pd.Series(nums).value_counts().sort_index()
    freq.index.name = '號碼'
    freq.name = '頻次'
    return freq

def get_hot_cold_numbers(freq: pd.Series, top_n=6, bottom_n=6):
    hot_numbers = freq.nlargest(top_n).index.tolist()
    cold_numbers = freq.nsmallest(bottom_n).index.tolist()
    return hot_numbers, cold_numbers

# --------------------------------------------------------
# 3. 可視化功能已移除 (原 plot_frequency 函數)

# --------------------------------------------------------
# 4. 模式分析功能 -------------------------------------------
def analyze_single_draw_patterns(numbers_row):
    """分析單次開獎的模式"""
    nums = sorted(numbers_row)
    
    # 計算0頭號碼數量（1-9）
    zero_head_count = sum(1 for n in nums if n <= 9)
    
    # 計算1頭號碼數量（10-19）
    one_head_count = sum(1 for n in nums if 10 <= n <= 19)
    
    # 計算2頭號碼數量（20-29）
    two_head_count = sum(1 for n in nums if 20 <= n <= 29)
    
    # 計算3頭號碼數量（30-39）
    three_head_count = sum(1 for n in nums if 30 <= n <= 39)
    
    # 計算連號情況
    consecutive_count = 0
    max_consecutive = 1
    current_consecutive = 1
    has_consecutive = False  # 是否有連號
    
    for i in range(1, len(nums)):
        if nums[i] == nums[i-1] + 1:
            current_consecutive += 1
            has_consecutive = True
        else:
            max_consecutive = max(max_consecutive, current_consecutive)
            current_consecutive = 1
    max_consecutive = max(max_consecutive, current_consecutive)
    
    # 計算奇偶比例
    odd_count = sum(1 for n in nums if n % 2 == 1)
    even_count = 5 - odd_count
    
    # 計算大小號比例（1-19為小號，20-39為大號）
    small_count = sum(1 for n in nums if n <= 19)
    large_count = 5 - small_count
    
    # 計算跨度（最大號-最小號）
    span = nums[-1] - nums[0]
    
    return {
        'zero_head_count': zero_head_count,
        'one_head_count': one_head_count,
        'two_head_count': two_head_count,
        'three_head_count': three_head_count,
        'max_consecutive': max_consecutive,
        'has_consecutive': has_consecutive,
        'odd_count': odd_count,
        'even_count': even_count,
        'small_count': small_count,
        'large_count': large_count,
        'span': span
    }

def pattern_statistics(df: pd.DataFrame):
    """分析歷史開獎模式統計"""
    patterns = []
    
    for _, row in df.iterrows():
        numbers = [row['號碼1'], row['號碼2'], row['號碼3'], row['號碼4'], row['號碼5']]
        pattern = analyze_single_draw_patterns(numbers)
        patterns.append(pattern)
    
    patterns_df = pd.DataFrame(patterns)
    
    print("\n=== 歷史開獎模式統計 ===")
    
    # 0頭號碼統計
    print(f"\n0頭號碼(1-9)數量分布:")
    zero_head_dist = patterns_df['zero_head_count'].value_counts().sort_index()
    for count, freq in zero_head_dist.items():
        percentage = freq / len(patterns_df) * 100
        print(f"  {count}個: {freq}次 ({percentage:.1f}%)")
    
    # 1頭號碼統計
    print(f"\n1頭號碼(10-19)數量分布:")
    one_head_dist = patterns_df['one_head_count'].value_counts().sort_index()
    for count, freq in one_head_dist.items():
        percentage = freq / len(patterns_df) * 100
        print(f"  {count}個: {freq}次 ({percentage:.1f}%)")
    
    # 2頭號碼統計
    print(f"\n2頭號碼(20-29)數量分布:")
    two_head_dist = patterns_df['two_head_count'].value_counts().sort_index()
    for count, freq in two_head_dist.items():
        percentage = freq / len(patterns_df) * 100
        print(f"  {count}個: {freq}次 ({percentage:.1f}%)")
    
    # 3頭號碼統計
    print(f"\n3頭號碼(30-39)數量分布:")
    three_head_dist = patterns_df['three_head_count'].value_counts().sort_index()
    for count, freq in three_head_dist.items():
        percentage = freq / len(patterns_df) * 100
        print(f"  {count}個: {freq}次 ({percentage:.1f}%)")
    
    # 連號統計
    print(f"\n最大連號數量分布:")
    consecutive_dist = patterns_df['max_consecutive'].value_counts().sort_index()
    for count, freq in consecutive_dist.items():
        percentage = freq / len(patterns_df) * 100
        print(f"  {count}個: {freq}次 ({percentage:.1f}%)")
    
    # 連號出現情況統計
    has_consecutive_count = patterns_df['has_consecutive'].sum()
    no_consecutive_count = len(patterns_df) - has_consecutive_count
    print(f"\n連號出現情況:")
    print(f"  有出現連號: {has_consecutive_count}期 ({has_consecutive_count/len(patterns_df)*100:.1f}%)")
    print(f"  沒有連號: {no_consecutive_count}期 ({no_consecutive_count/len(patterns_df)*100:.1f}%)")
    
    # 奇數比例統計
    print(f"\n奇數數量分布:")
    odd_dist = patterns_df['odd_count'].value_counts().sort_index()
    for count, freq in odd_dist.items():
        percentage = freq / len(patterns_df) * 100
        print(f"  {count}個奇數: {freq}次 ({percentage:.1f}%)")
    
    # 雙數比例統計
    print(f"\n雙數數量分布:")
    even_dist = patterns_df['even_count'].value_counts().sort_index()
    for count, freq in even_dist.items():
        percentage = freq / len(patterns_df) * 100
        print(f"  {count}個雙數: {freq}次 ({percentage:.1f}%)")
    
    # 大小號比例統計
    print(f"\n小號(1-19)數量分布:")
    small_dist = patterns_df['small_count'].value_counts().sort_index()
    for count, freq in small_dist.items():
        percentage = freq / len(patterns_df) * 100
        print(f"  {count}個: {freq}次 ({percentage:.1f}%)")
    
    # 跨度統計
    print(f"\n號碼跨度統計:")
    print(f"  平均跨度: {patterns_df['span'].mean():.1f}")
    print(f"  最小跨度: {patterns_df['span'].min()}")
    print(f"  最大跨度: {patterns_df['span'].max()}")
    
    # 綜合統計摘要
    print(f"\n📈 統計摘要 (共{len(patterns_df)}期):")
    print(f"  • 0頭號碼最常出現: {zero_head_dist.idxmax()}個 ({zero_head_dist.max()}次)")
    print(f"  • 1頭號碼最常出現: {one_head_dist.idxmax()}個 ({one_head_dist.max()}次)")
    print(f"  • 2頭號碼最常出現: {two_head_dist.idxmax()}個 ({two_head_dist.max()}次)")
    print(f"  • 3頭號碼最常出現: {three_head_dist.idxmax()}個 ({three_head_dist.max()}次)")
    print(f"  • 奇數最常出現: {odd_dist.idxmax()}個 ({odd_dist.max()}次)")
    print(f"  • 雙數最常出現: {even_dist.idxmax()}個 ({even_dist.max()}次)")
    print(f"  • 連號超過2個的機率: {(consecutive_dist[consecutive_dist.index > 2].sum() / len(patterns_df) * 100):.1f}%")
    print(f"  • 有連號出現的機率: {(has_consecutive_count / len(patterns_df) * 100):.1f}%")
    print(f"  • 完全沒有連號的機率: {(no_consecutive_count / len(patterns_df) * 100):.1f}%")
    
    return patterns_df

def is_pattern_reasonable(numbers, historical_stats=None):
    """檢查號碼組合是否符合歷史模式"""
    pattern = analyze_single_draw_patterns(numbers)
    
    # 如果有歷史統計數據，使用動態標準；否則使用固定標準
    if historical_stats is not None:
        # 基於歷史統計的動態驗證
        zero_head_common = historical_stats['zero_head_most_common']
        one_head_common = historical_stats['one_head_most_common']
        two_head_common = historical_stats['two_head_most_common']
        three_head_common = historical_stats['three_head_most_common']
        consecutive_prob = historical_stats['consecutive_prob']
        
        # 允許常見的模式組合
        if pattern['zero_head_count'] > zero_head_common + 1:
            return False
        if pattern['one_head_count'] > one_head_common + 1:
            return False
        if pattern['two_head_count'] > two_head_common + 1:
            return False
        if pattern['three_head_count'] > three_head_common + 1:
            return False
        if pattern['max_consecutive'] > 3:  # 連號仍保持最大3個的限制
            return False
    else:
        # 固定標準（向後兼容）
        if pattern['zero_head_count'] > 3:
            return False
        if pattern['max_consecutive'] > 3:
            return False
    
    # 跨度檢查
    if pattern['span'] < 10 or pattern['span'] > 35:
        return False
    
    return True

def is_trend_reasonable(numbers, historical_stats):
    """基於近期趨勢檢查號碼組合是否合理"""
    if not historical_stats.get('trend_detected', False):
        # 沒有趨勢，使用標準驗證
        return is_pattern_reasonable(numbers, historical_stats)
    
    pattern = analyze_single_draw_patterns(numbers)
    trends = historical_stats.get('trends', {})
    recent_data = historical_stats.get('recent_vs_historical', {})
    
    if not recent_data:
        return is_pattern_reasonable(numbers, historical_stats)
    
    recent_avg = recent_data['recent']
    
    # 根據趨勢調整驗證標準
    # 如果近期某個特徵增加，則更容易接受較高的值
    
    # 0頭號碼趨勢調整
    if 'zero_head_avg_trend' in trends:
        if trends['zero_head_avg_trend'] == 'increasing':
            # 近期0頭號碼增加趨勢，放寬標準
            if pattern['zero_head_count'] > int(recent_avg['zero_head_avg']) + 2:
                return False
        else:
            # 近期0頭號碼減少趨勢，收緊標準
            if pattern['zero_head_count'] > int(recent_avg['zero_head_avg']) + 1:
                return False
    
    # 連號趨勢調整
    if 'consecutive_rate_trend' in trends:
        if trends['consecutive_rate_trend'] == 'decreasing':
            # 近期連號減少，更傾向於選擇無連號的組合
            if pattern['max_consecutive'] > 2:
                return False
    
    # 基本的跨度檢查
    if pattern['span'] < 10 or pattern['span'] > 35:
        return False
    
    return True

def get_historical_stats(patterns_df, recent_periods=None, use_weighted=False):
    """
    從歷史數據中提取統計信息用於選號驗證
    Args:
        patterns_df: 完整的歷史模式數據
        recent_periods: 近期期數，None表示使用全部數據
        use_weighted: 是否使用加權分析（近期權重更高）
    """
    if recent_periods is not None and recent_periods > 0:
        # 使用近期數據
        analysis_df = patterns_df.tail(recent_periods).copy()
        print(f"🎯 使用近期 {len(analysis_df)} 期數據進行分析 (共{len(patterns_df)}期)")
    else:
        analysis_df = patterns_df.copy()
        print(f"📊 使用全部 {len(analysis_df)} 期數據進行分析")
    
    if use_weighted and len(analysis_df) > 10:
        # 加權分析：越近期的數據權重越高
        weights = create_time_weights(len(analysis_df))
        stats = calculate_weighted_stats(analysis_df, weights)
        print("⚖️ 採用時間加權分析（近期權重較高）")
    else:
        # 傳統統計分析
        stats = calculate_standard_stats(analysis_df)
        if use_weighted:
            print("📋 數據量不足，採用標準分析")
    
    # 新增趨勢分析
    trend_analysis = analyze_recent_trends(patterns_df, analysis_df)
    stats.update(trend_analysis)
    
    return stats

def create_time_weights(n_periods):
    """創建時間權重：近期資料權重更高"""
    # 使用指數衰減權重，最新的權重=1，最舊的權重約0.1
    weights = np.exp(np.linspace(-2.3, 0, n_periods))  # e^(-2.3) ≈ 0.1
    return weights / weights.sum()  # 正規化

def calculate_weighted_stats(df, weights):
    """計算加權統計"""
    stats = {}
    
    # 加權頻次計算
    for col in ['zero_head_count', 'one_head_count', 'two_head_count', 'three_head_count']:
        weighted_freq = {}
        for i, value in enumerate(df[col]):
            if value in weighted_freq:
                weighted_freq[value] += weights[i]
            else:
                weighted_freq[value] = weights[i]
        
        # 找出加權頻次最高的值
        most_common = max(weighted_freq.items(), key=lambda x: x[1])[0]
        stats[f'{col.replace("_count", "")}_most_common'] = most_common
    
    # 加權連號機率
    consecutive_weight = sum(weights[i] for i, has_consecutive in enumerate(df['has_consecutive']) if has_consecutive)
    stats['consecutive_prob'] = consecutive_weight
    stats['total_periods'] = len(df)
    
    return stats

def calculate_standard_stats(df):
    """計算標準統計"""
    zero_head_dist = df['zero_head_count'].value_counts()
    one_head_dist = df['one_head_count'].value_counts()
    two_head_dist = df['two_head_count'].value_counts()
    three_head_dist = df['three_head_count'].value_counts()
    
    consecutive_count = df['has_consecutive'].sum()
    total_count = len(df)
    
    return {
        'zero_head_most_common': zero_head_dist.idxmax(),
        'one_head_most_common': one_head_dist.idxmax(),
        'two_head_most_common': two_head_dist.idxmax(),
        'three_head_most_common': three_head_dist.idxmax(),
        'consecutive_prob': consecutive_count / total_count,
        'total_periods': total_count
    }

def analyze_recent_trends(full_df, recent_df):
    """分析近期趨勢變化"""
    if len(full_df) < 50 or len(recent_df) < 10:
        return {'trend_detected': False}
    
    # 比較近期與歷史平均的差異
    full_avg = {
        'zero_head_avg': full_df['zero_head_count'].mean(),
        'one_head_avg': full_df['one_head_count'].mean(),
        'two_head_avg': full_df['two_head_count'].mean(),
        'three_head_avg': full_df['three_head_count'].mean(),
        'consecutive_rate': full_df['has_consecutive'].mean()
    }
    
    recent_avg = {
        'zero_head_avg': recent_df['zero_head_count'].mean(),
        'one_head_avg': recent_df['one_head_count'].mean(),
        'two_head_avg': recent_df['two_head_count'].mean(),
        'three_head_avg': recent_df['three_head_count'].mean(),
        'consecutive_rate': recent_df['has_consecutive'].mean()
    }
    
    # 檢測顯著趨勢變化（差異超過20%）
    trends = {}
    trend_detected = False
    
    for key in full_avg:
        diff_ratio = abs(recent_avg[key] - full_avg[key]) / full_avg[key] if full_avg[key] > 0 else 0
        if diff_ratio > 0.2:  # 20%的變化閾值
            trends[f'{key}_trend'] = 'increasing' if recent_avg[key] > full_avg[key] else 'decreasing'
            trend_detected = True
    
    return {
        'trend_detected': trend_detected,
        'trends': trends,
        'recent_vs_historical': {
            'recent': recent_avg,
            'historical': full_avg
        }
    }

# --------------------------------------------------------
# 5. 改進的建議號碼產生器 ----------------------------------
def suggest_numbers(strategy='smart', n=9, historical_stats=None):
    """
    產生建議號碼
    Args:
        strategy: 選號策略 ('random', 'hot', 'cold', 'smart', 'balanced')
        n: 要選擇的號碼數量（預設9個）
        historical_stats: 歷史統計數據用於驗證
    Returns:
        排序後的n個號碼列表
    """
    numbers = list(range(1, 40))          # 1-39
    global hot_numbers, cold_numbers

    if strategy == 'random':
        # 隨機選號：直接選擇n個號碼
        return sorted(random.sample(numbers, n))

    elif strategy == 'hot':
        # 熱號優先：優先選擇熱號，不足時補充其他號碼
        sel = hot_numbers.copy()
        if len(sel) >= n:
            return sorted(random.sample(sel, n))
        else:
            remain = [x for x in numbers if x not in sel]
            sel += random.sample(remain, n - len(sel))
            return sorted(sel)

    elif strategy == 'cold':
        # 冷號優先：優先選擇冷號，不足時補充其他號碼
        sel = cold_numbers.copy()
        if len(sel) >= n:
            return sorted(random.sample(sel, n))
        else:
            remain = [x for x in numbers if x not in sel]
            sel += random.sample(remain, n - len(sel))
            return sorted(sel)
    
    elif strategy == 'smart':
        # 智能選號：嚴格基於歷史統計模式
        max_attempts = 5000  # 增加嘗試次數以確保找到符合歷史模式的組合
        attempts = 0
        
        while attempts < max_attempts:
            # 先生成一組5個號碼（主要組合）
            main_group = random.sample(numbers, 5)
            
            if is_pattern_reasonable(main_group, historical_stats):
                # 再從剩餘號碼中選4個
                remaining = [x for x in numbers if x not in main_group]
                additional = random.sample(remaining, 4)
                result = sorted(main_group + additional)
                return result
            attempts += 1
        
        # 如果無法找到理想組合，降低標準再試一次
        attempts = 0
        while attempts < 1000:
            main_group = random.sample(numbers, 5)
            if is_pattern_reasonable(main_group):  # 使用固定標準
                remaining = [x for x in numbers if x not in main_group]
                additional = random.sample(remaining, 4)
                result = sorted(main_group + additional)
                return result
            attempts += 1
        
        # 最後退回隨機選號
        return sorted(random.sample(numbers, n))
    
    elif strategy == 'balanced':
        # 平衡策略：結合熱號、冷號，並考慮基本模式
        max_attempts = 1000
        attempts = 0
        
        while attempts < max_attempts:
            # 選擇3-4個熱號
            hot_count = random.randint(3, 4)
            # 選擇2-3個冷號
            cold_count = random.randint(2, 3)
            # 剩餘隨機選擇
            random_count = n - hot_count - cold_count
            
            candidates = []
            
            # 從熱號中選擇
            if len(hot_numbers) >= hot_count:
                candidates.extend(random.sample(hot_numbers, hot_count))
            else:
                candidates.extend(hot_numbers)
                hot_count = len(hot_numbers)
            
            # 從冷號中選擇
            available_cold = [x for x in cold_numbers if x not in candidates]
            actual_cold_count = min(cold_count, len(available_cold))
            if actual_cold_count > 0:
                candidates.extend(random.sample(available_cold, actual_cold_count))
            
            # 隨機選擇剩餘的
            remaining_numbers = [x for x in numbers if x not in candidates]
            random_count = n - len(candidates)
            if random_count > 0 and len(remaining_numbers) >= random_count:
                candidates.extend(random.sample(remaining_numbers, random_count))
            
            if len(candidates) == n:
                return sorted(candidates)
            attempts += 1
        
        # 如果無法找到合理組合，退回隨機選號
        return sorted(random.sample(numbers, n))
    
    elif strategy == 'trend_adaptive':
        # 趨勢適應策略：根據近期趨勢調整選號
        if historical_stats is None or not historical_stats.get('trend_detected', False):
            # 沒有趨勢數據，退回智能選號
            return suggest_numbers('smart', n, historical_stats)
        
        max_attempts = 3000
        attempts = 0
        trends = historical_stats.get('trends', {})
        
        while attempts < max_attempts:
            main_group = random.sample(numbers, 5)
            
            # 根據趨勢調整驗證標準
            if is_trend_reasonable(main_group, historical_stats):
                remaining = [x for x in numbers if x not in main_group]
                additional = random.sample(remaining, 4)
                result = sorted(main_group + additional)
                return result
            attempts += 1
        
        # 趨勢策略失敗，退回智能選號
        return suggest_numbers('smart', n, historical_stats)

    else:
        raise ValueError(f'unknown strategy {strategy}')

# --------------------------------------------------------
# 6. 預測記錄功能 -------------------------------------------
def check_daily_analysis_exists(log_file="prediction_log.xlsx"):
    """
    檢查今日是否已有分析記錄
    Returns:
        bool: True表示今日已有記錄，False表示今日尚無記錄
    """
    current_time = datetime.now()
    date_str = current_time.strftime("%Y-%m-%d")
    
    log_path = Path(log_file)
    if not log_path.exists():
        return False
    
    try:
        existing_df = pd.read_excel(log_file, engine='openpyxl')
        today_records = existing_df[existing_df['日期'] == date_str]
        
        if len(today_records) > 0:
            print(f"📅 檢測到今日({date_str})已有分析記錄")
            # 顯示已有記錄的摘要
            latest_record = today_records.iloc[-1]
            print(f"   記錄時間: {latest_record.get('時間', 'N/A')}")
            if '智能選號' in latest_record and pd.notna(latest_record['智能選號']):
                print(f"   智能選號: {latest_record['智能選號']}")
            if '平衡策略' in latest_record and pd.notna(latest_record['平衡策略']):
                print(f"   平衡策略: {latest_record['平衡策略']}")
            return True
        else:
            return False
    except Exception as e:
        print(f"⚠️ 檢查今日記錄時發生錯誤: {e}")
        return False

def log_predictions_to_excel(predictions, log_file="prediction_log.xlsx", allow_overwrite=False):
    """
    將預測結果記錄到Excel檔案
    Args:
        predictions: 包含各策略預測結果的字典
        log_file: 記錄檔案名稱
        allow_overwrite: 是否允許覆蓋今日已有記錄（預設False不覆蓋）
    """
    current_time = datetime.now()
    date_str = current_time.strftime("%Y-%m-%d")
    time_str = current_time.strftime("%H:%M:%S")
    
    # 檢查今日是否已有記錄
    if not allow_overwrite and check_daily_analysis_exists(log_file):
        print(f"⚠️ 今日({date_str})已有分析記錄，跳過記錄以避免覆蓋")
        print(f"   如需強制覆蓋，請手動刪除今日記錄或使用 allow_overwrite=True 參數")
        return False
    
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
        '驗證結果': '',  # 留空，等待未來開獎結果
        '中獎號碼數': '',  # 留空，等待驗證
        '備註': f"基於歷史統計模式生成"
    }
    
    # 如果有趨勢適應策略，也記錄
    if 'trend_adaptive' in predictions:
        log_data['趨勢適應'] = str(predictions['trend_adaptive'])
        log_data['趨勢適應_前5號'] = str(predictions['trend_adaptive'][:5])
    
    # 檢查檔案是否存在
    log_path = Path(log_file)
    overwrite_occurred = False
    
    if log_path.exists():
        # 讀取現有數據
        try:
            existing_df = pd.read_excel(log_file, engine='openpyxl')
            
            # 檢查今天是否已有記錄
            today_records = existing_df[existing_df['日期'] == date_str]
            
            if len(today_records) > 0 and allow_overwrite:
                # 今天已有記錄且允許覆蓋，覆蓋最新的一筆
                latest_today_index = today_records.index[-1]  # 取得今天最後一筆的索引
                
                # 保留已驗證的結果（如果有的話）
                old_record = existing_df.loc[latest_today_index]
                if pd.notna(old_record.get('驗證結果', '')) and old_record.get('驗證結果', '') != '':
                    log_data['驗證結果'] = old_record['驗證結果']
                    log_data['中獎號碼數'] = old_record['中獎號碼數']
                    print(f"🔄 覆蓋今日記錄，保留已驗證結果")
                else:
                    print(f"🔄 覆蓋今日記錄")
                
                # 覆蓋該記錄
                for key, value in log_data.items():
                    if key in existing_df.columns:
                        existing_df.loc[latest_today_index, key] = value
                    else:
                        # 如果是新欄位（如趨勢適應），需要先添加欄位
                        existing_df[key] = ''
                        existing_df.loc[latest_today_index, key] = value
                
                combined_df = existing_df
                overwrite_occurred = True
            else:
                # 今天沒有記錄或不允許覆蓋，新增
                new_df = pd.DataFrame([log_data])
                
                # 檢查是否有新欄位需要添加到現有數據
                for col in new_df.columns:
                    if col not in existing_df.columns:
                        existing_df[col] = ''
                
                combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                
        except Exception as e:
            print(f"讀取現有記錄檔案時發生錯誤: {e}")
            # 如果讀取失敗，創建新的DataFrame
            combined_df = pd.DataFrame([log_data])
    else:
        # 創建新的DataFrame
        combined_df = pd.DataFrame([log_data])
    
    # 寫入Excel
    try:
        combined_df.to_excel(log_file, index=False, engine='openpyxl')
        
        if overwrite_occurred:
            print(f"\n🔄 預測記錄已覆蓋更新: {log_file}")
            print(f"   更新時間: {date_str} {time_str}")
            print(f"   ⚠️  同一天的舊記錄已被替換")
        else:
            print(f"\n✅ 預測記錄已保存到: {log_file}")
            print(f"   記錄時間: {date_str} {time_str}")
        
        return True
            
    except Exception as e:
        print(f"❌ 保存預測記錄時發生錯誤: {e}")
        return False

def analyze_prediction_accuracy(log_file="prediction_log.xlsx", actual_results_file="lottery_hist.xlsx"):
    """
    分析預測準確度（未來功能，當有足夠歷史預測記錄時使用）
    """
    if not Path(log_file).exists():
        print("尚無預測記錄檔案")
        return
    
    try:
        predictions_df = pd.read_excel(log_file, engine='openpyxl')
        print(f"\n📊 預測記錄統計:")
        print(f"   總預測次數: {len(predictions_df)}")
        print(f"   最早預測: {predictions_df['日期'].min()}")
        print(f"   最新預測: {predictions_df['日期'].max()}")
        print(f"   ⚠️  驗證功能將在累積足夠數據後實作")
    except Exception as e:
        print(f"讀取預測記錄時發生錯誤: {e}")

# --------------------------------------------------------
# 5. 期望獲利計算 -------------------------------------------
def expected_profit(n: int = 9,
                    cost_per_ticket: int = 1400,
                    payout_per_hit: int = 11000):
    prob_k = n * 5 / 39
    exp_payout = payout_per_hit * prob_k
    cost = cost_per_ticket * n
    return exp_payout - cost

# --------------------------------------------------------
# 6. 主程式 ------------------------------------------------------
if __name__ == "__main__":

    # 路徑（請自行改成實際檔名）
    excel_file = "lottery_hist.xlsx"

    if not Path(excel_file).exists():
        raise FileNotFoundError(f"找不到檔案: {excel_file}")

    # ---------- 自動驗證之前的預測 ----------
    if VERIFICATION_AVAILABLE:
        auto_verify_on_startup("prediction_log.xlsx", excel_file)
    
    print("\n" + "="*60)
    print("🎲 539彩票分析系統")
    print("="*60)
    
    # ---------- 檢查今日是否已有分析記錄 ----------
    if check_daily_analysis_exists("prediction_log.xlsx"):
        print("\n✅ 今日分析已完成，程式結束")
        print("💡 如需重新分析，請手動刪除今日記錄或修改程式碼")
        exit(0)

    df = load_lottery_excel(excel_file)

    # ---------- 基本統計 ----------
    freq_series = compute_num_frequency(df)
    hot_numbers, cold_numbers = get_hot_cold_numbers(freq_series, top_n=6, bottom_n=6)

    print("\n=== 最高頻次號碼（熱號） ===")
    print(hot_numbers)
    print("\n=== 最低頻次號碼（冷號） ===")
    print(cold_numbers)

    # ---------- 模式分析 ----------
    print("\n" + "="*50)
    print("開始進行歷史開獎模式分析...")
    patterns_df = pattern_statistics(df)
    
    # 提取歷史統計數據用於智能選號
    print(f"\n🔬 統計分析模式選擇:")
    
    # 根據數據量決定分析策略
    total_periods = len(patterns_df)
    
    if total_periods >= 200:
        # 數據充足，提供多種分析選項
        print(f"📊 檢測到充足的歷史數據 ({total_periods} 期)")
        
        # 1. 全部數據分析
        historical_stats_full = get_historical_stats(patterns_df)
        
        # 2. 近期數據分析（最近100期）
        recent_periods = min(100, total_periods // 2)
        historical_stats_recent = get_historical_stats(patterns_df, recent_periods=recent_periods)
        
        # 3. 加權分析（近期權重更高）
        historical_stats_weighted = get_historical_stats(patterns_df, recent_periods=recent_periods, use_weighted=True)
        
        # 選擇主要分析模式
        if historical_stats_recent.get('trend_detected', False):
            print(f"📈 檢測到近期趨勢變化，採用近期加權分析")
            historical_stats = historical_stats_weighted
            analysis_mode = "近期加權"
        else:
            print(f"📋 近期模式穩定，採用近期數據分析")
            historical_stats = historical_stats_recent
            analysis_mode = "近期數據"
    
    elif total_periods >= 50:
        # 中等數據量，使用近期分析
        recent_periods = min(30, total_periods // 2)
        historical_stats = get_historical_stats(patterns_df, recent_periods=recent_periods)
        analysis_mode = "近期數據"
        print(f"📊 使用近期 {recent_periods} 期數據進行分析")
    
    else:
        # 數據不足，使用全部數據
        historical_stats = get_historical_stats(patterns_df)
        analysis_mode = "全部數據"
        print(f"📊 數據量較少，使用全部 {total_periods} 期數據")
    
    # 顯示統計基準
    print(f"\n🎯 統計基準 ({analysis_mode}):")
    print(f"   最常見的0頭號碼數量: {historical_stats['zero_head_most_common']}個")
    print(f"   最常見的1頭號碼數量: {historical_stats['one_head_most_common']}個")
    print(f"   最常見的2頭號碼數量: {historical_stats['two_head_most_common']}個")
    print(f"   最常見的3頭號碼數量: {historical_stats['three_head_most_common']}個")
    print(f"   連號出現機率: {historical_stats['consecutive_prob']:.2%}")
    
    # 顯示趨勢分析結果
    if historical_stats.get('trend_detected', False):
        print(f"\n📈 近期趨勢分析:")
        trends = historical_stats.get('trends', {})
        for trend_key, trend_direction in trends.items():
            trend_name = trend_key.replace('_trend', '').replace('_', ' ')
            direction_text = "增加" if trend_direction == 'increasing' else "減少"
            print(f"   {trend_name}: {direction_text}趨勢")
        
        recent_data = historical_stats.get('recent_vs_historical', {})
        if recent_data:
            print(f"\n📊 近期 vs 歷史比較:")
            for key in ['zero_head_avg', 'consecutive_rate']:
                if key in recent_data['recent']:
                    recent_val = recent_data['recent'][key]
                    hist_val = recent_data['historical'][key]
                    change = ((recent_val - hist_val) / hist_val * 100) if hist_val > 0 else 0
                    key_name = key.replace('_avg', '').replace('_', ' ')
                    print(f"   {key_name}: {recent_val:.2f} vs {hist_val:.2f} ({change:+.1f}%)")

    # ------------- 可視化已移除 ----------
    # plot_frequency(freq_series)  # 已移除可視化功能

    # ------------- 基於歷史統計的建議號碼 ----------
    print("\n" + "="*50)
    print("根據歷史模式生成9個建議號碼...")
    
    # 生成各策略的建議號碼
    print("\n===== 智能選號 (9個號碼) =====")
    smart_numbers = suggest_numbers('smart', n=9, historical_stats=historical_stats)
    print(f"建議號碼: {smart_numbers}")
    # 分析前5個號碼的模式作為參考
    first_five = smart_numbers[:5]
    pattern = analyze_single_draw_patterns(first_five)
    print(f"前5號模式: 0頭:{pattern['zero_head_count']}個, 1頭:{pattern['one_head_count']}個, 2頭:{pattern['two_head_count']}個, 3頭:{pattern['three_head_count']}個")
    print(f"          連號:{pattern['max_consecutive']}個, 奇數:{pattern['odd_count']}個, 雙數:{pattern['even_count']}個, 跨度:{pattern['span']}")
    
    print("\n===== 平衡策略 (9個號碼) =====")
    balanced_numbers = suggest_numbers('balanced', n=9, historical_stats=historical_stats)
    print(f"建議號碼: {balanced_numbers}")
    first_five = balanced_numbers[:5]
    pattern = analyze_single_draw_patterns(first_five)
    print(f"前5號模式: 0頭:{pattern['zero_head_count']}個, 1頭:{pattern['one_head_count']}個, 2頭:{pattern['two_head_count']}個, 3頭:{pattern['three_head_count']}個")
    print(f"          連號:{pattern['max_consecutive']}個, 奇數:{pattern['odd_count']}個, 雙數:{pattern['even_count']}個, 跨度:{pattern['span']}")

    print("\n===== 隨機選號 (9個號碼) =====")
    random_numbers = suggest_numbers('random', n=9)
    print(f"建議號碼: {random_numbers}")
    first_five = random_numbers[:5]
    pattern = analyze_single_draw_patterns(first_five)
    print(f"前5號模式: 0頭:{pattern['zero_head_count']}個, 1頭:{pattern['one_head_count']}個, 2頭:{pattern['two_head_count']}個, 3頭:{pattern['three_head_count']}個")
    print(f"          連號:{pattern['max_consecutive']}個, 奇數:{pattern['odd_count']}個, 雙數:{pattern['even_count']}個, 跨度:{pattern['span']}")

    print("\n===== 熱號優先 (9個號碼) =====")
    hot_numbers_selected = suggest_numbers('hot', n=9)
    print(f"建議號碼: {hot_numbers_selected}")
    first_five = hot_numbers_selected[:5]
    pattern = analyze_single_draw_patterns(first_five)
    print(f"前5號模式: 0頭:{pattern['zero_head_count']}個, 1頭:{pattern['one_head_count']}個, 2頭:{pattern['two_head_count']}個, 3頭:{pattern['three_head_count']}個")
    print(f"          連號:{pattern['max_consecutive']}個, 奇數:{pattern['odd_count']}個, 雙數:{pattern['even_count']}個, 跨度:{pattern['span']}")

    print("\n===== 冷號優先 (9個號碼) =====")
    cold_numbers_selected = suggest_numbers('cold', n=9)
    print(f"建議號碼: {cold_numbers_selected}")
    first_five = cold_numbers_selected[:5]
    pattern = analyze_single_draw_patterns(first_five)
    print(f"前5號模式: 0頭:{pattern['zero_head_count']}個, 1頭:{pattern['one_head_count']}個, 2頭:{pattern['two_head_count']}個, 3頭:{pattern['three_head_count']}個")
    print(f"          連號:{pattern['max_consecutive']}個, 奇數:{pattern['odd_count']}個, 雙數:{pattern['even_count']}個, 跨度:{pattern['span']}")

    # 趨勢適應策略（如果檢測到趨勢）
    trend_numbers = None
    if historical_stats.get('trend_detected', False):
        print("\n===== 趨勢適應 (9個號碼) =====")
        trend_numbers = suggest_numbers('trend_adaptive', n=9, historical_stats=historical_stats)
        print(f"建議號碼: {trend_numbers}")
        first_five = trend_numbers[:5]
        pattern = analyze_single_draw_patterns(first_five)
        print(f"前5號模式: 0頭:{pattern['zero_head_count']}個, 1頭:{pattern['one_head_count']}個, 2頭:{pattern['two_head_count']}個, 3頭:{pattern['three_head_count']}個")
        print(f"          連號:{pattern['max_consecutive']}個, 奇數:{pattern['odd_count']}個, 雙數:{pattern['even_count']}個, 跨度:{pattern['span']}")
        print(f"💡 此策略根據檢測到的近期趨勢進行選號調整")
    else:
        print("\n📋 未檢測到顯著趨勢，跳過趨勢適應策略")
    
    # ------------- 記錄預測結果 ----------
    predictions = {
        'smart': smart_numbers,
        'balanced': balanced_numbers,
        'random': random_numbers,
        'hot': hot_numbers_selected,
        'cold': cold_numbers_selected
    }
    
    # 如果有趨勢策略結果，也記錄
    if trend_numbers is not None:
        predictions['trend_adaptive'] = trend_numbers
    
    print("\n" + "="*50)
    print("正在記錄預測結果...")
    log_predictions_to_excel(predictions, "prediction_log.xlsx")
    
    # 顯示預測記錄統計
    analyze_prediction_accuracy("prediction_log.xlsx")

    # ------------- 期望獲利 ----------
    exp = expected_profit()
    print(f"\n期望獲利（固定 9 個號碼）:  {exp:.2f} 台幣")

    # ------------- 分析總結 ----------
    print("\n" + "="*50)
    print("📊 分析總結:")
    print("1. 智能選號和平衡策略會自動避免歷史上很少出現的模式")
    print("2. 請參考上方的模式統計來了解合理的號碼組合特徵")
    print("3. 建議觀察不同策略生成的號碼模式差異")
    
    # ------------- 警示 ----------
    print("\n⚠️ 重要提醒:")
    print("• 任何選號策略都無法改變中獎機率")
    print("• 歷史模式分析僅供參考，不代表未來趨勢")
    print("• 理性投注，量力而為")
