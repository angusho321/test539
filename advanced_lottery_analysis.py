#!/usr/bin/env python3
"""
進階彩票分析系統
包含時間加權分析、未開出組合檢查、歷史板路相似性預測
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from itertools import combinations
from collections import defaultdict, Counter
import math
import json
from pathlib import Path

class AdvancedLotteryAnalyzer:
    def __init__(self, excel_file="lottery_hist.xlsx"):
        """初始化進階分析器"""
        self.excel_file = excel_file
        self.df = None
        self.load_data()
    
    def load_data(self):
        """載入歷史資料"""
        try:
            self.df = pd.read_excel(self.excel_file, engine='openpyxl')
            self.df['日期'] = pd.to_datetime(self.df['日期'])
            self.df = self.df.sort_values('日期').reset_index(drop=True)
            print(f"✅ 載入 {len(self.df)} 筆歷史記錄")
            print(f"📅 日期範圍: {self.df['日期'].min().date()} ~ {self.df['日期'].max().date()}")
        except Exception as e:
            print(f"❌ 載入資料失敗: {e}")
            self.df = None
    
    def compute_weighted_frequency(self, decay_factor=0.95, recent_days=365):
        """
        計算時間加權的號碼頻率
        
        Args:
            decay_factor: 衰減係數 (0-1)，越接近1衰減越慢
            recent_days: 只考慮最近N天的記錄
        
        Returns:
            dict: {號碼: 加權頻率}
        """
        if self.df is None:
            return {}
        
        print(f"\n⚖️ 計算時間加權頻率 (衰減係數: {decay_factor})")
        
        # 只取最近的記錄
        cutoff_date = datetime.now() - timedelta(days=recent_days)
        recent_df = self.df[self.df['日期'] >= cutoff_date].copy()
        
        if len(recent_df) == 0:
            print("⚠️ 沒有足夠的近期記錄")
            return {}
        
        print(f"📊 分析最近 {len(recent_df)} 筆記錄 (近 {recent_days} 天)")
        
        # 計算每筆記錄距今的天數
        today = datetime.now()
        recent_df['days_ago'] = (today - recent_df['日期']).dt.days
        
        # 計算加權頻率
        weighted_freq = defaultdict(float)
        
        for _, row in recent_df.iterrows():
            # 計算權重：越近期權重越高
            weight = decay_factor ** row['days_ago']
            
            # 累加每個號碼的加權頻率
            for col in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']:
                if pd.notna(row[col]):
                    number = int(row[col])
                    weighted_freq[number] += weight
        
        # 正規化頻率
        total_weight = sum(weighted_freq.values())
        if total_weight > 0:
            for num in weighted_freq:
                weighted_freq[num] = weighted_freq[num] / total_weight
        
        # 排序並顯示結果
        sorted_freq = sorted(weighted_freq.items(), key=lambda x: x[1], reverse=True)
        
        print("🔥 加權熱號 (前10):")
        for i, (num, freq) in enumerate(sorted_freq[:10]):
            print(f"   {i+1:2d}. 號碼 {num:2d}: {freq:.4f}")
        
        print("❄️ 加權冷號 (後10):")
        for i, (num, freq) in enumerate(sorted_freq[-10:]):
            print(f"   {i+1:2d}. 號碼 {num:2d}: {freq:.4f}")
        
        return dict(weighted_freq)
    
    def find_never_drawn_combinations(self, combo_size=5, sample_size=1000):
        """
        找出從未開出的號碼組合
        
        Args:
            combo_size: 組合大小 (預設5個號碼)
            sample_size: 抽樣檢查的組合數量
        
        Returns:
            list: 從未開出的組合列表
        """
        if self.df is None:
            return []
        
        print(f"\n🔍 分析從未開出的 {combo_size} 號組合...")
        
        # 提取所有歷史開獎組合
        historical_combinations = set()
        
        for _, row in self.df.iterrows():
            numbers = []
            for col in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']:
                if pd.notna(row[col]):
                    numbers.append(int(row[col]))
            
            if len(numbers) == 5:
                # 轉換為排序的 tuple
                combo = tuple(sorted(numbers))
                historical_combinations.add(combo)
        
        print(f"📊 歷史開出組合數: {len(historical_combinations):,}")
        
        # 計算總可能組合數
        total_combinations = len(list(combinations(range(1, 40), combo_size)))
        never_drawn_count = total_combinations - len(historical_combinations)
        
        print(f"🎲 總可能組合數: {total_combinations:,}")
        print(f"💎 從未開出組合數: {never_drawn_count:,} ({never_drawn_count/total_combinations*100:.2f}%)")
        
        # 隨機抽樣檢查從未開出的組合
        never_drawn_samples = []
        checked_count = 0
        
        print(f"\n🎯 隨機抽樣檢查 {sample_size} 個組合...")
        
        # 生成隨機組合並檢查
        while len(never_drawn_samples) < sample_size and checked_count < sample_size * 10:
            # 隨機生成一個組合
            random_combo = tuple(sorted(np.random.choice(range(1, 40), combo_size, replace=False)))
            checked_count += 1
            
            if random_combo not in historical_combinations:
                never_drawn_samples.append(random_combo)
        
        print(f"✅ 找到 {len(never_drawn_samples)} 個從未開出的組合")
        
        # 顯示前10個
        print("\n💎 從未開出組合範例 (前10個):")
        for i, combo in enumerate(never_drawn_samples[:10]):
            print(f"   {i+1:2d}. {list(combo)}")
        
        return never_drawn_samples
    
    def analyze_recent_patterns(self, pattern_length=5):
        """
        分析最近的開獎模式
        
        Args:
            pattern_length: 模式長度（分析最近N期）
        
        Returns:
            dict: 最近模式的特徵
        """
        if self.df is None or len(self.df) < pattern_length:
            return {}
        
        print(f"\n📈 分析最近 {pattern_length} 期的開獎模式...")
        
        # 取最近的記錄
        recent_records = self.df.tail(pattern_length)
        
        pattern_features = {
            'dates': [],
            'numbers_sequence': [],
            'sum_trend': [],
            'even_odd_ratio': [],
            'high_low_ratio': [],
            'consecutive_pairs': [],
            'number_gaps': []
        }
        
        for _, row in recent_records.iterrows():
            numbers = []
            for col in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']:
                if pd.notna(row[col]):
                    numbers.append(int(row[col]))
            
            if len(numbers) == 5:
                numbers = sorted(numbers)
                pattern_features['dates'].append(row['日期'])
                pattern_features['numbers_sequence'].append(numbers)
                
                # 計算各種特徵
                num_sum = sum(numbers)
                pattern_features['sum_trend'].append(num_sum)
                
                # 奇偶比例
                even_count = sum(1 for n in numbers if n % 2 == 0)
                pattern_features['even_odd_ratio'].append(even_count / 5)
                
                # 高低比例 (20以上為高)
                high_count = sum(1 for n in numbers if n > 20)
                pattern_features['high_low_ratio'].append(high_count / 5)
                
                # 連續號碼對
                consecutive = sum(1 for i in range(len(numbers)-1) if numbers[i+1] - numbers[i] == 1)
                pattern_features['consecutive_pairs'].append(consecutive)
                
                # 號碼間距
                gaps = [numbers[i+1] - numbers[i] for i in range(len(numbers)-1)]
                pattern_features['number_gaps'].append(gaps)
        
        # 顯示模式分析
        print("\n📋 最近模式特徵:")
        for i, (date, numbers) in enumerate(zip(pattern_features['dates'], pattern_features['numbers_sequence'])):
            print(f"   {date.date()}: {numbers} (和:{pattern_features['sum_trend'][i]})")
        
        # 計算趨勢
        if len(pattern_features['sum_trend']) >= 3:
            recent_avg = np.mean(pattern_features['sum_trend'][-3:])
            overall_avg = np.mean(pattern_features['sum_trend'])
            print(f"\n📊 數字和趨勢:")
            print(f"   最近3期平均: {recent_avg:.1f}")
            print(f"   整體平均: {overall_avg:.1f}")
            print(f"   趨勢: {'📈 偏高' if recent_avg > overall_avg else '📉 偏低'}")
        
        return pattern_features
    
    def find_similar_historical_patterns(self, current_pattern, similarity_threshold=0.8):
        """
        在歷史中尋找相似的板路模式
        
        Args:
            current_pattern: 當前模式特徵
            similarity_threshold: 相似度閾值
        
        Returns:
            list: 相似的歷史模式及其後續開獎
        """
        if self.df is None or not current_pattern:
            return []
        
        print(f"\n🧠 尋找歷史相似板路 (相似度閾值: {similarity_threshold})")
        
        pattern_length = len(current_pattern['numbers_sequence'])
        if pattern_length == 0:
            return []
        
        similar_patterns = []
        
        # 在歷史中滑動窗口尋找相似模式
        for i in range(len(self.df) - pattern_length):
            # 提取歷史模式
            historical_pattern = self.extract_pattern_features(i, pattern_length)
            
            if historical_pattern:
                # 計算相似度
                similarity = self.calculate_pattern_similarity(current_pattern, historical_pattern)
                
                if similarity >= similarity_threshold:
                    # 找到相似模式，記錄其後續開獎
                    next_index = i + pattern_length
                    if next_index < len(self.df):
                        next_record = self.df.iloc[next_index]
                        next_numbers = []
                        for col in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']:
                            if pd.notna(next_record[col]):
                                next_numbers.append(int(next_record[col]))
                        
                        if len(next_numbers) == 5:
                            similar_patterns.append({
                                'start_date': self.df.iloc[i]['日期'],
                                'end_date': self.df.iloc[i + pattern_length - 1]['日期'],
                                'similarity': similarity,
                                'pattern': historical_pattern['numbers_sequence'],
                                'next_draw': sorted(next_numbers),
                                'next_date': next_record['日期']
                            })
        
        # 排序相似模式
        similar_patterns.sort(key=lambda x: x['similarity'], reverse=True)
        
        print(f"🔍 找到 {len(similar_patterns)} 個相似的歷史模式")
        
        # 顯示前5個最相似的模式
        print("\n🎯 最相似的歷史模式:")
        for i, pattern in enumerate(similar_patterns[:5]):
            print(f"   {i+1}. 相似度: {pattern['similarity']:.3f}")
            print(f"      時期: {pattern['start_date'].date()} ~ {pattern['end_date'].date()}")
            print(f"      後續開獎: {pattern['next_draw']} ({pattern['next_date'].date()})")
        
        return similar_patterns
    
    def extract_pattern_features(self, start_index, length):
        """提取指定位置的模式特徵"""
        try:
            pattern_records = self.df.iloc[start_index:start_index + length]
            
            features = {
                'numbers_sequence': [],
                'sum_trend': [],
                'even_odd_ratio': [],
                'high_low_ratio': []
            }
            
            for _, row in pattern_records.iterrows():
                numbers = []
                for col in ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']:
                    if pd.notna(row[col]):
                        numbers.append(int(row[col]))
                
                if len(numbers) == 5:
                    numbers = sorted(numbers)
                    features['numbers_sequence'].append(numbers)
                    features['sum_trend'].append(sum(numbers))
                    
                    even_count = sum(1 for n in numbers if n % 2 == 0)
                    features['even_odd_ratio'].append(even_count / 5)
                    
                    high_count = sum(1 for n in numbers if n > 20)
                    features['high_low_ratio'].append(high_count / 5)
            
            return features if len(features['numbers_sequence']) == length else None
            
        except Exception as e:
            return None
    
    def calculate_pattern_similarity(self, pattern1, pattern2):
        """計算兩個模式的相似度"""
        try:
            # 比較數字和趨勢
            sum_similarity = 1 - abs(np.mean(pattern1['sum_trend']) - np.mean(pattern2['sum_trend'])) / 200
            
            # 比較奇偶比例趨勢
            even_odd_similarity = 1 - abs(np.mean(pattern1['even_odd_ratio']) - np.mean(pattern2['even_odd_ratio']))
            
            # 比較高低比例趨勢
            high_low_similarity = 1 - abs(np.mean(pattern1['high_low_ratio']) - np.mean(pattern2['high_low_ratio']))
            
            # 綜合相似度
            total_similarity = (sum_similarity + even_odd_similarity + high_low_similarity) / 3
            
            return max(0, total_similarity)
            
        except Exception as e:
            return 0
    
    def weighted_smart_prediction(self, n_predictions=9, decay_factor=0.95):
        """
        基於時間加權的智能選號
        
        Args:
            n_predictions: 預測號碼數量
            decay_factor: 時間衰減係數
        
        Returns:
            list: 預測號碼
        """
        print(f"\n🎯 時間加權智能選號 (預測 {n_predictions} 個號碼)")
        
        # 計算加權頻率
        weighted_freq = self.compute_weighted_frequency(decay_factor)
        
        if not weighted_freq:
            print("❌ 無法計算加權頻率")
            return []
        
        # 轉換為可用於選號的權重
        numbers = list(range(1, 40))
        weights = [weighted_freq.get(num, 0.001) for num in numbers]  # 給未出現號碼小權重
        
        # 加入一些隨機性，避免過度依賴歷史
        randomness_factor = 0.3
        for i in range(len(weights)):
            weights[i] = weights[i] * (1 - randomness_factor) + np.random.random() * randomness_factor
        
        # 正規化權重
        weights = np.array(weights)
        weights = weights / weights.sum()
        
        # 根據權重選號
        selected_numbers = np.random.choice(numbers, size=n_predictions, replace=False, p=weights)
        result = sorted(selected_numbers.tolist())
        
        print(f"✅ 加權智能選號結果: {result}")
        
        return result
    
    def pattern_based_prediction(self, pattern_length=5):
        """
        基於歷史板路相似性的預測
        
        Args:
            pattern_length: 模式長度
        
        Returns:
            dict: 預測結果和信心度
        """
        print(f"\n🧠 基於板路相似性預測 (模式長度: {pattern_length})")
        
        # 分析最近模式
        current_pattern = self.analyze_recent_patterns(pattern_length)
        
        if not current_pattern:
            print("❌ 無法分析當前模式")
            return {}
        
        # 尋找相似歷史模式
        similar_patterns = self.find_similar_historical_patterns(current_pattern)
        
        if not similar_patterns:
            print("❌ 沒有找到相似的歷史模式")
            return {}
        
        # 統計相似模式的後續開獎
        next_number_freq = defaultdict(int)
        total_patterns = len(similar_patterns)
        
        for pattern in similar_patterns:
            for number in pattern['next_draw']:
                # 根據相似度加權
                weight = pattern['similarity']
                next_number_freq[number] += weight
        
        # 正規化並排序
        total_weight = sum(next_number_freq.values())
        if total_weight > 0:
            for num in next_number_freq:
                next_number_freq[num] = next_number_freq[num] / total_weight
        
        sorted_predictions = sorted(next_number_freq.items(), key=lambda x: x[1], reverse=True)
        
        # 選擇前9個最有可能的號碼
        predicted_numbers = [num for num, freq in sorted_predictions[:9]]
        
        # 計算整體信心度
        confidence = np.mean([p['similarity'] for p in similar_patterns[:5]])
        
        print(f"\n🎯 板路預測結果:")
        print(f"   預測號碼: {sorted(predicted_numbers)}")
        print(f"   信心度: {confidence:.3f}")
        print(f"   基於 {total_patterns} 個相似歷史模式")
        
        return {
            'predicted_numbers': sorted(predicted_numbers),
            'confidence': confidence,
            'similar_patterns_count': total_patterns,
            'top_probabilities': sorted_predictions[:15]
        }
    
    def comprehensive_analysis(self):
        """綜合分析報告"""
        print("\n" + "="*60)
        print("🎲 539 彩票進階分析報告")
        print("="*60)
        
        if self.df is None:
            print("❌ 無法載入歷史資料")
            return
        
        # 1. 時間加權分析
        weighted_prediction = self.weighted_smart_prediction()
        
        # 2. 從未開出組合分析
        never_drawn = self.find_never_drawn_combinations(sample_size=500)
        
        # 3. 板路相似性預測
        pattern_prediction = self.pattern_based_prediction()
        
        # 4. 綜合建議
        print("\n" + "="*60)
        print("🎯 綜合預測建議")
        print("="*60)
        
        print(f"⚖️ 時間加權選號: {weighted_prediction}")
        
        if pattern_prediction:
            print(f"🧠 板路相似預測: {pattern_prediction['predicted_numbers']}")
            print(f"   信心度: {pattern_prediction['confidence']:.3f}")
        
        if never_drawn:
            # 從未開出組合中隨機選一個
            random_never_drawn = never_drawn[np.random.randint(0, min(len(never_drawn), 10))]
            print(f"💎 從未開出組合: {list(random_never_drawn)}")
        
        # 生成最終建議
        final_recommendation = self.generate_final_recommendation(
            weighted_prediction, 
            pattern_prediction.get('predicted_numbers', []),
            never_drawn
        )
        
        print(f"\n🏆 最終建議號碼: {final_recommendation}")
        
        return {
            'weighted_prediction': weighted_prediction,
            'pattern_prediction': pattern_prediction,
            'never_drawn_samples': never_drawn[:10],
            'final_recommendation': final_recommendation
        }
    
    def generate_final_recommendation(self, weighted_nums, pattern_nums, never_drawn):
        """生成最終建議號碼"""
        # 結合各種策略
        all_candidates = set(weighted_nums + pattern_nums)
        
        # 如果候選號碼不足，從未開出組合中補充
        if len(all_candidates) < 9 and never_drawn:
            random_never = never_drawn[0]  # 取第一個從未開出的組合
            all_candidates.update(random_never)
        
        # 如果還是不足，隨機補充
        if len(all_candidates) < 9:
            remaining = [i for i in range(1, 40) if i not in all_candidates]
            all_candidates.update(np.random.choice(remaining, 9 - len(all_candidates), replace=False))
        
        # 取前9個
        final_nums = sorted(list(all_candidates)[:9])
        
        return final_nums

def main():
    """主要執行函數"""
    print("🚀 進階彩票分析系統啟動")
    
    analyzer = AdvancedLotteryAnalyzer()
    
    if analyzer.df is None:
        print("❌ 無法載入歷史資料")
        return
    
    # 執行綜合分析
    results = analyzer.comprehensive_analysis()
    
    # 保存分析結果
    if results:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        results_file = f"advanced_analysis_{timestamp}.json"
        
        # 轉換資料類型以便 JSON 序列化
        def convert_for_json(obj):
            if isinstance(obj, (np.integer, np.int64)):
                return int(obj)
            elif isinstance(obj, (np.floating, np.float64)):
                return float(obj)
            elif isinstance(obj, list):
                return [convert_for_json(item) for item in obj]
            elif isinstance(obj, dict):
                return {k: convert_for_json(v) for k, v in obj.items()}
            else:
                return obj
        
        json_results = convert_for_json(results)
        
        try:
            with open(results_file, 'w', encoding='utf-8') as f:
                json.dump(json_results, f, ensure_ascii=False, indent=2)
            print(f"\n💾 分析結果已保存至: {results_file}")
        except Exception as e:
            print(f"⚠️ 保存分析結果失敗: {e}")

if __name__ == "__main__":
    main()
