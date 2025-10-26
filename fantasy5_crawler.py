#!/usr/bin/env python3
"""
加州Fantasy 5開獎號碼爬蟲
格式與539歷史紀錄保持一致
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import re
import time
import logging
from typing import List, Dict, Optional

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class CaliforniaFantasy5Crawler:
    def __init__(self):
        self.base_url = "https://www.calottery.com"
        self.fantasy5_url = "https://www.calottery.com/fantasy-5"
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
        })
        
    def get_weekday_chinese(self, date_obj):
        """將日期轉換為中文星期"""
        weekdays = ['一', '二', '三', '四', '五', '六', '日']
        return weekdays[date_obj.weekday()]
    
    def parse_date(self, date_str):
        """解析日期字串，支援多種格式"""
        date_formats = [
            '%Y-%m-%d',
            '%m/%d/%Y',
            '%m-%d-%Y',
            '%B %d, %Y',
            '%b %d, %Y'
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except ValueError:
                continue
        
        logger.warning(f"無法解析日期格式: {date_str}")
        return None
    
    def crawl_calottery_results(self):
        """從加州彩票官方網站爬取Fantasy 5開獎結果"""
        try:
            logger.info("🕷️ 開始爬取加州Fantasy 5開獎結果...")
            
            response = self.session.get(self.fantasy5_url, timeout=30)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            results = []
            
            # 尋找開獎結果表格或列表
            # 加州彩票網站結構可能不同，需要根據實際HTML調整
            result_elements = soup.find_all(['div', 'table', 'ul'], class_=re.compile(r'result|draw|winning|number', re.I))
            
            for element in result_elements:
                # 嘗試提取日期和號碼
                date_match = re.search(r'(\d{1,2}[/-]\d{1,2}[/-]\d{4}|\d{4}[/-]\d{1,2}[/-]\d{1,2})', str(element))
                numbers_match = re.findall(r'\b(\d{1,2})\b', str(element))
                
                if date_match and len(numbers_match) >= 5:
                    date_str = date_match.group(1)
                    parsed_date = self.parse_date(date_str)
                    
                    if parsed_date:
                        # 取前5個數字作為開獎號碼
                        winning_numbers = [int(num) for num in numbers_match[:5]]
                        
                        if all(1 <= num <= 39 for num in winning_numbers):
                            results.append({
                                'date': parsed_date,
                                'numbers': winning_numbers
                            })
                            logger.info(f"✅ 找到開獎結果: {parsed_date.strftime('%Y-%m-%d')} -> {winning_numbers}")
            
            if not results:
                logger.warning("⚠️ 未找到開獎結果，嘗試備用方法...")
                return self.crawl_alternative_source()
            
            return results
            
        except Exception as e:
            logger.error(f"❌ 爬取加州彩票網站失敗: {e}")
            return self.crawl_alternative_source()
    
    def crawl_alternative_source(self):
        """備用數據來源"""
        try:
            logger.info("🔄 嘗試備用數據來源...")
            
            # 嘗試第三方網站
            alternative_urls = [
                "https://www.lotteryusa.com/california/fantasy-5/",
                "https://www.lotterypost.com/game/california-fantasy-5",
            ]
            
            for url in alternative_urls:
                try:
                    response = self.session.get(url, timeout=30)
                    response.raise_for_status()
                    
                    soup = BeautifulSoup(response.content, 'html.parser')
                    results = self.parse_alternative_results(soup)
                    
                    if results:
                        logger.info(f"✅ 從備用來源獲取 {len(results)} 筆結果")
                        return results
                        
                except Exception as e:
                    logger.warning(f"⚠️ 備用來源 {url} 失敗: {e}")
                    continue
            
            # 如果所有來源都失敗，返回空結果
            logger.error("❌ 所有數據來源都失敗")
            return []
            
        except Exception as e:
            logger.error(f"❌ 備用數據來源失敗: {e}")
            return []
    
    def parse_alternative_results(self, soup):
        """解析備用來源的結果"""
        results = []
        
        # 根據常見的第三方網站結構解析
        # 這裡需要根據實際網站結構調整
        result_rows = soup.find_all(['tr', 'div'], class_=re.compile(r'result|draw|winning', re.I))
        
        for row in result_rows:
            try:
                # 提取日期
                date_text = row.get_text()
                date_match = re.search(r'(\d{1,2}[/-]\d{1,2}[/-]\d{4})', date_text)
                
                if date_match:
                    date_str = date_match.group(1)
                    parsed_date = self.parse_date(date_str)
                    
                    if parsed_date:
                        # 提取號碼
                        numbers = re.findall(r'\b(\d{1,2})\b', date_text)
                        winning_numbers = [int(num) for num in numbers if 1 <= int(num) <= 39]
                        
                        if len(winning_numbers) >= 5:
                            results.append({
                                'date': parsed_date,
                                'numbers': winning_numbers[:5]
                            })
                            
        except Exception as e:
            logger.warning(f"⚠️ 解析備用結果時發生錯誤: {e}")
        
        return results
    
    def generate_period_number(self, date_obj):
        """生成期別號碼，格式類似539"""
        # 使用日期生成期別號碼
        # 格式: YYMMDDXXX (年月日+序號)
        year = date_obj.year % 100
        month = date_obj.month
        day = date_obj.day
        
        # 生成期別號碼
        period = f"{year:02d}{month:02d}{day:02d}001"
        return period
    
    def format_results_to_dataframe(self, results: List[Dict]) -> pd.DataFrame:
        """將爬取結果格式化為DataFrame，與539格式一致"""
        if not results:
            logger.warning("⚠️ 沒有結果需要格式化")
            return pd.DataFrame()
        
        formatted_data = []
        
        for result in results:
            date_obj = result['date']
            numbers = result['numbers']
            
            # 確保有5個號碼
            if len(numbers) != 5:
                logger.warning(f"⚠️ 號碼數量不正確: {numbers}")
                continue
            
            # 按539格式整理數據
            formatted_row = {
                '日期': date_obj.strftime('%Y-%m-%d %H:%M:%S'),
                '星期': self.get_weekday_chinese(date_obj),
                '號碼1': numbers[0],
                '號碼2': numbers[1], 
                '號碼3': numbers[2],
                '號碼4': numbers[3],
                '號碼5': numbers[4],
                '期別': self.generate_period_number(date_obj)
            }
            
            formatted_data.append(formatted_row)
            logger.info(f"📊 格式化結果: {date_obj.strftime('%Y-%m-%d')} -> {numbers}")
        
        df = pd.DataFrame(formatted_data)
        logger.info(f"✅ 成功格式化 {len(df)} 筆加州Fantasy 5開獎結果")
        
        return df
    
    def save_to_excel(self, df: pd.DataFrame, filename: str = "fantasy5_hist.xlsx"):
        """保存結果到Excel檔案"""
        try:
            if df.empty:
                logger.warning("⚠️ 沒有數據需要保存")
                return False
            
            # 檢查檔案是否存在
            if Path(filename).exists():
                # 讀取現有數據
                existing_df = pd.read_excel(filename, engine='openpyxl')
                
                # 合併數據，避免重複
                combined_df = pd.concat([existing_df, df], ignore_index=True)
                
                # 移除重複記錄（基於日期）
                combined_df = combined_df.drop_duplicates(subset=['日期'], keep='last')
                combined_df = combined_df.sort_values('日期')
                
                logger.info(f"📊 合併數據: 原有 {len(existing_df)} 筆，新增 {len(df)} 筆，合計 {len(combined_df)} 筆")
                
            else:
                combined_df = df
                logger.info(f"📊 建立新檔案，包含 {len(df)} 筆數據")
            
            # 保存到Excel
            combined_df.to_excel(filename, index=False, engine='openpyxl')
            logger.info(f"✅ 加州Fantasy 5歷史數據已保存到: {filename}")
            
            return True
            
        except Exception as e:
            logger.error(f"❌ 保存Excel檔案失敗: {e}")
            return False
    
    def crawl_and_save(self, filename: str = "fantasy5_hist.xlsx"):
        """完整的爬取和保存流程"""
        logger.info("🎲 開始加州Fantasy 5數據爬取流程")
        logger.info("="*60)
        
        # 爬取數據
        results = self.crawl_calottery_results()
        
        if not results:
            logger.error("❌ 未能獲取任何開獎結果")
            return False
        
        # 格式化數據
        df = self.format_results_to_dataframe(results)
        
        if df.empty:
            logger.error("❌ 數據格式化失敗")
            return False
        
        # 保存到Excel
        success = self.save_to_excel(df, filename)
        
        if success:
            logger.info("🎉 加州Fantasy 5數據爬取完成！")
            return True
        else:
            logger.error("❌ 數據保存失敗")
            return False

def main():
    """主程式"""
    crawler = CaliforniaFantasy5Crawler()
    success = crawler.crawl_and_save()
    
    if success:
        print("✅ 加州Fantasy 5爬蟲執行成功")
        return True
    else:
        print("❌ 加州Fantasy 5爬蟲執行失敗")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
