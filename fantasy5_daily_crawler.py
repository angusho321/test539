#!/usr/bin/env python3
"""
加州Fantasy 5每日開獎爬蟲
專門用於獲取當日最新的開獎結果
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import re
import time
import logging
import pytz

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class Fantasy5DailyCrawler:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
        })
        
        # 加州時區
        self.ca_timezone = pytz.timezone('America/Los_Angeles')
        self.tw_timezone = pytz.timezone('Asia/Taipei')
        
    def get_today_ca_date(self):
        """取得加州今天的日期"""
        ca_now = datetime.now(self.ca_timezone)
        return ca_now.date()
    
    def get_today_tw_date(self):
        """取得台灣今天的日期"""
        tw_now = datetime.now(self.tw_timezone)
        return tw_now.date()
    
    def crawl_daily_results(self):
        """爬取當日開獎結果"""
        logger.info("🎯 開始爬取當日加州Fantasy 5開獎結果...")
        
        # 取得今天的日期
        today_ca = self.get_today_ca_date()
        today_tw = self.get_today_tw_date()
        
        logger.info(f"📅 加州今天: {today_ca}")
        logger.info(f"📅 台灣今天: {today_tw}")
        
        # 嘗試多個數據源
        data_sources = [
            {
                'name': '加州彩票官方網站',
                'url': 'https://www.calottery.com/fantasy-5',
                'method': 'official'
            },
            {
                'name': 'LotteryUSA',
                'url': 'https://www.lotteryusa.com/california/fantasy-5/',
                'method': 'lotteryusa'
            },
            {
                'name': '台灣彩券網',
                'url': 'https://twlottery.in/en/lotteryCA5',
                'method': 'taiwan'
            }
        ]
        
        for source in data_sources:
            try:
                logger.info(f"🔍 嘗試數據源: {source['name']}")
                response = self.session.get(source['url'], timeout=30)
                response.raise_for_status()
                
                soup = BeautifulSoup(response.content, 'html.parser')
                results = self.parse_daily_results(soup, source['method'], today_ca)
                
                if results:
                    logger.info(f"✅ 從 {source['name']} 成功獲取 {len(results)} 筆當日開獎結果")
                    return results
                else:
                    logger.warning(f"⚠️ {source['name']} 未找到當日開獎結果")
                    
            except Exception as e:
                logger.warning(f"⚠️ {source['name']} 爬取失敗: {e}")
                continue
        
        logger.error("❌ 所有數據源都無法獲取當日開獎結果")
        return []
    
    def parse_daily_results(self, soup, method, target_date):
        """解析當日開獎結果"""
        results = []
        
        if method == 'official':
            results = self.parse_official_daily(soup, target_date)
        elif method == 'lotteryusa':
            results = self.parse_lotteryusa_daily(soup, target_date)
        elif method == 'taiwan':
            results = self.parse_taiwan_daily(soup, target_date)
        
        return results
    
    def parse_official_daily(self, soup, target_date):
        """解析加州彩票官方網站的當日結果"""
        results = []
        
        # 尋找最新的開獎結果
        result_containers = soup.find_all(['div', 'section'], class_=re.compile(r'result|draw|winning|number|game', re.I))
        
        for container in result_containers:
            # 檢查是否包含今天的日期
            container_text = container.get_text()
            if str(target_date) in container_text or target_date.strftime('%m/%d/%Y') in container_text:
                # 提取號碼
                numbers = self.extract_numbers_from_text(container_text)
                
                if len(numbers) >= 5:
                    results.append({
                        'date': target_date,
                        'numbers': sorted(numbers[:5])
                    })
                    logger.info(f"✅ 官方網站找到當日結果: {target_date} -> {sorted(numbers[:5])}")
                    break
        
        return results
    
    def parse_lotteryusa_daily(self, soup, target_date):
        """解析LotteryUSA的當日結果"""
        results = []
        
        # 尋找結果表格
        tables = soup.find_all('table')
        
        for table in tables:
            rows = table.find_all('tr')
            
            for row in rows:
                row_text = row.get_text()
                
                # 檢查是否包含今天的日期
                if str(target_date) in row_text or target_date.strftime('%m/%d/%Y') in row_text:
                    numbers = self.extract_numbers_from_text(row_text)
                    
                    if len(numbers) >= 5:
                        results.append({
                            'date': target_date,
                            'numbers': sorted(numbers[:5])
                        })
                        logger.info(f"✅ LotteryUSA找到當日結果: {target_date} -> {sorted(numbers[:5])}")
                        break
        
        return results
    
    def parse_taiwan_daily(self, soup, target_date):
        """解析台灣彩券網的當日結果"""
        results = []
        
        # 尋找結果表格
        tables = soup.find_all('table')
        
        for table in tables:
            rows = table.find_all('tr')
            
            for row in rows:
                row_text = row.get_text()
                
                # 檢查是否包含今天的日期
                if str(target_date) in row_text or target_date.strftime('%Y-%m-%d') in row_text:
                    numbers = self.extract_numbers_from_text(row_text)
                    
                    if len(numbers) >= 5:
                        results.append({
                            'date': target_date,
                            'numbers': sorted(numbers[:5])
                        })
                        logger.info(f"✅ 台灣彩券網找到當日結果: {target_date} -> {sorted(numbers[:5])}")
                        break
        
        return results
    
    def extract_numbers_from_text(self, text):
        """從文字中提取號碼"""
        numbers = re.findall(r'\b(\d{1,2})\b', text)
        return [int(num) for num in numbers if 1 <= int(num) <= 39]
    
    def format_daily_result(self, result):
        """格式化當日結果為Excel格式"""
        date_obj = result['date']
        numbers = result['numbers']
        
        # 生成期別號碼
        year = date_obj.year % 100
        month = date_obj.month
        day = date_obj.day
        period = f"{year:02d}{month:02d}{day:02d}001"
        
        # 取得中文星期
        weekdays = ['一', '二', '三', '四', '五', '六', '日']
        weekday = weekdays[date_obj.weekday()]
        
        return {
            '日期': date_obj.strftime('%Y-%m-%d %H:%M:%S'),
            '星期': weekday,
            '號碼1': numbers[0],
            '號碼2': numbers[1],
            '號碼3': numbers[2],
            '號碼4': numbers[3],
            '號碼5': numbers[4],
            '期別': period
        }
    
    def update_history_file(self, new_result):
        """更新歷史檔案"""
        try:
            # 讀取現有歷史檔案
            if pd.io.common.file_exists('fantasy5_hist.xlsx'):
                df = pd.read_excel('fantasy5_hist.xlsx', engine='openpyxl')
                logger.info(f"📁 讀取現有歷史檔案，包含 {len(df)} 筆記錄")
            else:
                df = pd.DataFrame(columns=['日期', '星期', '號碼1', '號碼2', '號碼3', '號碼4', '號碼5', '期別'])
                logger.info("📁 建立新的歷史檔案")
            
            # 格式化新結果
            new_record = self.format_daily_result(new_result)
            
            # 檢查是否已存在
            existing_dates = df['日期'].str[:10] if not df.empty else []
            new_date = new_record['日期'][:10]
            
            if new_date in existing_dates.values:
                logger.warning(f"⚠️ 日期 {new_date} 的記錄已存在，跳過更新")
                return False
            
            # 新增記錄
            new_df = pd.DataFrame([new_record])
            updated_df = pd.concat([df, new_df], ignore_index=True)
            
            # 按日期排序
            updated_df = updated_df.sort_values('日期')
            
            # 保存檔案
            updated_df.to_excel('fantasy5_hist.xlsx', index=False, engine='openpyxl')
            
            logger.info(f"✅ 成功更新歷史檔案，新增記錄: {new_date} -> {new_record['號碼1']}, {new_record['號碼2']}, {new_record['號碼3']}, {new_record['號碼4']}, {new_record['號碼5']}")
            return True
            
        except Exception as e:
            logger.error(f"❌ 更新歷史檔案失敗: {e}")
            return False

def main():
    """主程式"""
    logger.info("🎲 加州Fantasy 5每日開獎爬蟲")
    logger.info("="*60)
    
    crawler = Fantasy5DailyCrawler()
    
    # 爬取當日開獎結果
    results = crawler.crawl_daily_results()
    
    if results:
        logger.info(f"🎉 成功獲取 {len(results)} 筆當日開獎結果")
        
        # 更新歷史檔案
        for result in results:
            success = crawler.update_history_file(result)
            if success:
                logger.info("✅ 歷史檔案更新成功")
            else:
                logger.warning("⚠️ 歷史檔案更新失敗或記錄已存在")
    else:
        logger.warning("⚠️ 未找到當日開獎結果")
        logger.info("💡 可能的原因:")
        logger.info("   1. 開獎時間未到 (加州時間 6:30 PM)")
        logger.info("   2. 網站結構變更")
        logger.info("   3. 網路連線問題")

if __name__ == "__main__":
    main()
