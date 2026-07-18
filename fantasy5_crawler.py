#!/usr/bin/env python3
"""
加州Fantasy 5爬蟲（從台灣彩券網站）
專門用於從 https://twlottery.in/en/lotteryCA5 獲取加州Fantasy 5最新開獎號碼
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import re
import time
import logging
import pytz
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class Fantasy5Crawler:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
            'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Referer': 'https://twlottery.in/',
        })
        
        # 時區設定
        self.ca_timezone = pytz.timezone('America/Los_Angeles')
        self.tw_timezone = pytz.timezone('Asia/Taipei')
        
        # 目標URL
        self.target_url = 'https://twlottery.in/en/lotteryCA5'
        
    def get_today_ca_date(self):
        """取得加州今天的日期"""
        ca_now = datetime.now(self.ca_timezone)
        return ca_now.date()
    
    def get_today_tw_date(self):
        """取得台灣今天的日期"""
        tw_now = datetime.now(self.tw_timezone)
        return tw_now.date()
    
    def convert_tw_date_to_ca_date(self, tw_date):
        """將台灣日期轉換為對應的美國日期（考慮時差）"""
        try:
            # 將台灣日期轉換為datetime對象
            tw_datetime = datetime.combine(tw_date, datetime.min.time())
            tw_datetime = self.tw_timezone.localize(tw_datetime)
            
            # 轉換為加州時間
            ca_datetime = tw_datetime.astimezone(self.ca_timezone)
            ca_date = ca_datetime.date()
            
            logger.info(f"🕐 時區轉換: 台灣 {tw_date} -> 加州 {ca_date}")
            return ca_date
            
        except Exception as e:
            logger.warning(f"⚠️ 時區轉換失敗: {e}，使用原日期")
            return tw_date
    
    def crawl_latest_results(self):
        """爬取最新開獎結果"""
        logger.info("🎯 開始爬取加州Fantasy 5最新開獎結果...")
        
        # 首先嘗試使用requests
        results = self.crawl_with_requests()
        if results:
            return results
        
        # 如果requests失敗，嘗試使用Selenium
        logger.info("🔄 嘗試使用Selenium處理動態內容...")
        results = self.crawl_with_selenium()
        if results:
            return results
        
        logger.warning("⚠️ 所有方法都無法獲取開獎結果")
        return []
    
    def crawl_with_requests(self):
        """使用requests爬取"""
        try:
            # 發送請求
            response = self.session.get(self.target_url, timeout=30)
            response.raise_for_status()
            
            logger.info(f"✅ 成功連接到 {self.target_url}")
            
            # 解析HTML
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # 尋找目標class
            results = self.parse_results(soup)
            
            if results:
                logger.info(f"🎉 成功獲取 {len(results)} 筆開獎結果")
                return results
            else:
                logger.warning("⚠️ 未找到開獎結果")
                return []
                
        except Exception as e:
            logger.error(f"❌ requests爬取失敗: {e}")
            return []
    
    def crawl_with_selenium(self):
        """使用Selenium爬取動態內容"""
        driver = None
        try:
            # 設定Chrome選項
            chrome_options = Options()
            chrome_options.add_argument('--headless')  # 無頭模式
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
            
            # 使用 webdriver-manager 自動下載和匹配正確版本的 chromedriver
            logger.info("🔧 正在設置 Chrome WebDriver...")
            try:
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=chrome_options)
                logger.info("✅ Chrome WebDriver 設置成功")
            except Exception as e:
                logger.error(f"❌ Chrome WebDriver 設置失敗: {e}")
                # 如果 webdriver-manager 失敗，嘗試使用系統預設的 chromedriver
                logger.info("🔄 嘗試使用系統預設的 Chrome WebDriver...")
                try:
                    driver = webdriver.Chrome(options=chrome_options)
                    logger.info("✅ 使用系統預設 Chrome WebDriver 成功")
                except Exception as e2:
                    logger.error(f"❌ 系統預設 Chrome WebDriver 也失敗: {e2}")
                    raise
            
            logger.info("🌐 使用Selenium開啟瀏覽器...")
            driver.get(self.target_url)
            
            # 等待頁面載入
            wait = WebDriverWait(driver, 10)
            
            # 等待可能的動態內容載入
            time.sleep(3)
            
            # 尋找目標元素
            try:
                # 嘗試尋找 List_listItem__C_wls 類別
                elements = driver.find_elements(By.CLASS_NAME, "List_listItem__C_wls")
                logger.info(f"🔍 Selenium找到 {len(elements)} 個 List_listItem__C_wls 元素")
                
                if elements:
                    results = []
                    logger.info(f"🔍 開始解析 {len(elements)} 個元素...")
                    
                    for i, element in enumerate(elements):
                        try:
                            # 獲取元素的HTML內容
                            element_html = element.get_attribute('outerHTML')
                            logger.info(f"🔍 元素 {i+1} HTML: {element_html[:200]}...")
                            
                            result = self.parse_selenium_element(element)
                            if result:
                                results.append(result)
                                logger.info(f"✅ 解析成功: {result['date']} -> {result['numbers']}")
                            else:
                                logger.warning(f"⚠️ 元素 {i+1} 解析失敗")
                        except Exception as e:
                            logger.warning(f"⚠️ 解析元素 {i+1} 失敗: {e}")
                            continue
                    
                    if results:
                        logger.info(f"🎉 Selenium成功獲取 {len(results)} 筆開獎結果")
                        return results
                    else:
                        logger.warning("⚠️ 雖然找到了元素，但無法解析出有效的開獎結果")
                
                # 如果沒找到目標類別，嘗試其他方法
                logger.info("🔍 嘗試尋找其他可能的元素...")
                
                # 尋找所有包含數字的元素
                number_elements = driver.find_elements(By.XPATH, "//*[contains(text(), '1') or contains(text(), '2') or contains(text(), '3') or contains(text(), '4') or contains(text(), '5')]")
                logger.info(f"🔍 找到 {len(number_elements)} 個包含數字的元素")
                
                # 尋找表格
                tables = driver.find_elements(By.TAG_NAME, "table")
                logger.info(f"🔍 找到 {len(tables)} 個表格")
                
                # 尋找列表
                lists = driver.find_elements(By.XPATH, "//ul | //ol")
                logger.info(f"🔍 找到 {len(lists)} 個列表")
                
                # 獲取頁面原始碼
                page_source = driver.page_source
                soup = BeautifulSoup(page_source, 'html.parser')
                results = self.parse_results(soup)
                
                if results:
                    logger.info(f"🎉 從頁面原始碼成功獲取 {len(results)} 筆開獎結果")
                    return results
                
            except Exception as e:
                logger.warning(f"⚠️ 尋找元素失敗: {e}")
            
            logger.warning("⚠️ Selenium未找到開獎結果")
            return []
            
        except Exception as e:
            logger.error(f"❌ Selenium爬取失敗: {e}")
            return []
        finally:
            if driver:
                driver.quit()
    
    def parse_selenium_element(self, element):
        """解析Selenium元素"""
        try:
            # 提取文字內容
            text = element.text
            logger.info(f"🔍 元素文字內容: {text}")
            
            # 提取日期 - 嘗試多種格式
            date_obj = None
            
            # 格式1: "Sun, Oct 26, 2025"
            date_match = re.search(r'(\w{3}, \w{3} \d{1,2}, \d{4})', text)
            if date_match:
                try:
                    date_str = date_match.group(1)
                    date_obj = datetime.strptime(date_str, '%a, %b %d, %Y').date()
                    logger.info(f"🔍 找到日期格式1: {date_str} -> {date_obj}")
                except:
                    pass
            
            # 格式2: "2025-10-26"
            if not date_obj:
                date_match = re.search(r'(\d{4}-\d{2}-\d{2})', text)
                if date_match:
                    try:
                        date_str = date_match.group(1)
                        date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
                        logger.info(f"🔍 找到日期格式2: {date_str} -> {date_obj}")
                    except:
                        pass
            
            # 格式3: "10/26/2025"
            if not date_obj:
                date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', text)
                if date_match:
                    try:
                        date_str = date_match.group(1)
                        date_obj = datetime.strptime(date_str, '%m/%d/%Y').date()
                        logger.info(f"🔍 找到日期格式3: {date_str} -> {date_obj}")
                    except:
                        pass
            
            if not date_obj:
                logger.warning("⚠️ 無法解析日期")
                return None
            
            # 將台灣日期轉換為美國日期
            ca_date = self.convert_tw_date_to_ca_date(date_obj)
            
            # 提取號碼 - 使用BeautifulSoup解析HTML
            element_html = element.get_attribute('outerHTML')
            soup = BeautifulSoup(element_html, 'html.parser')
            
            # 尋找 Ball_ball__Mmfkz 類別的元素
            ball_elements = soup.find_all(class_='Ball_ball__Mmfkz')
            logger.info(f"🔍 找到 {len(ball_elements)} 個球元素")
            
            numbers = []
            for ball in ball_elements:
                ball_text = ball.get_text().strip()
                if ball_text.isdigit():
                    num = int(ball_text)
                    if 1 <= num <= 39:
                        numbers.append(num)
                        logger.info(f"🔍 找到號碼: {num}")
            
            if len(numbers) < 5:
                logger.warning(f"⚠️ 號碼不足，只找到 {len(numbers)} 個: {numbers}")
                return None
            
            # 只取前5個號碼並排序
            result_numbers = sorted(numbers[:5])
            logger.info(f"✅ 解析成功: 台灣 {date_obj} -> 加州 {ca_date} -> {result_numbers}")
            
            return {
                'date': ca_date,  # 使用轉換後的美國日期
                'numbers': result_numbers
            }
            
        except Exception as e:
            logger.warning(f"⚠️ 解析Selenium元素失敗: {e}")
            return None
    
    def parse_results(self, soup):
        """解析開獎結果"""
        results = []
        
        # 首先檢查所有可能的類別
        logger.info("🔍 檢查網站結構...")
        
        # 尋找所有包含 'list' 的類別
        list_classes = soup.find_all(class_=re.compile(r'list', re.I))
        logger.info(f"🔍 找到 {len(list_classes)} 個包含 'list' 的類別")
        
        # 尋找所有包含 'item' 的類別
        item_classes = soup.find_all(class_=re.compile(r'item', re.I))
        logger.info(f"🔍 找到 {len(item_classes)} 個包含 'item' 的類別")
        
        # 尋找所有包含 'result' 的類別
        result_classes = soup.find_all(class_=re.compile(r'result', re.I))
        logger.info(f"🔍 找到 {len(result_classes)} 個包含 'result' 的類別")
        
        # 尋找所有包含 'number' 的類別
        number_classes = soup.find_all(class_=re.compile(r'number', re.I))
        logger.info(f"🔍 找到 {len(number_classes)} 個包含 'number' 的類別")
        
        # 檢查所有可能的類別名稱
        all_classes = set()
        for element in soup.find_all(class_=True):
            for class_name in element.get('class', []):
                all_classes.add(class_name)
        
        logger.info(f"🔍 網站上所有類別名稱: {sorted(list(all_classes))[:20]}...")  # 只顯示前20個
        
        # 嘗試尋找 List_listItem__C_wls 類別
        list_items = soup.find_all(class_='List_listItem__C_wls')
        logger.info(f"🔍 找到 {len(list_items)} 個 List_listItem__C_wls 元素")
        
        # 如果沒找到，嘗試其他可能的類別
        if not list_items:
            # 嘗試尋找其他可能的類別
            possible_classes = [
                'list-item',
                'result-item',
                'draw-item',
                'game-item',
                'lottery-item'
            ]
            
            for class_name in possible_classes:
                items = soup.find_all(class_=class_name)
                if items:
                    logger.info(f"🔍 找到 {len(items)} 個 {class_name} 元素")
                    list_items = items
                    break
        
        # 如果還是沒找到，嘗試尋找包含數字的元素
        if not list_items:
            # 尋找可能包含開獎號碼的元素
            number_elements = soup.find_all(text=re.compile(r'\d{1,2}'))
            logger.info(f"🔍 找到 {len(number_elements)} 個包含數字的文字元素")
            
            # 尋找表格
            tables = soup.find_all('table')
            logger.info(f"🔍 找到 {len(tables)} 個表格")
            
            # 尋找列表
            lists = soup.find_all(['ul', 'ol'])
            logger.info(f"🔍 找到 {len(lists)} 個列表")
        
        for item in list_items:
            try:
                result = self.parse_single_result(item)
                if result:
                    results.append(result)
                    logger.info(f"✅ 解析成功: {result['date']} -> {result['numbers']}")
            except Exception as e:
                logger.warning(f"⚠️ 解析單一結果失敗: {e}")
                continue
        
        return results
    
    def parse_single_result(self, item):
        """解析單一開獎結果"""
        try:
            # 提取日期
            date_text = self.extract_date(item)
            if not date_text:
                return None
            
            # 提取號碼
            numbers = self.extract_numbers(item)
            if len(numbers) < 5:
                return None
            
            return {
                'date': date_text,
                'numbers': sorted(numbers[:5])
            }
            
        except Exception as e:
            logger.warning(f"⚠️ 解析單一結果時發生錯誤: {e}")
            return None
    
    def extract_date(self, item):
        """從元素中提取日期"""
        # 尋找日期相關的文字
        date_patterns = [
            r'(\d{4}-\d{2}-\d{2})',  # YYYY-MM-DD
            r'(\d{2}/\d{2}/\d{4})',  # MM/DD/YYYY
            r'(\d{1,2}/\d{1,2}/\d{4})',  # M/D/YYYY
        ]
        
        item_text = item.get_text()
        
        for pattern in date_patterns:
            match = re.search(pattern, item_text)
            if match:
                date_str = match.group(1)
                try:
                    # 嘗試解析日期
                    if '-' in date_str:
                        return datetime.strptime(date_str, '%Y-%m-%d').date()
                    elif '/' in date_str:
                        return datetime.strptime(date_str, '%m/%d/%Y').date()
                except:
                    continue
        
        # 如果沒有找到標準格式，嘗試尋找其他日期表示
        logger.debug(f"未找到標準日期格式，元素文字: {item_text[:200]}...")
        return None
    
    def extract_numbers(self, item):
        """從元素中提取號碼"""
        numbers = []
        
        # 尋找所有可能的號碼元素
        number_elements = item.find_all(['span', 'div', 'td', 'li'], class_=re.compile(r'number|ball|digit', re.I))
        
        if not number_elements:
            # 如果沒有找到特定的號碼元素，從文字中提取
            item_text = item.get_text()
            numbers = re.findall(r'\b(\d{1,2})\b', item_text)
        else:
            # 從號碼元素中提取
            for element in number_elements:
                text = element.get_text().strip()
                if text.isdigit() and 1 <= int(text) <= 39:
                    numbers.append(text)
        
        # 轉換為整數並過濾有效範圍
        valid_numbers = []
        for num in numbers:
            try:
                num_int = int(num)
                if 1 <= num_int <= 39:
                    valid_numbers.append(num_int)
            except:
                continue
        
        return valid_numbers
    
    def format_result(self, result):
        """格式化結果為Excel格式"""
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
    
    def save_results(self, results):
        """保存結果到Excel檔案，合併到現有歷史檔案"""
        if not results:
            logger.warning("⚠️ 沒有結果可保存")
            return False
        
        try:
            # 格式化結果
            formatted_results = []
            for result in results:
                formatted_result = self.format_result(result)
                formatted_results.append(formatted_result)
            
            # 檢查是否有現有的歷史檔案
            history_filename = "fantasy5_hist.xlsx"
            if pd.io.common.file_exists(history_filename):
                logger.info(f"📁 讀取現有歷史檔案: {history_filename}")
                try:
                    existing_df = pd.read_excel(history_filename, engine='openpyxl')
                    logger.info(f"📊 現有記錄數: {len(existing_df)} 筆")
                except Exception as e:
                    # 讀不動現有檔案時中止，不可建立殘缺檔覆蓋雲端完整歷史
                    logger.error(f"🛑 讀取現有歷史檔案失敗: {e}。中止，避免以殘缺檔覆蓋雲端。")
                    return False
            else:
                # 歷史檔不存在通常代表下載步驟失敗；中止而非建立只含今日一筆的殘缺檔
                logger.error(f"🛑 歷史檔案不存在: {history_filename}。中止，避免建立殘缺檔覆蓋雲端。若確為首次建置，請手動放置種子檔後再執行。")
                return False
            
            # 建立新結果的DataFrame
            new_df = pd.DataFrame(formatted_results)
            
            # 檢查重複記錄
            new_records = []
            existing_dates = set()
            seen_new_dates = set()  # 追蹤新記錄中已處理的日期，避免新記錄之間重複
            
            if not existing_df.empty:
                # 取得現有記錄的日期
                for _, row in existing_df.iterrows():
                    date_str = str(row['日期'])[:10]  # 只取日期部分
                    existing_dates.add(date_str)
            
            # 過濾新記錄（同時檢查與現有記錄的重複，以及新記錄之間的重複）
            for _, row in new_df.iterrows():
                date_str = str(row['日期'])[:10]  # 只取日期部分
                if date_str not in existing_dates and date_str not in seen_new_dates:
                    new_records.append(row)
                    seen_new_dates.add(date_str)  # 標記此日期已處理
                    logger.info(f"✅ 新增記錄: {date_str} -> {row['號碼1']}, {row['號碼2']}, {row['號碼3']}, {row['號碼4']}, {row['號碼5']}")
                elif date_str in existing_dates:
                    logger.info(f"⚠️ 跳過重複記錄（已存在於歷史檔案）: {date_str}")
                else:
                    logger.info(f"⚠️ 跳過重複記錄（新記錄中重複）: {date_str}")
            
            if not new_records:
                logger.info("ℹ️ 沒有新的記錄需要添加")
                return True
            
            # 合併記錄
            new_records_df = pd.DataFrame(new_records)
            
            # 確保日期格式一致
            if not existing_df.empty:
                # 將現有資料的日期轉換為字串格式
                existing_df['日期'] = existing_df['日期'].astype(str)
            
            if not new_records_df.empty:
                # 將新資料的日期轉換為字串格式
                new_records_df['日期'] = new_records_df['日期'].astype(str)
            
            updated_df = pd.concat([existing_df, new_records_df], ignore_index=True)
            
            # 按日期排序
            updated_df = updated_df.sort_values('日期')
            
            # 保存更新後的檔案
            updated_df.to_excel(history_filename, index=False, engine='openpyxl')
            
            logger.info(f"✅ 歷史檔案已更新: {history_filename}")
            logger.info(f"📊 總記錄數: {len(updated_df)} 筆")
            logger.info(f"📈 新增記錄數: {len(new_records)} 筆")
            
            # 顯示最新結果
            if new_records:
                latest = new_records[-1]  # 最新的記錄
                logger.info(f"🎯 最新開獎結果:")
                logger.info(f"   日期: {latest['日期']}")
                logger.info(f"   號碼: {latest['號碼1']}, {latest['號碼2']}, {latest['號碼3']}, {latest['號碼4']}, {latest['號碼5']}")
            
            return True
            
        except Exception as e:
            logger.error(f"❌ 保存結果失敗: {e}")
            return False

def main():
    """主程式"""
    logger.info("🎲 加州Fantasy 5爬蟲")
    logger.info("="*60)
    
    crawler = Fantasy5Crawler()
    
    # 爬取最新開獎結果
    results = crawler.crawl_latest_results()
    
    if results:
        # 保存結果
        success = crawler.save_results(results)
        if success:
            logger.info("✅ 程式執行完成")
        else:
            logger.error("❌ 保存結果失敗")
            import sys
            sys.exit(1)
    else:
        logger.warning("⚠️ 未找到開獎結果")
        logger.info("💡 可能的原因:")
        logger.info("   1. 網站結構變更")
        logger.info("   2. 網路連線問題")
        logger.info("   3. 網站暫時無法訪問")

if __name__ == "__main__":
    main()
