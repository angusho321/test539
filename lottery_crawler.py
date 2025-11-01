# --------------------------------------------------------
#  lottery_crawler.py - 539開獎號碼爬蟲
# --------------------------------------------------------

import requests
from bs4 import BeautifulSoup
import pandas as pd
from io import BytesIO
from pathlib import Path
from datetime import datetime, timedelta
import re
import time
import json
import sys
import traceback

# 嘗試導入 TaiwanLotteryCrawler 作為備用
try:
    from TaiwanLottery import TaiwanLotteryCrawler
    TAIWAN_LOTTERY_AVAILABLE = True
    print("✅ TaiwanLotteryCrawler 套件可用")
except ImportError:
    TAIWAN_LOTTERY_AVAILABLE = False
    print("⚠️ TaiwanLotteryCrawler 套件不可用，使用原有爬蟲方法")

class Lottery539Crawler:
    def __init__(self):
        self.base_url = "https://www.pilio.idv.tw"
        self.lottery_page = "https://www.pilio.idv.tw/lto539/list.asp"
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
    
    def crawl_pilio_results(self):
        """從pilio網站爬取539開獎結果"""
        try:
            print("🔍 正在從pilio網站獲取539開獎資料...")
            response = self.session.get(self.lottery_page, timeout=10)
            response.raise_for_status()
            
            # 設定正確的編碼
            response.encoding = 'utf-8'
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 尋找包含開獎結果的表格
            lottery_data = []
            
            # 找到包含開獎號碼的表格行
            tables = soup.find_all('table')
            
            for table in tables:
                rows = table.find_all('tr')
                for row in rows:
                    cells = row.find_all('td')
                    if len(cells) >= 2:
                        # 檢查第一欄是否為日期格式
                        date_cell = cells[0].get_text(strip=True)
                        numbers_cell = cells[1].get_text(strip=True)
                        
                        # 檢查日期格式 (例如: 2025/08/30(六))
                        date_match = re.match(r'(\d{4}/\d{2}/\d{2})\([一二三四五六日]\)', date_cell)
                        if date_match:
                            date_str = date_match.group(1)
                            
                            # 解析號碼 (例如: "04, 05, 07, 13, 14")
                            numbers = re.findall(r'\d+', numbers_cell)
                            if len(numbers) == 5:
                                # 轉換為整數並確保是兩位數格式
                                formatted_numbers = [int(num) for num in numbers]
                                
                                lottery_data.append({
                                    '日期': date_str,
                                    '星期': date_cell.split('(')[1].split(')')[0] if '(' in date_cell else '',
                                    '號碼1': formatted_numbers[0],
                                    '號碼2': formatted_numbers[1],
                                    '號碼3': formatted_numbers[2],
                                    '號碼4': formatted_numbers[3],
                                    '號碼5': formatted_numbers[4]
                                })
            
            if lottery_data:
                print(f"✅ 成功解析 {len(lottery_data)} 期開獎資料")
                
                # 顯示最新幾期作為確認
                print("\n最新5期開獎資料:")
                for i, data in enumerate(lottery_data[:5]):
                    numbers = [data['號碼1'], data['號碼2'], data['號碼3'], data['號碼4'], data['號碼5']]
                    print(f"   {data['日期']} ({data['星期']}): {numbers}")
                
                # 轉換為DataFrame
                df = pd.DataFrame(lottery_data)
                
                # 轉換日期格式
                df['日期'] = pd.to_datetime(df['日期'])
                
                # 按日期排序（舊的在前）
                df = df.sort_values('日期').reset_index(drop=True)
                
                return df
            else:
                print("❌ 未找到開獎資料")
                return None
                
        except Exception as e:
            print(f"❌ 爬取pilio資料時發生錯誤: {e}")
            return None

    def get_539_download_links(self):
        """獲取539開獎結果的下載連結"""
        try:
            print("🔍 正在查找539開獎資料下載連結...")
            
            # 嘗試多個可能的網址
            possible_urls = [
                "https://www.taiwanlottery.com/lotto/history/result_download/",
                "https://www.taiwanlottery.com/lotto/lotto539/history/",
                "https://www.taiwanlottery.com/download/",
                "https://www.taiwanlottery.com/result_download.aspx"
            ]
            
            links = []
            
            for url in possible_urls:
                try:
                    print(f"   檢查網址: {url}")
                    response = self.session.get(url, timeout=10)
                    response.raise_for_status()
                    
                    soup = BeautifulSoup(response.text, 'html.parser')
                    
                    # 方法1: 尋找包含"539"或"今彩539"的連結
                    for link in soup.find_all('a', href=True):
                        link_text = link.text.strip()
                        href = link.get('href', '')
                        
                        if any(keyword in link_text.lower() for keyword in ['539', '今彩539', 'lotto539']):
                            if not href.startswith('http'):
                                href = self.base_url + href
                            links.append({
                                'text': link_text,
                                'url': href,
                                'source': url
                            })
                        elif any(keyword in href.lower() for keyword in ['539', 'lotto539']):
                            if not href.startswith('http'):
                                href = self.base_url + href
                            links.append({
                                'text': link_text or '539相關連結',
                                'url': href,
                                'source': url
                            })
                    
                    # 方法2: 尋找Excel檔案連結
                    for link in soup.find_all('a', href=True):
                        href = link.get('href', '')
                        if href.endswith('.xlsx') or href.endswith('.xls'):
                            if any(keyword in href.lower() for keyword in ['539', 'lotto539']):
                                if not href.startswith('http'):
                                    href = self.base_url + href
                                links.append({
                                    'text': f"Excel檔案: {link.text.strip() or href.split('/')[-1]}",
                                    'url': href,
                                    'source': url
                                })
                    
                    # 方法3: 尋找包含"開獎結果"和"下載"的連結
                    for link in soup.find_all('a', href=True):
                        link_text = link.text.strip()
                        if any(keyword in link_text for keyword in ['開獎結果', '歷史資料', '下載']) and '539' in link_text:
                            href = link.get('href', '')
                            if not href.startswith('http'):
                                href = self.base_url + href
                            links.append({
                                'text': link_text,
                                'url': href,
                                'source': url
                            })
                    
                except Exception as e:
                    print(f"   無法存取 {url}: {e}")
                    continue
            
            # 去重複
            unique_links = []
            seen_urls = set()
            for link in links:
                if link['url'] not in seen_urls:
                    unique_links.append(link)
                    seen_urls.add(link['url'])
            
            if unique_links:
                print(f"✅ 找到 {len(unique_links)} 個相關連結:")
                for i, link in enumerate(unique_links):
                    print(f"   {i+1}. {link['text']} -> {link['url']}")
                return unique_links
            else:
                print("❌ 未找到539開獎資料下載連結")
                print("💡 建議手動檢查台彩官網或使用手動輸入功能")
                return []
                
        except Exception as e:
            print(f"❌ 獲取下載連結時發生錯誤: {e}")
            return []
    
    def download_lottery_data(self, download_url):
        """下載開獎資料檔案"""
        try:
            print(f"📥 正在下載開獎資料...")
            response = self.session.get(download_url, timeout=30)
            response.raise_for_status()
            
            # 檢查是否為Excel檔案
            content_type = response.headers.get('content-type', '')
            if 'excel' not in content_type and 'spreadsheet' not in content_type:
                print(f"⚠️ 檔案類型可能不正確: {content_type}")
            
            print(f"✅ 下載完成，檔案大小: {len(response.content)} bytes")
            return response.content
            
        except Exception as e:
            print(f"❌ 下載檔案時發生錯誤: {e}")
            return None
    
    def parse_lottery_data(self, excel_content):
        """解析開獎資料Excel檔案"""
        try:
            print("🔍 正在解析開獎資料...")
            
            # 嘗試讀取Excel檔案
            try:
                df = pd.read_excel(BytesIO(excel_content), engine='openpyxl')
            except:
                # 如果openpyxl失敗，嘗試xlrd
                df = pd.read_excel(BytesIO(excel_content), engine='xlrd')
            
            print(f"📊 原始資料包含 {len(df)} 行, {len(df.columns)} 列")
            print(f"欄位名稱: {list(df.columns)}")
            
            # 顯示前幾行供檢查
            print("\n前5行資料:")
            print(df.head())
            
            # 嘗試標準化欄位名稱
            standardized_df = self.standardize_columns(df)
            
            if standardized_df is not None:
                print(f"✅ 成功解析開獎資料，共 {len(standardized_df)} 期")
                return standardized_df
            else:
                print("❌ 無法標準化資料欄位")
                return df  # 返回原始資料
                
        except Exception as e:
            print(f"❌ 解析Excel檔案時發生錯誤: {e}")
            return None
    
    def standardize_columns(self, df):
        """標準化欄位名稱"""
        try:
            # 建立欄位對應表
            column_mapping = {}
            
            # 尋找可能的欄位名稱
            for col in df.columns:
                col_str = str(col).strip()
                
                # 日期欄位
                if any(keyword in col_str for keyword in ['日期', 'date', '開獎日']):
                    column_mapping[col] = '日期'
                # 期數欄位
                elif any(keyword in col_str for keyword in ['期數', '期別', 'period']):
                    column_mapping[col] = '期數'
                # 星期欄位
                elif any(keyword in col_str for keyword in ['星期', 'day', '週']):
                    column_mapping[col] = '星期'
                # 開獎號碼欄位
                elif any(keyword in col_str for keyword in ['號碼', '開獎號碼', 'number']):
                    if '1' in col_str or 'first' in col_str.lower():
                        column_mapping[col] = '號碼1'
                    elif '2' in col_str or 'second' in col_str.lower():
                        column_mapping[col] = '號碼2'
                    elif '3' in col_str or 'third' in col_str.lower():
                        column_mapping[col] = '號碼3'
                    elif '4' in col_str or 'fourth' in col_str.lower():
                        column_mapping[col] = '號碼4'
                    elif '5' in col_str or 'fifth' in col_str.lower():
                        column_mapping[col] = '號碼5'
            
            # 如果找不到明確的號碼欄位，嘗試其他方法
            if not any('號碼' in v for v in column_mapping.values()):
                # 尋找數字欄位
                numeric_cols = []
                for col in df.columns:
                    if df[col].dtype in ['int64', 'float64']:
                        # 檢查數值範圍是否合理（1-39）
                        valid_range = df[col].dropna().between(1, 39).all()
                        if valid_range:
                            numeric_cols.append(col)
                
                # 如果找到5個數字欄位，假設為開獎號碼
                if len(numeric_cols) >= 5:
                    for i, col in enumerate(numeric_cols[:5]):
                        column_mapping[col] = f'號碼{i+1}'
            
            # 應用欄位對應
            if column_mapping:
                standardized_df = df.rename(columns=column_mapping)
                
                # 確保必要欄位存在
                required_cols = ['號碼1', '號碼2', '號碼3', '號碼4', '號碼5']
                if all(col in standardized_df.columns for col in required_cols):
                    print(f"✅ 成功標準化欄位: {column_mapping}")
                    return standardized_df
                else:
                    missing_cols = [col for col in required_cols if col not in standardized_df.columns]
                    print(f"⚠️ 缺少必要欄位: {missing_cols}")
                    return None
            else:
                print("⚠️ 無法識別欄位結構")
                return None
                
        except Exception as e:
            print(f"❌ 標準化欄位時發生錯誤: {e}")
            return None
    
    def update_excel_file(self, new_data, excel_file="lottery_hist.xlsx"):
        """更新本地Excel檔案"""
        try:
            print(f"📝 正在更新本地檔案: {excel_file}")
            
            excel_path = Path(excel_file)
            
            if excel_path.exists():
                # 讀取現有資料
                existing_df = pd.read_excel(excel_file, engine='openpyxl')
                print(f"現有資料: {len(existing_df)} 期")
                
                # 合併資料
                combined_df = pd.concat([existing_df, new_data], ignore_index=True)
                
                # 去重（基於日期）
                before_count = len(combined_df)
                if '日期' in combined_df.columns:
                    combined_df = combined_df.drop_duplicates(subset=['日期'], keep='last')
                else:
                    combined_df = combined_df.drop_duplicates()
                after_count = len(combined_df)
                
                if before_count > after_count:
                    print(f"🔄 移除了 {before_count - after_count} 筆重複資料")
                
                # 按日期排序（如果有日期欄位）
                if '日期' in combined_df.columns:
                    try:
                        combined_df['日期'] = pd.to_datetime(combined_df['日期'])
                        combined_df = combined_df.sort_values('日期')
                    except:
                        print("⚠️ 無法按日期排序")
                
                final_df = combined_df
                
            else:
                print("📄 創建新的Excel檔案")
                final_df = new_data
            
            # 儲存檔案
            final_df.to_excel(excel_file, index=False, engine='openpyxl')
            print(f"✅ 成功更新 {excel_file}，共 {len(final_df)} 期資料")
            
            # 顯示最新幾期資料
            print("\n最新5期資料:")
            print(final_df.tail())
            
            return True
            
        except Exception as e:
            print(f"❌ 更新Excel檔案時發生錯誤: {e}")
            return False
    
    def try_direct_download_urls(self):
        """嘗試直接下載常見的539資料檔案URL"""
        # 根據台彩官網的一般模式，嘗試直接下載
        direct_urls = [
            "https://www.taiwanlottery.com/upload/lotto539/lotto539.xlsx",
            "https://www.taiwanlottery.com/upload/lotto539/539.xlsx",
            "https://www.taiwanlottery.com/file/lotto539.xlsx",
            "https://www.taiwanlottery.com/download/lotto539.xlsx",
            "https://www.taiwanlottery.com/static/upload/lotto539.xlsx"
        ]
        
        print("📥 嘗試直接下載已知的539資料URL...")
        for url in direct_urls:
            try:
                print(f"   嘗試: {url}")
                response = self.session.get(url, timeout=15)
                if response.status_code == 200:
                    content_type = response.headers.get('content-type', '')
                    if 'excel' in content_type or 'spreadsheet' in content_type or len(response.content) > 1000:
                        print(f"✅ 成功下載: {url}")
                        return response.content
            except Exception as e:
                print(f"   失敗: {e}")
                continue
        
        print("❌ 直接下載失敗")
        return None

    def crawl_with_taiwan_lottery(self, excel_file="lottery_hist.xlsx"):
        """使用 TaiwanLotteryCrawler 套件獲取最新記錄"""
        if not TAIWAN_LOTTERY_AVAILABLE:
            return False
            
        try:
            print("🔄 使用 TaiwanLotteryCrawler 套件獲取最新記錄...")
            lottery = TaiwanLotteryCrawler()
            
            # 檢查是否已有歷史檔案
            excel_path = Path(excel_file)
            if excel_path.exists():
                existing_df = pd.read_excel(excel_file, engine='openpyxl')
                existing_df['日期'] = pd.to_datetime(existing_df['日期'])
                latest_date = existing_df['日期'].max()
                print(f"📊 現有歷史記錄: {len(existing_df)} 筆，最新日期: {latest_date.date()}")
            else:
                print("📄 沒有現有歷史檔案，將獲取完整記錄")
                latest_date = pd.to_datetime('2019-01-01')
            
            # 獲取當前月份的記錄，但只保留比現有記錄更新的
            current_date = datetime.now()
            year_str = str(current_date.year)
            month_str = f"{current_date.month:02d}"
            
            print(f"📅 嘗試獲取 {year_str} 年 {month_str} 月的記錄...")
            
            try:
                current_month_data = lottery.daily_cash([year_str, month_str])
                
                # 顯示返回值的詳細資訊
                if current_month_data is None:
                    print(f"⚠️ daily_cash 返回 None")
                elif isinstance(current_month_data, list):
                    print(f"📊 daily_cash 返回列表，長度: {len(current_month_data)}")
                    if len(current_month_data) > 0:
                        print(f"   第一筆資料範例: {current_month_data[0] if current_month_data else '無資料'}")
                else:
                    print(f"⚠️ daily_cash 返回非預期類型: {type(current_month_data)}")
                    print(f"   返回值: {current_month_data}")
                    
            except Exception as e:
                print(f"❌ 調用 daily_cash 時發生異常: {e}")
                print(f"   異常類型: {type(e).__name__}")
                print(f"   詳細錯誤:\n{traceback.format_exc()}")
                return False
            
            if current_month_data and len(current_month_data) > 0:
                print(f"📅 獲取到 {len(current_month_data)} 筆當月記錄")
                
                # 如果有現有記錄，只保留比最新記錄更新的資料
                if excel_path.exists():
                    filtered_data = []
                    for record in current_month_data:
                        try:
                            record_date = pd.to_datetime(record['開獎日期'])
                            if record_date > latest_date:
                                filtered_data.append(record)
                        except:
                            continue
                    
                    if filtered_data:
                        print(f"🆕 找到 {len(filtered_data)} 筆新記錄")
                        all_new_data = filtered_data
                    else:
                        print("📋 沒有比現有記錄更新的資料")
                        return True  # 沒有新資料但不算失敗
                else:
                    all_new_data = current_month_data
                    
            else:
                print("❌ 無法獲取當月記錄")
                print(f"   當前日期: {current_date.date()}")
                print(f"   查詢參數: 年={year_str}, 月={month_str}")
                if current_month_data is None:
                    print("   原因: daily_cash() 返回 None")
                elif isinstance(current_month_data, list) and len(current_month_data) == 0:
                    print("   原因: daily_cash() 返回空列表（該月份可能還沒有開獎記錄）")
                else:
                    print(f"   原因: 返回資料格式不符合預期")
                    print(f"   返回資料類型: {type(current_month_data)}")
                return False
            
            if all_new_data:
                print(f"✅ 準備更新 {len(all_new_data)} 筆新記錄")
                
                # 轉換為標準格式
                standardized_data = []
                for record in all_new_data:
                    try:
                        draw_date = pd.to_datetime(record['開獎日期'])
                        winning_numbers = record['獎號']
                        
                        if len(winning_numbers) >= 5:
                            weekday_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
                            weekday = weekday_map[draw_date.weekday()]
                            
                            standardized_record = {
                                '日期': draw_date,
                                '星期': weekday,
                                '號碼1': winning_numbers[0],
                                '號碼2': winning_numbers[1],
                                '號碼3': winning_numbers[2],
                                '號碼4': winning_numbers[3],
                                '號碼5': winning_numbers[4],
                                '期別': record.get('期別', '')
                            }
                            standardized_data.append(standardized_record)
                    except Exception as e:
                        print(f"⚠️ 處理記錄時發生錯誤: {e}")
                        continue
                
                if standardized_data:
                    new_df = pd.DataFrame(standardized_data)
                    success = self.update_excel_file(new_df, excel_file)
                    if success:
                        print("✅ 使用 TaiwanLotteryCrawler 更新成功")
                        return True
            
        except Exception as e:
            print(f"❌ TaiwanLotteryCrawler 執行失敗: {e}")
            print(f"   異常類型: {type(e).__name__}")
            print(f"   詳細錯誤:\n{traceback.format_exc()}")
        
        return False

    def crawl_latest_results(self, excel_file="lottery_hist.xlsx"):
        """爬取最新開獎結果的主要函數"""
        print("🚀 開始爬取539開獎結果...")
        print("="*50)
        
        # 1. 優先嘗試使用 TaiwanLotteryCrawler 套件
        if TAIWAN_LOTTERY_AVAILABLE:
            taiwan_success = self.crawl_with_taiwan_lottery(excel_file)
            if taiwan_success:
                print("\n🎉 使用 TaiwanLotteryCrawler 更新完成！")
                return True
            else:
                print("\n⚠️ TaiwanLotteryCrawler 失敗，嘗試原有方法...")
        
        # 2. 從pilio網站爬取資料（備用方法）
        lottery_data = self.crawl_pilio_results()
        if lottery_data is not None and len(lottery_data) > 0:
            success = self.update_excel_file(lottery_data, excel_file)
            if success:
                print("\n🎉 開獎資料更新完成！")
                return True
            else:
                print("\n❌ 更新本地檔案失敗")
        
        print("\n❌ 從pilio網站獲取資料失敗")
        print("💡 可能的原因:")
        print("   - 網路連線問題")
        print("   - pilio網站暫時無法存取")
        print("   - 網站結構改變")
        print("   - 建議稍後再試或使用手動輸入功能")
        return False

def manual_data_entry():
    """手動輸入開獎資料的備用方案"""
    print("\n📝 手動輸入開獎資料")
    print("="*30)
    
    try:
        date_str = input("請輸入開獎日期 (YYYY-MM-DD): ").strip()
        period = input("請輸入期數 (可選): ").strip()
        
        print("請輸入5個開獎號碼 (1-39):")
        numbers = []
        for i in range(5):
            while True:
                try:
                    num = int(input(f"號碼{i+1}: "))
                    if 1 <= num <= 39:
                        numbers.append(num)
                        break
                    else:
                        print("請輸入1-39之間的數字")
                except ValueError:
                    print("請輸入有效的數字")
        
        # 建立資料
        manual_data = {
            '日期': [date_str],
            '期數': [period if period else ''],
            '號碼1': [numbers[0]],
            '號碼2': [numbers[1]],
            '號碼3': [numbers[2]],
            '號碼4': [numbers[3]],
            '號碼5': [numbers[4]]
        }
        
        df = pd.DataFrame(manual_data)
        
        # 更新Excel檔案
        crawler = Lottery539Crawler()
        success = crawler.update_excel_file(df)
        
        if success:
            print("✅ 手動資料已新增到Excel檔案")
            return True
        else:
            print("❌ 手動資料新增失敗")
            return False
            
    except (EOFError, KeyboardInterrupt):
        print("\n⏹️ 手動輸入已取消")
        return False
    except Exception as e:
        print(f"❌ 手動輸入時發生錯誤: {e}")
        return False

if __name__ == "__main__":
    print("🎲 539開獎號碼爬蟲系統")
    print("="*50)
    
    crawler = Lottery539Crawler()
    
    # 嘗試自動爬取
    success = crawler.crawl_latest_results()
    
    if not success:
        # 檢查是否在互動終端中執行
        is_interactive = sys.stdin.isatty() and sys.stdout.isatty()
        
        if is_interactive:
            try:
                print("\n🤔 自動爬取失敗，是否要手動輸入開獎資料？")
                choice = input("輸入 'y' 進行手動輸入，其他鍵退出: ").strip().lower()
                
                if choice == 'y':
                    manual_data_entry()
            except (EOFError, KeyboardInterrupt):
                print("\n⏹️ 已取消手動輸入")
        else:
            print("\n⚠️ 自動爬取失敗，且目前處於非互動環境")
            print("💡 建議稍後再試或使用互動終端執行程式以進行手動輸入")
    
    print("\n" + "="*50)
    print("💡 使用說明:")
    print("1. 程式會自動嘗試從台灣彩券官網下載最新開獎資料")
    print("2. 如果自動下載失敗，可以選擇手動輸入")
    print("3. 所有資料會自動更新到 lottery_hist.xlsx")
    print("4. 建議定期執行以保持資料最新")

