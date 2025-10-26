#!/usr/bin/env python3
"""
測試加州Fantasy 5爬蟲功能
"""

import pandas as pd
from datetime import datetime
import os

def test_fantasy5_crawler():
    """測試加州Fantasy 5爬蟲"""
    print("🧪 測試加州Fantasy 5爬蟲...")
    
    try:
        # 導入爬蟲
        from fantasy5_crawler import CaliforniaFantasy5Crawler
        
        # 建立爬蟲實例
        crawler = CaliforniaFantasy5Crawler()
        
        # 測試爬取功能
        print("🕷️ 開始測試爬取功能...")
        results = crawler.crawl_calottery_results()
        
        if results:
            print(f"✅ 成功爬取 {len(results)} 筆結果")
            
            # 測試格式化
            print("📊 測試數據格式化...")
            df = crawler.format_results_to_dataframe(results)
            
            if not df.empty:
                print(f"✅ 成功格式化 {len(df)} 筆數據")
                print("📋 數據格式:")
                print(df.head())
                
                # 檢查欄位格式
                expected_columns = ['日期', '星期', '號碼1', '號碼2', '號碼3', '號碼4', '號碼5', '期別']
                missing_columns = [col for col in expected_columns if col not in df.columns]
                
                if missing_columns:
                    print(f"❌ 缺少欄位: {missing_columns}")
                else:
                    print("✅ 所有必要欄位都存在")
                
                # 測試保存功能
                print("💾 測試保存功能...")
                success = crawler.save_to_excel(df, "test_fantasy5_hist.xlsx")
                
                if success:
                    print("✅ 保存功能正常")
                    
                    # 檢查檔案
                    if os.path.exists("test_fantasy5_hist.xlsx"):
                        file_size = os.path.getsize("test_fantasy5_hist.xlsx")
                        print(f"📁 測試檔案已建立: test_fantasy5_hist.xlsx ({file_size} bytes)")
                        
                        # 讀取並驗證
                        test_df = pd.read_excel("test_fantasy5_hist.xlsx", engine='openpyxl')
                        print(f"📊 檔案包含 {len(test_df)} 筆記錄")
                        print("📋 檔案內容預覽:")
                        print(test_df.head())
                        
                        # 清理測試檔案
                        os.remove("test_fantasy5_hist.xlsx")
                        print("🗑️ 已清理測試檔案")
                    else:
                        print("❌ 測試檔案未建立")
                else:
                    print("❌ 保存功能失敗")
            else:
                print("❌ 數據格式化失敗")
        else:
            print("❌ 爬取功能失敗")
            
    except ImportError as e:
        print(f"❌ 無法導入爬蟲模組: {e}")
    except Exception as e:
        print(f"❌ 測試過程中發生錯誤: {e}")

def test_fantasy5_prediction():
    """測試加州Fantasy 5預測功能"""
    print("\n🧪 測試加州Fantasy 5預測功能...")
    
    try:
        # 建立測試歷史數據
        test_data = {
            '日期': ['2025-01-01 00:00:00', '2025-01-02 00:00:00', '2025-01-03 00:00:00'],
            '星期': ['三', '四', '五'],
            '號碼1': [1, 5, 12],
            '號碼2': [8, 15, 22],
            '號碼3': [18, 25, 31],
            '號碼4': [28, 33, 36],
            '號碼5': [35, 39, 2],
            '期別': ['250101001', '250102001', '250103001']
        }
        
        test_df = pd.DataFrame(test_data)
        test_df.to_excel("test_fantasy5_hist.xlsx", index=False, engine='openpyxl')
        print("✅ 測試歷史數據已建立")
        
        # 測試預測功能
        from fantasy5_prediction_only import main as prediction_main
        
        # 執行預測
        success = prediction_main()
        
        if success:
            print("✅ 預測功能正常")
            
            # 檢查預測檔案
            if os.path.exists("fantasy5_prediction_log.xlsx"):
                pred_df = pd.read_excel("fantasy5_prediction_log.xlsx", engine='openpyxl')
                print(f"📊 預測檔案包含 {len(pred_df)} 筆記錄")
                print("📋 預測結果:")
                print(pred_df.head())
            else:
                print("❌ 預測檔案未建立")
        else:
            print("❌ 預測功能失敗")
            
    except Exception as e:
        print(f"❌ 預測測試過程中發生錯誤: {e}")
    finally:
        # 清理測試檔案
        for file in ["test_fantasy5_hist.xlsx", "fantasy5_prediction_log.xlsx"]:
            if os.path.exists(file):
                os.remove(file)
                print(f"🗑️ 已清理測試檔案: {file}")

if __name__ == "__main__":
    print("🎲 加州Fantasy 5系統測試")
    print("="*50)
    
    # 測試爬蟲
    test_fantasy5_crawler()
    
    # 測試預測
    test_fantasy5_prediction()
    
    print("\n🎉 測試完成！")
