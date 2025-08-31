# --------------------------------------------------------
#  auto_lottery_system.py - 全自動539彩票分析系統
# --------------------------------------------------------

import sys
import os
from pathlib import Path
from datetime import datetime
import time

def check_dependencies():
    """檢查必要的檔案和模組"""
    required_files = [
        'lottery_analysis.py',
        'verify_predictions.py', 
        'lottery_crawler.py'
    ]
    
    missing_files = []
    for file in required_files:
        if not Path(file).exists():
            missing_files.append(file)
    
    if missing_files:
        print(f"❌ 缺少必要檔案: {missing_files}")
        return False
    
    # 檢查必要的Python模組
    required_modules = ['pandas', 'requests', 'beautifulsoup4', 'openpyxl']
    missing_modules = []
    
    for module in required_modules:
        try:
            if module == 'beautifulsoup4':
                import bs4
            else:
                __import__(module)
        except ImportError:
            missing_modules.append(module)
    
    if missing_modules:
        print(f"❌ 缺少必要模組，請安裝: pip install {' '.join(missing_modules)}")
        return False
    
    return True

def run_crawler():
    """執行爬蟲更新開獎資料"""
    print("\n🕷️ 步驟1: 更新開獎資料")
    print("="*40)
    
    try:
        from lottery_crawler import Lottery539Crawler
        
        crawler = Lottery539Crawler()
        success = crawler.crawl_latest_results("lottery_hist.xlsx")
        
        if success:
            print("✅ 開獎資料更新成功")
            return True
        else:
            print("⚠️ 自動更新失敗，請檢查網路連線或手動更新")
            print("\n💡 解決方法:")
            print("   1. 檢查網路連線是否正常")
            print("   2. 稍後再試（台彩官網可能暫時無法存取）")
            print("   3. 手動執行 lottery_crawler.py 進行手動輸入")
            print("   4. 如果有舊的 lottery_hist.xlsx，分析功能仍可執行")
            return False
            
    except Exception as e:
        print(f"❌ 爬蟲執行失敗: {e}")
        print("\n💡 可能的原因:")
        print("   - 網路連線問題")
        print("   - 缺少必要的Python套件 (requests, beautifulsoup4)")
        print("   - 台彩官網結構改變")
        return False

def run_analysis_and_prediction():
    """執行分析和預測"""
    print("\n🔍 步驟2: 執行分析和預測")
    print("="*40)
    
    try:
        # 檢查是否有開獎資料
        if not Path("lottery_hist.xlsx").exists():
            print("❌ 找不到開獎資料檔案，請先執行爬蟲或手動準備資料")
            return False
        
        # 執行主分析程式
        import subprocess
        import locale
        import os
        
        # 設定環境變數以確保正確的編碼
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        
        # 取得系統預設編碼
        system_encoding = locale.getpreferredencoding()
        
        result = None
        encoding_attempts = ['utf-8', system_encoding, 'cp950', 'big5']
        
        for encoding in encoding_attempts:
            try:
                result = subprocess.run(
                    [sys.executable, "lottery_analysis.py"], 
                    capture_output=True, 
                    text=True, 
                    encoding=encoding,
                    errors='replace',  # 忽略無法解碼的字符
                    env=env
                )
                break  # 成功執行就跳出迴圈
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        if result is None:
            # 如果所有編碼都失敗，使用二進制模式
            try:
                result = subprocess.run(
                    [sys.executable, "lottery_analysis.py"], 
                    capture_output=True, 
                    text=False,
                    env=env
                )
                # 手動解碼輸出
                try:
                    stdout = result.stdout.decode('utf-8', errors='replace')
                    stderr = result.stderr.decode('utf-8', errors='replace')
                except:
                    stdout = result.stdout.decode('cp950', errors='replace')
                    stderr = result.stderr.decode('cp950', errors='replace')
                
                # 創建一個模擬的結果對象
                class MockResult:
                    def __init__(self, returncode, stdout, stderr):
                        self.returncode = returncode
                        self.stdout = stdout
                        self.stderr = stderr
                
                result = MockResult(result.returncode, stdout, stderr)
            except Exception as e:
                print(f"❌ 執行分析程式時發生嚴重錯誤: {e}")
                return False
        
        if result.returncode == 0:
            print("✅ 分析和預測完成")
            # 顯示部分輸出
            output_lines = result.stdout.split('\n')
            for line in output_lines[-10:]:  # 顯示最後10行
                if line.strip():
                    print(f"   {line}")
            return True
        else:
            print(f"❌ 分析程式執行失敗: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ 分析執行失敗: {e}")
        return False

def show_summary():
    """顯示系統摘要"""
    print("\n📊 系統執行摘要")
    print("="*40)
    
    try:
        # 檢查開獎資料
        if Path("lottery_hist.xlsx").exists():
            import pandas as pd
            df = pd.read_excel("lottery_hist.xlsx", engine='openpyxl')
            print(f"📋 開獎資料: {len(df)} 期")
            if len(df) > 0:
                latest_date = df.iloc[-1].get('日期', '未知')
                print(f"   最新期數日期: {latest_date}")
        
        # 檢查預測記錄
        if Path("prediction_log.xlsx").exists():
            pred_df = pd.read_excel("prediction_log.xlsx", engine='openpyxl')
            print(f"🎯 預測記錄: {len(pred_df)} 次")
            if len(pred_df) > 0:
                latest_prediction = pred_df.iloc[-1]
                print(f"   最新預測: {latest_prediction.get('日期', '未知')} {latest_prediction.get('時間', '')}")
                
                # 統計驗證結果
                verified_count = pred_df['驗證結果'].notna().sum()
                if verified_count > 0:
                    print(f"   已驗證: {verified_count} 次")
        
        print(f"\n⏰ 系統執行時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
    except Exception as e:
        print(f"⚠️ 生成摘要時發生錯誤: {e}")

def main():
    """主函數"""
    print("🎲 全自動539彩票分析系統")
    print("="*50)
    print(f"啟動時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 檢查相依性
    if not check_dependencies():
        return
    
    success_count = 0
    total_steps = 2
    
    # 步驟1: 更新開獎資料
    if run_crawler():
        success_count += 1
    
    # 等待一下讓檔案寫入完成
    time.sleep(1)
    
    # 步驟2: 執行分析和預測
    if run_analysis_and_prediction():
        success_count += 1
    
    # 顯示摘要
    show_summary()
    
    # 最終狀態
    print(f"\n🏁 系統執行完成")
    print(f"成功步驟: {success_count}/{total_steps}")
    
    if success_count == total_steps:
        print("🎉 所有步驟執行成功！")
    else:
        print("⚠️ 部分步驟執行失敗，請檢查錯誤訊息")
    
    print("\n💡 建議工作流程:")
    print("   1. 每日一次: python lottery_analysis.py (生成預測)")
    print("   2. 開獎後: python auto_lottery_system.py --update-only (更新開獎資料)")
    print("   3. 接著: python verify_predictions.py (驗證預測結果)")

def update_only_mode():
    """僅更新開獎資料模式"""
    print("📥 開獎資料更新模式")
    print("="*30)
    print(f"執行時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 檢查相依性
    if not check_dependencies():
        return
    
    # 只執行爬蟲更新
    if run_crawler():
        print("✅ 開獎資料更新完成")
        
        # 顯示摘要（不包含預測部分）
        show_summary()
        
        print("\n💡 下一步:")
        print("   執行 python verify_predictions.py 來驗證預測結果")
    else:
        print("❌ 開獎資料更新失敗")

def prediction_only_mode():
    """僅執行預測分析模式"""
    print("🔮 預測分析模式")
    print("="*25)
    print(f"執行時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 檢查相依性
    if not check_dependencies():
        return
    
    # 檢查今天是否已有預測記錄
    today = datetime.now().strftime('%Y-%m-%d')
    if Path("prediction_log.xlsx").exists():
        try:
            import pandas as pd
            df = pd.read_excel("prediction_log.xlsx", engine='openpyxl')
            today_records = df[df['日期'] == today]
            
            if len(today_records) > 0:
                print(f"⚠️ 今日 ({today}) 已有預測記錄")
                print("是否要覆蓋現有記錄？")
                choice = input("輸入 'y' 覆蓋，其他鍵取消: ").strip().lower()
                if choice != 'y':
                    print("📋 取消預測，保留現有記錄")
                    return
                else:
                    print("🔄 將覆蓋今日現有預測記錄")
        except Exception as e:
            print(f"⚠️ 檢查現有記錄時發生錯誤: {e}")
    
    # 只執行分析和預測
    if run_analysis_and_prediction():
        print("✅ 預測分析完成")
        show_summary()
    else:
        print("❌ 預測分析失敗")

def schedule_mode():
    """排程模式（靜默執行）"""
    """可以用於定時任務的靜默模式"""
    try:
        # 簡化的執行流程，減少輸出
        from lottery_crawler import Lottery539Crawler
        
        # 更新資料
        crawler = Lottery539Crawler()
        crawler_success = crawler.crawl_latest_results("lottery_hist.xlsx")
        
        if crawler_success:
            # 執行分析
            import subprocess
            import locale
            import os
            
            # 設定環境變數以確保正確的編碼
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            
            # 取得系統預設編碼
            system_encoding = locale.getpreferredencoding()
            
            result = None
            encoding_attempts = ['utf-8', system_encoding, 'cp950', 'big5']
            
            for encoding in encoding_attempts:
                try:
                    result = subprocess.run(
                        [sys.executable, "lottery_analysis.py"], 
                        capture_output=True, 
                        text=True, 
                        encoding=encoding,
                        errors='replace',
                        env=env
                    )
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            
            if result is None:
                # 使用二進制模式作為最後手段
                try:
                    result = subprocess.run(
                        [sys.executable, "lottery_analysis.py"], 
                        capture_output=True, 
                        text=False,
                        env=env
                    )
                    stdout = result.stdout.decode('utf-8', errors='replace')
                    stderr = result.stderr.decode('utf-8', errors='replace')
                    
                    class MockResult:
                        def __init__(self, returncode, stdout, stderr):
                            self.returncode = returncode
                            self.stdout = stdout
                            self.stderr = stderr
                    
                    result = MockResult(result.returncode, stdout, stderr)
                except Exception:
                    result = None
            
            if result and result.returncode == 0:
                # 記錄成功執行
                with open("auto_system_log.txt", "a", encoding='utf-8') as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 自動執行成功\n")
                return True
            else:
                # 記錄分析失敗
                error_msg = result.stderr if result else "無法執行分析程式"
                with open("auto_system_log.txt", "a", encoding='utf-8') as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 分析失敗: {error_msg}\n")
                return False
        else:
            # 記錄爬蟲失敗
            with open("auto_system_log.txt", "a", encoding='utf-8') as f:
                f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 爬蟲失敗\n")
            return False
            
    except Exception as e:
        # 記錄系統錯誤
        with open("auto_system_log.txt", "a", encoding='utf-8') as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 系統錯誤: {e}\n")
        return False

if __name__ == "__main__":
    # 檢查命令列參數
    if len(sys.argv) > 1:
        mode = sys.argv[1]
        
        if mode == "--schedule":
            # 排程模式（靜默執行）
            schedule_mode()
        elif mode == "--update-only":
            # 僅更新開獎資料模式
            update_only_mode()
            # 等待用戶按鍵
            input("\n按 Enter 鍵退出...")
        elif mode == "--predict-only":
            # 僅執行預測分析模式
            prediction_only_mode()
            # 等待用戶按鍵
            input("\n按 Enter 鍵退出...")
        elif mode == "--help" or mode == "-h":
            # 顯示說明
            print("🎲 539彩票分析系統 - 使用說明")
            print("="*50)
            print("用法:")
            print("  python auto_lottery_system.py                 # 完整模式（爬蟲+預測）")
            print("  python auto_lottery_system.py --update-only   # 僅更新開獎資料")
            print("  python auto_lottery_system.py --predict-only  # 僅執行預測分析")
            print("  python auto_lottery_system.py --schedule      # 排程模式（靜默）")
            print("  python auto_lottery_system.py --help          # 顯示說明")
            print()
            print("建議工作流程:")
            print("  1. 每日一次: python lottery_analysis.py")
            print("  2. 開獎後: python auto_lottery_system.py --update-only")
            print("  3. 接著: python verify_predictions.py")
        else:
            print(f"❌ 未知參數: {mode}")
            print("使用 --help 查看可用選項")
    else:
        # 預設：完整模式
        print("⚠️ 注意：完整模式會同時執行爬蟲和預測分析")
        print("如果今日已有預測記錄，建議使用 --update-only 模式")
        choice = input("繼續執行完整模式？(y/N): ").strip().lower()
        if choice == 'y':
            main()
        else:
            print("已取消執行")
            print("使用 --help 查看可用選項")
        
        # 等待用戶按鍵
        input("\n按 Enter 鍵退出...")

