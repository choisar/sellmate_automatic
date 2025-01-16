import os
import time
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import List

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, NoAlertPresentException
import pandas as pd

class SellmateAutomation:
    def __init__(self, download_timeout: int = 60):
        """
        Initialize SellmateAutomation with configuration
        
        Args:
            download_timeout (int): Maximum time to wait for downloads in seconds
        """
        load_dotenv()
        self.DOWNLOAD_TIMEOUT = download_timeout
        self.driver = None
        self.wait = None
        self.main_window = None
        
    # def initialize_driver(self) -> None:
    #   """Initialize Chrome driver with required settings"""
    #   options = webdriver.ChromeOptions()
      
    #   # 기본 다운로드 설정
    #   download_prefs = {
    #       "download.default_directory": os.path.abspath(os.getenv("DOWNLOAD_PATH")),
    #       "download.prompt_for_download": False,
    #       "download.directory_upgrade": True,
    #       "safebrowsing.enabled": True,
    #       # 추가된 설정들
    #       "profile.default_content_settings.popups": 0,
    #       "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
    #       "plugins.always_open_pdf_externally": True
    #   }
      
    #   options.add_experimental_option("prefs", download_prefs)
    #   options.add_experimental_option('excludeSwitches', ['enable-logging'])
    #   options.add_argument('--disable-gpu')
    #   options.add_argument('--no-sandbox')
    #   options.add_argument('--disable-dev-shm-usage')
      
    #   self.driver = webdriver.Chrome(options=options)
    #   self.driver.execute_cdp_cmd('Page.setDownloadBehavior', {
    #       'behavior': 'allow',
    #       'downloadPath': os.path.abspath(os.getenv("DOWNLOAD_PATH"))
    #   })
      
    #   self.wait = WebDriverWait(self.driver, 30)

    def initialize_driver(self) -> None:
        """Initialize Chrome driver with extended timeout settings"""
        options = webdriver.ChromeOptions()
        
        # 다운로드 설정
        download_prefs = {
            "download.default_directory": os.path.abspath(os.getenv("DOWNLOAD_PATH")),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0,
            "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
            "plugins.always_open_pdf_externally": True
        }
        
        options.add_experimental_option("prefs", download_prefs)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        
        # Timeouts 설정
        options.set_capability('pageLoadStrategy', 'normal')
        options.set_capability('timeouts', {
            'implicit': 10000,
            'pageLoad': 300000,  # 5분
            'script': 300000     # 5분
        })
        
        # Service 설정
        service = webdriver.ChromeService()
        
        # Driver 초기화
        self.driver = webdriver.Chrome(
            options=options,
            service=service
        )
        
        # 추가 타임아웃 설정
        self.driver.set_script_timeout(300)
        self.driver.set_page_load_timeout(300)
        
        # CDP 명령어로 타임아웃 설정 추가
        self.driver.execute_cdp_cmd('Network.enable', {})
        self.driver.execute_cdp_cmd('Network.setUserAgentOverride', {
            "userAgent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/91.0.4472.124 Safari/537.36"
        })
        
        # CDP를 통한 네트워크 타임아웃 설정
        self.driver.execute_cdp_cmd('Network.enable', {})
        self.driver.execute_cdp_cmd('Network.setBypassServiceWorker', {'bypass': True})
        
        # 다운로드 경로 설정
        self.driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': os.path.abspath(os.getenv("DOWNLOAD_PATH"))
        })
        
        # WebDriverWait 초기화
        self.wait = WebDriverWait(
            self.driver, 
            timeout=300,  # 5분
            poll_frequency=1,  # 1초마다 체크
            ignored_exceptions=None
        )
        
        print("Chrome driver initialized with extended capabilities and timeouts")
        
    def login(self) -> None:
        """Perform login to Sellmate"""
        self.driver.get(os.getenv('BASE_URL'))
        time.sleep(2)
        self.main_window = self.driver.current_window_handle
        
        userid_field = self.driver.find_element(By.ID, "userid")
        password_field = self.driver.find_element(By.ID, "passwd")
        
        userid_field.clear()
        password_field.clear()
        
        userid_field.send_keys(os.getenv("SELLMATE_ID"))
        password_field.send_keys(os.getenv("SELLMATE_PW"))
        
        self.driver.find_element(By.CLASS_NAME, "login_btn").click()
        time.sleep(2)
        
    def navigate_to_stock_report(self) -> None:
        """Navigate to stock report page"""
        self.driver.get(f"{os.getenv('BASE_URL')}/stock/progressRpt.asp")
        time.sleep(2)
        
    def handle_date_selection(self, date_type: str) -> None:
        """
        Handle different date selection types
        
        Args:
            date_type (str): Type of date selection ('today', 'month', 'custom')
        """
        if date_type == 'today':
            self.driver.find_element(By.ID, "date_today").click()
        elif date_type == 'month':
            self.driver.find_element(By.ID, "date_month").click()
        elif date_type == 'custom':
            self.driver.find_element(By.ID, "inout_custom").click()
            self.set_custom_date_range()
            
    def set_custom_date_range(self) -> None:
        """Set custom date range for 90 days"""
        sDate = self.driver.find_element(By.ID, "sDate")
        eDate = self.driver.find_element(By.ID, "eDate")
        
        today = datetime.now().date()
        ninety_days_ago = today - timedelta(days=90)
        
        eDate.clear()
        eDate.send_keys(today.strftime('%Y-%m-%d'))
        
        sDate.clear()
        sDate.send_keys(ninety_days_ago.strftime('%Y-%m-%d'))
        
    def select_suppliers(self, suppliers: List[str]) -> None:
        """
        Select specified suppliers
        
        Args:
            suppliers (List[str]): List of supplier names to select
        """
        extend_toggle_cell = self.wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "extend_toggle_cell"))
        )
        extend_toggle_cell.click()
        
        supply_button = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='공급처 선택하기']"))
        )
        supply_button.click()
        
        self.handle_supplier_window(suppliers)
        
    def handle_supplier_window(self, suppliers: List[str]) -> None:
        """Handle supplier selection in popup window"""
        for handle in self.driver.window_handles:
            if handle != self.main_window:
                self.driver.switch_to.window(handle)
                time.sleep(3)
                
                try:
                    all_checkbox = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "label.customizing.ui_checkBox .custom_checkbox"))
                    )
                    all_checkbox.click()
                    
                    self.driver.execute_script("collapse_0()")
                    
                    for supplier in suppliers:
                        supplier_checkbox = self.driver.find_element(
                            By.CSS_SELECTOR, f"input.supplier_code[supplier_name='{supplier}']"
                        )
                        self.driver.execute_script("arguments[0].click();", supplier_checkbox)
                    
                    selected_button = self.wait.until(
                        EC.element_to_be_clickable((By.ID, "selected"))
                    )
                    selected_button.click()
                except Exception as e:
                    print(f"공급처 선택 중 오류 발생: {e}")
                    raise
                
                break
        
        self.driver.switch_to.window(self.main_window)
        time.sleep(2)
        
    def download_excel(self, template_name: str) -> str:
        """
        Download Excel file with specified template
        
        Args:
            template_name (str): Name of the template to use
            
        Returns:
            str: Downloaded file name
        """
        self.driver.find_element(By.ID, "search-btn").click()
        
        excel_dropdown = self.wait.until(EC.element_to_be_clickable((By.ID, "excel_dropdown")))
        excel_dropdown.click()
        time.sleep(1)
        
        self.driver.find_element(By.XPATH, "//a[contains(text(), 'XLS양식설정')]").click()
        
        # Handle template selection window
        excel_file_name = self.handle_template_window(template_name)
        
        return excel_file_name
        
    def handle_template_window(self, template_name: str) -> str:
        """
        Handle template selection window with improved alert handling
        
        Args:
            template_name (str): Name of the template to select
        """
        wait_start = time.time()
        max_wait = 30
        template_handled = False
        
        while time.time() - wait_start < max_wait and not template_handled:
            try:
                # 현재 활성화된 창들의 핸들 목록을 가져옴
                current_handles = self.driver.window_handles
                new_handles = [h for h in current_handles if h != self.main_window]
                
                if not new_handles:
                    time.sleep(1)
                    continue
                    
                # 새 창으로 전환
                new_window = new_handles[-1]
                self.driver.switch_to.window(new_window)
                
                # 템플릿 선택
                template_select = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "xls-template-options"))
                )
                
                options = WebDriverWait(template_select, 10).until(
                    EC.presence_of_all_elements_located((By.TAG_NAME, "option"))
                )
                
                template_found = False
                for option in options:
                    if option.text == template_name:
                        option.click()
                        template_found = True
                        print(f"- '{template_name}' 템플릿을 선택했습니다.")
                        break
                        
                if not template_found:
                    raise Exception(f"템플릿 '{template_name}'을 찾을 수 없습니다.")
                
                time.sleep(1)  # 템플릿 선택 후 잠시 대기
                
                # Alert 처리 추가
                try:
                    alert = WebDriverWait(self.driver, 3).until(EC.alert_is_present())
                    alert_text = alert.text
                    print(f"- Alert 메시지: {alert_text}")  # 디버깅용
                    alert.accept()
                    time.sleep(0.5)  # alert 처리 후 잠시 대기
                except:
                    pass  # alert가 없으면 계속 진행
                
                try:
                    self.driver.close()
                except Exception as e:
                    print(f"- 창 닫기 실패: {str(e)}")
                
                self.driver.switch_to.window(self.main_window)
                template_handled = True
                
            except UnexpectedAlertPresentException as e:
                try:
                    alert = self.driver.switch_to.alert
                    alert_text = alert.text
                    print(f"- Alert 메시지: {alert_text}")  # 디버깅용
                    alert.accept()
                    time.sleep(0.5)
                except:
                    pass
                continue
                
            except Exception as e:
                print(f"- 템플릿 선택 창 처리 중 오류 발생: {str(e)}")
                time.sleep(1)
                
                try:
                    self.driver.switch_to.window(self.main_window)
                except:
                    pass
                    
        if not template_handled:
            raise TimeoutException("템플릿 선택 창을 찾을 수 없거나 처리하지 못했습니다.")
                
        # Trigger download
        excel_dropdown = self.wait.until(EC.element_to_be_clickable((By.ID, "excel_dropdown")))
        excel_dropdown.click()
        time.sleep(1)
        
        excel_download_option = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//li[@role='presentation' and @onclick='popXls();']/a"))
        )
        excel_download_option.click()
        
        # Handle download window
        excel_file_name = self.handle_download_window()
        
        return excel_file_name
        
    def handle_alerts(self) -> None:
        """Handle any alerts that appear"""
        try:
            while True:
                try:
                    alert = self.driver.switch_to.alert
                    alert.accept()
                    time.sleep(1)
                except:
                    break
        except Exception as e:
            if isinstance(e, UnexpectedAlertPresentException):
                print("예기치 않은 Alert가 발생했습니다.")
            elif isinstance(e, NoAlertPresentException):
                print("Alert가 없습니다.")
            else:
                print(f"처리되지 않은 예외 발생: {e}")
            raise e
            
    def handle_download_window(self) -> str:
        """Handle file download window and return filename"""
        excel_file_name = ""
        wait_start = time.time()
        max_wait = 30  # 최대 30초 동안 새 창을 기다림
        
        while not excel_file_name and time.time() - wait_start < max_wait:
            try:
                # 현재 활성화된 창들의 핸들 목록을 가져옴
                current_handles = self.driver.window_handles
                new_handles = [h for h in current_handles if h != self.main_window]
                
                if not new_handles:
                    time.sleep(1)
                    continue
                    
                # 새 창으로 전환
                new_window = new_handles[-1]
                self.driver.switch_to.window(new_window)
                
                # 다운로드 창 처리
                row = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "table.origin_table tbody:first-of-type tr:first-of-type")
                    )
                )
                
                download_button = WebDriverWait(row, 60).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "td:last-child button.btn-success"))
                )
                
                completed_time_cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(5)")
                completed_time = datetime.strptime(completed_time_cell.text, "%Y-%m-%d %H:%M:%S")
                excel_file_name = f"stk_optOrdProgress_{completed_time.strftime('%Y%m%d_%H%M%S')}.csv"
                
                # 다운로드 버튼 클릭
                download_button.click()
                print("- 다운로드 버튼을 클릭했습니다...")
                
                # 다운로드 완료 대기
                self.wait_for_download(excel_file_name)
                
                try:
                    self.driver.close()
                except:
                    pass
                    
                # 메인 창으로 복귀
                self.driver.switch_to.window(self.main_window)
                break
                
            except Exception as e:
                print(f"- 다운로드 창 처리 중 오류 발생: {str(e)}")
                time.sleep(1)
                continue
        
        if not excel_file_name:
            raise TimeoutException("다운로드 창을 찾을 수 없거나 처리하지 못했습니다.")
            
        return excel_file_name
        
    def wait_for_download(self, file_name: str) -> None:
        """Wait for file download to complete"""
        downloaded_file_path = os.path.join(os.getenv('DOWNLOAD_PATH'), file_name)
        download_path = Path(downloaded_file_path)
        wait_start = time.time()
        
        while not download_path.exists() or download_path.stat().st_size == 0:
            if time.time() - wait_start > self.DOWNLOAD_TIMEOUT:
                raise TimeoutException(f"파일 다운로드가 {self.DOWNLOAD_TIMEOUT}초 내에 완료되지 않았습니다.")
            time.sleep(0.5)
        print(f"다운로드 완료: {downloaded_file_path}")
        
    def process_overseas_stock(self, csv_path: str) -> List[str]:
        """Process overseas stock data"""
        index_df = pd.read_csv(csv_path, encoding='CP949')
        index_df['안정재고'] = index_df['주문수'].apply(
            lambda x: int((x / 90 * 50) if x >= 500 else ((x / 90 * 35) if x >= 1 else 0))
        )
        df_barcode_stock = index_df[['바코드번호', '안정재고']]
        
        return self.save_processed_data(df_barcode_stock)
        
    def process_domestic_stock(self, csv_path: str) -> List[str]:
        """Process domestic stock data"""
        index_df = pd.read_csv(csv_path, encoding='CP949')
        index_df['안정재고'] = index_df['주문수'].apply(lambda x: int(x/30*5))
        df_barcode_stock = index_df[['바코드번호', '안정재고']]
        
        return self.save_processed_data(df_barcode_stock)
        
    def process_negative_stock(self, csv_path: str) -> List[str]:
        """Process negative stock data"""
        index_df = pd.read_csv(csv_path, encoding='CP949')
        index_df = index_df.iloc[:, :-2]
        index_df = index_df[index_df['현재고'] < 0]
        index_df['현재고'] = 0
        
        current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'option_modify_form_{current_time}.csv'
        index_df.to_csv(output_filename, index=False, encoding='EUC-KR')
        return [output_filename]
        
    def save_processed_data(self, df: pd.DataFrame) -> List[str]:
        """Save processed data to CSV files"""
        output_files = []
        
        if len(df) > 3000:
            df_top_3000 = df.head(3000)
            df_bottom = df.iloc[3000:]
            
            current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_top_filename = f'option_modify_form_{current_time}.csv'
            df_top_3000.to_csv(output_top_filename, index=False, encoding='EUC-KR')
            output_files.append(output_top_filename)
            
            time.sleep(5)
            
            current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_bottom_filename = f'option_modify_form_{current_time}.csv'
            df_bottom.to_csv(output_bottom_filename, index=False, encoding='EUC-KR')
            output_files.append(output_bottom_filename)
        else:
            current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f'option_modify_form_{current_time}.csv'
            df.to_csv(output_filename, index=False, encoding='EUC-KR')
            output_files.append(output_filename)
            
        return output_files
        
    def upload_files(self, output_files: List[str]) -> None:
        """Upload processed files"""
        def progress_complete(driver):
            try:
                progress_bar = driver.find_element(By.CLASS_NAME, "progress-bar")
                progress_text = progress_bar.text
                
                if "완료" in progress_text:
                    numbers = [int(num) for num in re.findall(r"\d+", progress_text)]
                    if len(numbers) >= 2:
                        a, b = numbers[0], numbers[1]
                        if a / b > 1:  # a/b가 1 초과면 정상 완료
                            return True
                return False
            except:
                return False

        for output_filename in output_files:
            
            # 상품 옵션 일괄수정 페이지로 이동
            self.driver.get(f"{os.getenv('BASE_URL')}/stock/opt_csv_edit.asp")
            
            # 파일 입력 필드가 존재할 때까지 대기
            file_input = self.wait.until(
                EC.presence_of_element_located((By.NAME, "pro_filename"))
            )
            
            # CSV 파일의 절대 경로를 send_keys로 전송
            file_input.send_keys(os.path.abspath(output_filename))
            
            # 파일 전송 버튼이 클릭 가능할 때까지 대기
            submit_button = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(text(), '파일전송')]")
                )
            )
            submit_button.click()
            
            try:
                # 최대 300초(5분) 동안 progress complete 조건을 기다림
                WebDriverWait(self.driver, 600).until(progress_complete)
                print(f"{output_filename} 업로드 완료")
            except TimeoutException:
                print(f"{output_filename} 업로드 실패 - 5분 초과")

    def process_stock_data(self, process_type: str) -> None:
        """
        Main method to process stock data based on type
        
        Args:
            process_type (str): Type of processing ('overseas', 'domestic', 'negative')
        """
        try:
            print("\n[1/6] 브라우저를 초기화하고 있습니다...")
            self.initialize_driver()
            
            print("[2/6] 로그인을 진행합니다...")
            self.login()
            
            print("[3/6] 재고 보고서 페이지로 이동합니다...")
            self.navigate_to_stock_report()
            
            print("[4/6] 데이터를 필터링하고 있습니다...")
            if process_type == 'overseas':
                print("- 날짜 범위를 90일로 설정합니다...")
                self.handle_date_selection('custom')
                print("- 해외 공급처(일본, 중국, PermaBlend)를 선택합니다...")
                self.select_suppliers(['일본', '중국', 'PermaBlend'])
                print("- 안정재고수정 양식으로 데이터를 다운로드합니다...")
                excel_file = self.download_excel('안정재고수정')
                output_files = self.process_overseas_stock(
                    os.path.join(os.getenv('DOWNLOAD_PATH'), excel_file)
                )
            elif process_type == 'domestic':
                print("- 날짜를 이번 달로 설정합니다...")
                self.handle_date_selection('month')
                print("- 국내 공급처를 선택합니다...")
                self.select_suppliers(['한국'])
                print("- 안정재고수정 양식으로 데이터를 다운로드합니다...")
                excel_file = self.download_excel('안정재고수정')
                output_files = self.process_domestic_stock(
                    os.path.join(os.getenv('DOWNLOAD_PATH'), excel_file)
                )
            elif process_type == 'negative':
                print("- 날짜를 오늘로 설정합니다...")
                self.handle_date_selection('today')
                print("- 음수재고 0 만들기 양식으로 데이터를 다운로드합니다...")
                excel_file = self.download_excel('음수재고 0 만들기')
                output_files = self.process_negative_stock(
                    os.path.join(os.getenv('DOWNLOAD_PATH'), excel_file)
                )
            
            print("\n[5/6] 데이터 처리가 완료되었습니다.")
            print("[6/6] 처리된 파일을 업로드합니다...")
            self.upload_files(output_files)
            print("\n모든 작업이 성공적으로 완료되었습니다.")
            
        except Exception as e:
            print(f"처리 중 오류 발생: {e}")
            raise
        finally:
            if self.driver:
                self.driver.quit()

def main():
    """Main entry point for the script"""
    print("\n어떤 작업을 하시겠습니까?")
    print("1. 음수 재고 처리")
    print("2. 국내 재고 처리")
    print("3. 해외 재고 처리")
    print("4. 종료")
    
    while True:
        try:
            choice = input("\n작업 번호를 선택하세요 (1-4): ")
            
            if choice == "4":
                print("\n프로그램을 종료합니다.")
                break
                
            if choice not in ["1", "2", "3"]:
                print("\n잘못된 선택입니다. 1-4 사이의 숫자를 입력해주세요.")
                continue
                
            process_type = {
                "1": "negative",
                "2": "domestic",
                "3": "overseas"
            }[choice]
            
            process_name = {
                "negative": "음수 재고 처리",
                "domestic": "국내 재고 처리",
                "overseas": "해외 재고 처리"
            }[process_type]
            
            print(f"\n{process_name} 작업을 시작합니다...")
            
            automation = SellmateAutomation()
            automation.process_stock_data(process_type)
            
            print(f"\n{process_name} 작업이 완료되었습니다.")
            print("\n다른 작업을 선택하시겠습니까?")
            print("1. 음수 재고 처리")
            print("2. 국내 재고 처리")
            print("3. 해외 재고 처리")
            print("4. 종료")
            
        except Exception as e:
            print(f"\n오류가 발생했습니다: {e}")
            print("\n다시 시도하시겠습니까?")
            print("1. 음수 재고 처리")
            print("2. 국내 재고 처리")
            print("3. 해외 재고 처리")
            print("4. 종료")

if __name__ == "__main__":
    main()