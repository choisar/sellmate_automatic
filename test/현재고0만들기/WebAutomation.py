import os
import time
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import (
    TimeoutException,
    UnexpectedAlertPresentException,
    NoAlertPresentException
)

class WebAutomation:
    def __init__(self):
        load_dotenv()
        self.DOWNLOAD_TIMEOUT = 60
        self.driver = None
        self.wait = None
        self.main_window = None
        self.setup_driver()

    def setup_driver(self) -> None:
        """Chrome 드라이버 설정"""
        options = webdriver.ChromeOptions()
        download_prefs = {
            "download.default_directory": os.path.expanduser(os.getenv('DOWNLOAD_PATH')),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", download_prefs)
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 2)

    def login(self) -> None:
        """웹사이트 로그인"""
        self.driver.get(os.getenv('BASE_URL'))
        time.sleep(2)

        userid_field = self.driver.find_element(By.ID, "userid")
        password_field = self.driver.find_element(By.ID, "passwd")
        
        userid_field.clear()
        password_field.clear()
        
        userid_field.send_keys(os.getenv("SELLMATE_ID"))
        password_field.send_keys(os.getenv("SELLMATE_PW"))
        
        self.driver.find_element(By.CLASS_NAME, "login_btn").click()
        time.sleep(2)

    def navigate_to_stock_report(self) -> None:
        """재고 보고서 페이지로 이동 및 설정"""
        self.driver.get(f"{os.getenv('BASE_URL')}/stock/progressRpt.asp")
        time.sleep(2)

        self.driver.find_element(By.ID, "date_today").click()
        self.driver.find_element(By.ID, "search-btn").click()
        time.sleep(2)

    def handle_template_settings(self) -> None:
        """엑셀 템플릿 설정 처리"""
        excel_dropdown = self.driver.find_element(By.ID, "excel_dropdown")
        excel_dropdown.click()
        time.sleep(1)

        self.driver.find_element(By.XPATH, "//a[contains(text(), 'XLS양식설정')]").click()
        self.main_window = self.driver.current_window_handle

        self._handle_template_window()
        return excel_dropdown

    def _handle_template_window(self) -> None:
        """템플릿 설정 창 처리"""
        for handle in self.driver.window_handles:
            if handle != self.main_window:
                self.driver.switch_to.window(handle)
                time.sleep(3)
                
                self._handle_alerts()
                self._select_template_option()
                
                self.driver.close()
                break
        
        self.driver.switch_to.window(self.main_window)

    def _handle_alerts(self) -> None:
        """알림창 처리"""
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
                raise e
            elif isinstance(e, NoAlertPresentException):
                print("Alert가 없습니다.")
                raise e
            else:
                print(f"처리되지 않은 예외 발생: {e}")
                raise e

    def _select_template_option(self) -> None:
        """템플릿 옵션 선택"""
        template_select = self.driver.find_element(By.CLASS_NAME, "xls-template-options")
        for option in template_select.find_elements(By.TAG_NAME, "option"):
            if option.text == "음수재고 0 만들기":
                option.click()
                break

    def download_excel(self) -> str:
        """엑셀 파일 다운로드"""
        excel_download_option = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//li[@role='presentation' and @onclick='popXls();']/a"))
        )
        excel_download_option.click()

        return self._handle_download_window()

    def _handle_download_window(self) -> str:
        """다운로드 창 처리 및 파일명 반환"""
        excel_file_name = ""
        for handle in self.driver.window_handles:
            if handle != self.main_window:
                self.driver.switch_to.window(handle)
                excel_file_name = self._process_download()
                self.driver.close()
                break
        return excel_file_name

    def _process_download(self) -> str:
        """다운로드 처리 및 완료 대기"""
        try:
            row = WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "table.origin_table tbody:first-of-type tr:first-of-type")
                )
            )

            download_button = WebDriverWait(row, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "td:last-child button.btn-success"))
            )

            completed_time_cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(5)")
            completed_time = datetime.strptime(completed_time_cell.text, "%Y-%m-%d %H:%M:%S")
            excel_file_name = f"stk_optOrdProgress_{completed_time.strftime('%Y%m%d_%H%M%S')}.csv"

            download_button.click()
            self._wait_for_download(excel_file_name)
            
            return excel_file_name

        except TimeoutException as e:
            print(f"오류 발생: {e}")
            raise e

    def _wait_for_download(self, file_name: str) -> None:
        """파일 다운로드 완료 대기"""
        downloaded_file_path = os.path.join(os.getenv('DOWNLOAD_PATH'), file_name)
        download_path = Path(downloaded_file_path)
        wait_start = time.time()
        
        while not download_path.exists() or download_path.stat().st_size == 0:
            if time.time() - wait_start > self.DOWNLOAD_TIMEOUT:
                raise TimeoutException(f"파일 다운로드가 {self.DOWNLOAD_TIMEOUT}초 내에 완료되지 않았습니다.")
            time.sleep(0.5)
        print(f"다운로드 완료: {downloaded_file_path}")

    def upload_csv_to_options(self, csv_file_path: str) -> None:
        """옵션 일괄수정 페이지로 이동하여 CSV 파일 업로드"""
        try:
            self.driver.switch_to.window(self.main_window)  # 메인 창으로 돌아가기
            
            # 상품 옵션 일괄수정 페이지로 이동
            self.driver.get(f"{os.getenv('BASE_URL')}/stock/opt_csv_edit.asp")

            # 파일 입력 필드가 존재할 때까지 대기 (최대 10초)
            file_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.NAME, "pro_filename"))
            )

            # 생성된 CSV 파일의 절대 경로를 send_keys로 전송
            file_input.send_keys(os.path.abspath(csv_file_path))

            # 파일 전송 버튼이 클릭 가능할 때까지 대기 (최대 10초)
            submit_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '파일전송')]"))
            )
            submit_button.click()
            
        except Exception as e:
            print(f"CSV 파일 업로드 중 오류 발생: {e}")
            raise e

    def close(self) -> None:
        """브라우저 종료"""
        if self.driver:
            self.driver.quit()