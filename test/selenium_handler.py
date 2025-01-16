from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, NoAlertPresentException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import os
from datetime import datetime
from dotenv import load_dotenv
from pathlib import Path
import re
from datetime import timedelta

# .env 파일 로드
load_dotenv()

DOWNLOAD_TIMEOUT = 60  # 다운로드 대기 시간


def process_stock_data():
    # Chrome 옵션 설정
    options = webdriver.ChromeOptions()
    download_prefs = {
        "download.default_directory": os.path.expanduser(os.getenv("DOWNLOAD_PATH")),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", download_prefs)

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    try:
        # 페이지 열기 및 로그인
        driver.get(os.getenv("BASE_URL"))
        time.sleep(2)
        main_window = driver.current_window_handle  # 현재 창 핸들 저장

        # 로그인 정보 입력
        userid_field = driver.find_element(By.ID, "userid")
        password_field = driver.find_element(By.ID, "passwd")

        userid_field.clear()
        password_field.clear()

        userid_field.send_keys(os.getenv("SELLMATE_ID"))
        password_field.send_keys(os.getenv("SELLMATE_PW"))

        driver.find_element(By.CLASS_NAME, "login_btn").click()
        time.sleep(2)

        # 재고 보고서 페이지 접근
        driver.get(f"{os.getenv('BASE_URL')}/stock/progressRpt.asp")
        time.sleep(2)

        driver.find_element(By.ID, "inout_custom").click()

        # 웹 요소 찾기
        sDate = driver.find_element(By.ID, "sDate")
        eDate = driver.find_element(By.ID, "eDate")

        # 오늘 날짜 설정
        today = datetime.now().date()

        # eDate를 오늘 날짜로 설정
        eDate.clear()
        eDate.send_keys(today.strftime('%Y-%m-%d'))

        # sDate를 90일 전으로 설정
        ninety_days_ago = today - timedelta(days=90)
        sDate.clear()
        sDate.send_keys(ninety_days_ago.strftime('%Y-%m-%d'))

        # 날짜 검증
        current_sdate = datetime.strptime(sDate.get_attribute("value"), '%Y-%m-%d').date()
        current_edate = datetime.strptime(eDate.get_attribute("value"), '%Y-%m-%d').date()

        # 날짜 간격이 90일인지 확인
        if (current_edate - current_sdate).days != 90:
            print(f"Warning: Date difference is not 90 days")
            sDate.clear()
            sDate.send_keys(ninety_days_ago.strftime('%Y-%m-%d'))

        extend_toggle_cell = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "extend_toggle_cell"))
        )
        extend_toggle_cell.click()

        supply_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='공급처 선택하기']"))
        )
        supply_button.click()

        for handle in driver.window_handles:
            if handle != main_window:
                driver.switch_to.window(handle)
                time.sleep(3)  # 잠시 대기
                try:
                    # 전체 선택 체크박스를 클릭하여 모두 해제
                    all_checkbox = wait.until(
                        EC.element_to_be_clickable(
                            (
                                By.CSS_SELECTOR,
                                "label.customizing.ui_checkBox .custom_checkbox",
                            )
                        )
                    )
                    all_checkbox.click()

                    driver.execute_script("collapse_0()")

                    suppliers = ["일본", "중국", "PermaBlend"]

                    for supplier in suppliers:
                        supplier_checkbox = driver.find_element(
                            By.CSS_SELECTOR,
                            f"input.supplier_code[supplier_name='{supplier}']",
                        )
                        driver.execute_script("arguments[0].click();", supplier_checkbox)

                    selected_button = wait.until(
                        EC.element_to_be_clickable((By.ID, "selected"))
                    )
                    selected_button.click()
                except:
                    print("공급처 선택 중 오류 발생")
                    raise

        driver.switch_to.window(main_window)  # 메인 창으로 돌아가기
        time.sleep(2)

        driver.find_element(By.ID, "search-btn").click()
        # time.sleep(2)

        # excel_dropdown = driver.find_element(By.ID, "excel_dropdown")
        excel_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "excel_dropdown")))
        excel_dropdown.click()
        time.sleep(1)

        driver.find_element(By.XPATH, "//a[contains(text(), 'XLS양식설정')]").click()

        for handle in driver.window_handles:
            if handle != main_window:
                driver.switch_to.window(handle)
                time.sleep(3)  # 잠시 대기

                # Alert이 있는 경우 처리
                try:
                    while True:
                        try:
                            alert = driver.switch_to.alert
                            alert.accept()  # Alert 확인 클릭
                            time.sleep(1)  # 잠시 대기
                        except:
                            break
                except Exception as e:  # 모든 예외를 포착
                    if isinstance(
                        e, UnexpectedAlertPresentException
                    ):  # UnexpectedAlertPresentException 처리 코드
                        print("예기치 않은 Alert가 발생했습니다.")
                        raise e

                    elif isinstance(
                        e, NoAlertPresentException
                    ):  # NoAlertPresentException 처리 코드
                        print("Alert가 없습니다.")
                        raise e

                    else:  # 다른 예외 처리
                        print(f"처리되지 않은 예외 발생: {e}")
                        raise e

                # "안정재고수정" 선택
                template_select = driver.find_element(By.CLASS_NAME, "xls-template-options")
                for option in template_select.find_elements(By.TAG_NAME, "option"):
                    if option.text == "안정재고수정":
                        option.click()
                        break

                driver.close()
                break

        driver.switch_to.window(main_window)  # 메인 창으로 돌아가기

        excel_dropdown.click()
        time.sleep(1)

        # '엑셀다운로드' 항목 선택
        excel_download_option = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//li[@role='presentation' and @onclick='popXls();']/a")
            )
        )
        excel_download_option.click()

        excel_file_name = ""

        # 새 창에서 다운로드 버튼 클릭
        for handle in driver.window_handles:
            if handle != main_window:
                driver.switch_to.window(handle)
                try:
                    # 다운로드 버튼이 있는 테이블 행 추출
                    row = wait.until(
                        EC.presence_of_element_located(
                            (
                                By.CSS_SELECTOR,
                                "table.origin_table tbody:first-of-type tr:first-of-type",
                            )
                        )
                    )

                    # 다운로드 버튼 클릭
                    download_button = WebDriverWait(row, 120).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "td:last-child button.btn-success")
                        )
                    )

                    # 완료시간 추출
                    completed_time_cell = row.find_element(
                        By.CSS_SELECTOR, "td:nth-child(5)"
                    )
                    completed_time = datetime.strptime(
                        completed_time_cell.text, "%Y-%m-%d %H:%M:%S"
                    )
                    excel_file_name = (
                        f"stk_optOrdProgress_{completed_time.strftime('%Y%m%d_%H%M%S')}.csv"
                    )

                    # 다운로드 버튼 클릭
                    download_button.click()

                    # 다운로드 완료 대기 개선
                    downloaded_file_path = os.path.join(
                        os.getenv("DOWNLOAD_PATH"), excel_file_name
                    )
                    download_path = Path(downloaded_file_path)
                    wait_start = time.time()

                    while not download_path.exists() or download_path.stat().st_size == 0:
                        if time.time() - wait_start > DOWNLOAD_TIMEOUT:
                            raise TimeoutException(
                                f"파일 다운로드가 {DOWNLOAD_TIMEOUT}초 내에 완료되지 않았습니다."
                            )
                        time.sleep(0.5)
                    print(f"다운로드 완료: {downloaded_file_path}")

                except TimeoutException as e:
                    print(f"오류 발생: {e}")
                    raise e

                driver.close()
                break

        # CSV 파일 처리
        time.sleep(5)
        csv_path = os.path.join(os.getenv("DOWNLOAD_PATH"), excel_file_name)

        # 파일 존재 확인
        if not os.path.exists(csv_path):
            raise FileNotFoundError(f"CSV 파일을 찾을 수 없습니다: {csv_path}")

        try:
            index_df = pd.read_csv(csv_path, encoding="CP949")
            index_df["안정재고"] = index_df['주문수'].apply(lambda x: int((x / 90 * 50) if x >= 500 else ((x / 90 * 35) if x >= 1 else 0)))
            df_barcode_stock = index_df[["바코드번호", "안정재고"]]

            output_files = []

            if len(df_barcode_stock) >3000:
                df_top_3000 = df_barcode_stock.head(3000)
                df_bottom = df_barcode_stock.iloc[3000:]

                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_top_filename = f"option_modify_form_{current_time}.csv"
                df_top_3000.to_csv(output_top_filename, index=False, encoding="EUC-KR")
                output_files.append(output_top_filename)

                time.sleep(5)  # 5초 텀을 줍니다

                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_bottom_filename = f"option_modify_form_{current_time}.csv"
                df_bottom.to_csv(output_bottom_filename, index=False, encoding="EUC-KR")
                output_files.append(output_bottom_filename)

                print(f"\n결과 저장 완료: {output_top_filename}, {output_bottom_filename}")
            else:
                # # 결과 저장
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = f"option_modify_form_{current_time}.csv"
                df_barcode_stock.to_csv(output_filename, index=False, encoding="EUC-KR")
                output_files.append(output_filename)
                print(f"\n결과 저장 완료: {output_filename}")

        except Exception as e:
            print(f"CSV 파일 처리 중 오류 발생: {e}")
            raise  # 에러 전파하여 디버깅 용이하게

        driver.switch_to.window(main_window)  # 메인 창으로 돌아가기

        # Custom expected condition 정의
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

        # 각 파일별로 순차적으로 업로드 진행
        for output_filename in output_files:
            upload_complete = False

            # 상품 옵션 일괄수정 페이지로 이동
            driver.get(f"{os.getenv('BASE_URL')}/stock/opt_csv_edit.asp")

            # 파일 입력 필드가 존재할 때까지 대기 (최대 10초)
            file_input = wait.until(
                EC.presence_of_element_located((By.NAME, "pro_filename"))
            )

            # CSV 파일의 절대 경로를 send_keys로 전송
            file_input.send_keys(os.path.abspath(output_filename))

            # 파일 전송 버튼이 클릭 가능할 때까지 대기 (최대 10초)
            submit_button = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(text(), '파일전송')]")
                )
            )
            submit_button.click()

            try:
                # 최대 300초(5분) 동안 progress complete 조건을 기다림
                WebDriverWait(driver, 300).until(progress_complete)
                print(f"{output_filename} 업로드 완료")
                upload_complete = True
            except TimeoutException:
                print(f"{output_filename} 업로드 실패 - 5분 초과")

        print("모든 파일 처리 완료")

    finally:
        # 브라우저 닫기
        driver.quit()
    return "작업 완료"
