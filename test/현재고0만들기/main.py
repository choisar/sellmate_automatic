import os
import time
from datetime import datetime
import pandas as pd
import WebAutomation

def process_csv(file_path: str) -> str:
    """CSV 파일 처리 및 결과 저장"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"CSV 파일을 찾을 수 없습니다: {file_path}")
    
    try:
        index_df = pd.read_csv(file_path, encoding='CP949')
        index_df = index_df.iloc[:, :-2]
        index_df = index_df[index_df['현재고'] < 0]
        index_df['현재고'] = 0

        current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'option_modify_form_{current_time}.csv'
        index_df.to_csv(output_filename, index=False, encoding='CP949')
        print(f"\n결과 저장 완료: {output_filename}")
        return output_filename

    except Exception as e:
        print(f"CSV 파일 처리 중 오류 발생: {e}")
        raise

def main():
    """메인 실행 함수"""
    automation = WebAutomation()
    try:
        automation.login()
        automation.navigate_to_stock_report()
        
        excel_dropdown = automation.handle_template_settings()
        excel_dropdown.click()
        time.sleep(1)
        
        excel_file_name = automation.download_excel()
        time.sleep(5)
        
        csv_path = os.path.join(os.getenv('DOWNLOAD_PATH'), excel_file_name)
        output_filename = process_csv(csv_path)
        time.sleep(5)

        # CSV 파일 업로드
        automation.upload_csv_to_options(output_filename)
        time.sleep(5)

    finally:
        automation.close()

if __name__ == "__main__":
    main()