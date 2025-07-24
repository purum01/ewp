import os
import time
from datetime import timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import re
import pandas as pd
from glob import glob
from datetime import datetime

from selenium.common.exceptions import (
    NoAlertPresentException,
    TimeoutException,
    UnexpectedAlertPresentException
)

base_dir = os.path.expanduser("~")
download_dir = os.path.join(base_dir, "Downloads", "KEPCO_Download")
os.makedirs(download_dir, exist_ok=True)
for file in os.listdir(download_dir):
    file_path = os.path.join(download_dir, file)
    if os.path.isfile(file_path):
        os.remove(file_path)

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(options=chrome_options)

# -----------------------
# 로그인
# -----------------------
def login(user_id, user_pw):
    global driver
    driver.get("https://pp.kepco.co.kr/")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "RSA_USER_ID")))
    driver.find_element(By.ID, "RSA_USER_ID").send_keys(user_id)
    driver.find_element(By.ID, "RSA_USER_PWD").send_keys(user_pw)
    driver.find_element(By.CLASS_NAME, "intro_btn").click()

    time.sleep(1)  # 알림창 발생 대기

    # 알림창 확인
    try:
        alert = driver.switch_to.alert
        if "로그인에 실패" in alert.text:
            alert.accept()
            return False
    except NoAlertPresentException:
        pass  # 알림 없으면 정상 진행

    # 로그인 성공 여부 판단 (로그인 버튼이 다시 뜨는 경우 실패로 간주)
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "intro_btn")))
        return False  # 로그인 실패
    except TimeoutException:
        return True  # 로그인 성공
    
# -----------------------
# 파일 다운로드
# -----------------------
start_date = None
end_date = None

def download_data(start, end):
    global start_date, end_date
    start_date = start
    end_date = end
    try:
        global driver

        # 전기사용량 페이지 이동
        driver.get("https://pp.kepco.co.kr/rs/rs0101N.do?menu_id=O010201")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "SELECT_DT")))
        current_date = start_date
        time.sleep(2)  # 페이지 로딩 대기

        start_time = time.time()  # 시작 시간 기록
        log_filename = f"download_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        log_path = os.path.join(download_dir, log_filename)
        log_file = open(log_path, "a", encoding="utf-8")
        while current_date <= end_date:
            new_date = current_date.strftime("%Y-%m-%d")
            element = driver.find_element(By.ID, "SELECT_DT")
            driver.execute_script("arguments[0].value = arguments[1];", element, new_date)
            driver.execute_script("getTotalData();")
            time.sleep(2)  # 데이터 로딩 대기
            
            driver.execute_script("excelDown();")

            log_file.write(f"{new_date} data download complete\n")
            time.sleep(1) 

            current_date += timedelta(days=1)

        end_time = time.time()
        elapsed = end_time - start_time if 'start_time' in locals() else None
        msg = "파일 다운로드 완료"
        if elapsed is not None:
            msg += f" (소요 시간: {elapsed:.2f}초)"
        print(msg)
        log_file.write(msg + "\n")

        log_file.close()
        driver.quit()
    except UnexpectedAlertPresentException as e:
        raise ValueError("📆 최대 12개월 이내의 날짜만 조회할 수 있습니다. 날짜를 다시 선택해주세요.")


# -----------------------
# 파일 Merging 작업
# -----------------------
def merge_files():

    result_dir = os.path.join(base_dir, "Downloads", "KEPCO_Results")
    excel_files = glob(os.path.join(download_dir, "*.xls*"))

    # 년도 추출
    if excel_files:
        first_file = os.path.basename(excel_files[0])
        match = re.search(r"\d{4}", first_file)
        year = match.group(0) if match else ""
    else:
        year = ""

    # Results list
    results = []

    for file in excel_files:
        try:
            # Extract date from filename
            date_str = os.path.basename(file).split("(")[-1].split(")")[0]

            # Calculate day of week
            date_obj = datetime.strptime(date_str, "%Y%m%d")
            day_of_week = ["월", "화", "수", "목", "금", "토", "일"][date_obj.weekday()]
    
           # 공휴일 엑셀 불러오기
            holiday_file = "korean_holidays.xlsx"  
            holiday_df = pd.read_excel(holiday_file)

            # 공휴일을 리스트로 변환
            holiday_list = holiday_df["date"].astype(str).tolist()

            # 휴무 여부 판단
            is_holiday = (day_of_week in ["토", "일"]) or (date_str in holiday_list)
            holiday_flag = "휴" if is_holiday else ""

            # Read HTML tables
            tables = pd.read_html(file)
            df = tables[1]

            # Rename columns
            df.columns = [
                "hour", "usage_kWh", "peak_demand_kW", "reactive_power_kVarh", "lagging", "leading",
                "CO2_t", "power_factor", "hour_2", "usage_kWh_2", "peak_demand_kW_2",
                "reactive_power_kVarh_2", "lagging_2", "leading_2", "CO2_t_2", "power_factor_2"
            ]

            # Extract hourly usage
            usage_dict = {}
            for _, row in df.iterrows():
                usage_dict[f"{int(row['hour']):02d}"] = row["usage_kWh"]
                usage_dict[f"{int(row['hour_2']):02d}"] = row["usage_kWh_2"]

            # Calculate total usage
            total_usage = sum(usage_dict.values())

            # Assemble result row
            result_row = {
                "일자": date_str,
                "요일": f"({day_of_week})",
                "휴무": holiday_flag
            }
            for h in range(1, 25):
                result_row[f"{h:02d}"] = usage_dict.get(f"{h:02d}", None)
            result_row["합계"] = total_usage

            results.append(result_row)

        except Exception as e:
            print(f"Error processing {file}: {e}")

    # Convert to DataFrame
    results_df = pd.DataFrame(results)
    results_df = results_df.sort_values(by="일자")

    # "일자" 컬럼명을 연도(예: "2025년")로 변경
    results_df = results_df.rename(columns={"일자": f"{year}년"})

    # Convert first column to datetime and format as 'YYYY-MM-DD'
    first_col = results_df.columns[0]
    results_df[first_col] = pd.to_datetime(results_df[first_col], format="%Y%m%d").dt.strftime("%Y-%m-%d")


    # Save to Excel
    os.makedirs(result_dir, exist_ok=True)
    today_str = datetime.now().strftime("%Y%m%d_%H%M")
    output_file = os.path.join(
        result_dir,
        f"{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}_시간대별_사용량_집계_{today_str}.xlsx"
    )

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        results_df.to_excel(writer, index=False, sheet_name="시간대별_사용량")

        workbook = writer.book
        worksheet = writer.sheets["시간대별_사용량"]

        # === 서식 정의 ===
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'center',
            'align': 'center',
            'fg_color': "#749ADB",  
            'border': 1
            , 'top': 2, 'bottom': 6  # top: 실선(2), bottom: 이중선(6)
        })

        # === 열 너비 조정 ===
        worksheet.set_column("A:A", 10)  # 2025년
        worksheet.set_column("B:B", 5)  # 요일 
        worksheet.set_column("C:C", 5)  # 휴무 
        worksheet.set_column("D:AA", 8)  # 01 ~ 24 시간대
        worksheet.set_column("AB:AB", 10)  # 합계


        # === 헤더 서식 적용 ===
        for col_num, value in enumerate(results_df.columns):
            worksheet.write(0, col_num, value, header_format)

        # A1 셀의 기존 서식과 동일하게 만들고 배경색만 노란색(#FFFF00)으로 지정
        a1_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'center',
            'align': 'center',
            'fg_color': "#FFFF00",  # 노란색
            'border': 1,
            'top': 2,
            'bottom': 6
        })
        worksheet.write(0, 0, results_df.columns[0], a1_format)

    print(f"엑셀 저장 완료: {output_file}") 