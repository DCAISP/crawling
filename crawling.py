from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd

# WebDriver 설정
chrome_options = Options()
chrome_options.add_argument('--disable-popup-blocking')
chrome_options.add_argument('--start-maximized')

# ChromeDriver 경로 설정
webdriver_service = Service('C:/chromedriver-win64/chromedriver.exe')

# Chrome WebDriver 초기화
driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)

# 크롤링할 웹페이지 주소 설정
url = 'https://cleansys.or.kr/dataMS.do#none'
driver.get(url)

def select_dropdown_option(dropdown_id, option_text):
    dropdown = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, dropdown_id))
    )
    dropdown.click()
    option = dropdown.find_element(By.XPATH, f"//option[contains(text(), '{option_text}')]")
    option.click()

def get_table_data():
    table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '.table-box.type01'))
    )
    rows = table.find_elements(By.TAG_NAME, 'tr')
    data = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, 'td')
        data.append([col.text for col in cols])
    return data

try:
    # '서울특별시' 선택
    select_dropdown_option('selectArea', '서울특별시')

    # '서울시 중랑물재생센터' 선택
    select_dropdown_option('keyWordSelect', '서울시 중랑물재생센터')

    # 조회 버튼 클릭
    search_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'btnSearch'))
    )
    search_button.click()

    # 데이터 읽어오기
    time.sleep(5)  # 데이터 로딩을 위한 대기 시간
    data = get_table_data()

    # 데이터 프레임으로 변환
    df = pd.DataFrame(data, columns=[
        '지역', '사업장명', '배출구', '측정시간', '먼지 배출허용기준', '먼지 측정값',
        '황산화물 배출허용기준', '황산화물 측정값', '질소산화물 배출허용기준', '질소산화물 측정값',
        '염화수소 배출허용기준', '염화수소 측정값', '불화수소 배출허용기준', '불화수소 측정값',
        '암모니아 배출허용기준', '암모니아 측정값', '일산화탄소 배출허용기준', '일산화탄소 측정값'
    ])

    # 엑셀 파일로 저장
    df.to_excel('output.xlsx', index=False)
    print('데이터를 성공적으로 엑셀 파일로 저장했습니다.')

finally:
    driver.quit()
